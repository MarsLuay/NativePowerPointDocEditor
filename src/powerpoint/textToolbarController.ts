import type { TranslateFn } from '../i18n/translate';

import { isSVGTSpanElement } from '../domGuards';
import {
	PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS,
} from '../editorChromeRegions';
import { bindPopoverDismissHandlers, createMenuItem, createPopoverShell, createToolbarIconButton, positionPopoverBelow } from '../menuControls';
import type {
  ParagraphAlignment,
  ParagraphTextRange,
  RunStyleChange,
  RunStyleInfo,
  RunTarget
} from '../PresentationEngine';
import type { PresentationEngine } from '../PresentationEngine';
import { debugLog, logPptxAction } from '../logger';
import {
  TEXT_TOOLBAR_FONTS,
  TEXT_TOOLBAR_MAX_FONT_SIZE,
  TEXT_TOOLBAR_MIN_FONT_SIZE,
  TEXT_TOOLBAR_SWATCHES
} from './constants';
import { parsePrimaryFontFamily } from './textUtils';
import type {
  ShapeTextEditTarget,
  TextStyleContext,
  TextToolbarControls,
  ToolbarFormattingSnapshot
} from './types';

export interface TextToolbarHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  readonly svgEl: SVGSVGElement | null;
  readonly canvasPane: HTMLElement | null;
  currentSlide: number;
  readonly activeEditor: HTMLTextAreaElement | null;
  readonly activeEditorTarget: SVGTextElement | SVGTSpanElement | null;
  readonly activeTextStyleTarget: ShapeTextEditTarget | null;
  currentRunStyle: RunStyleInfo | null;
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  getTextStyleContext(): TextStyleContext | null;
  getElementBox(element: Element): { left: number; top: number; width: number; height: number } | null;
  getSelectedBox(): { left: number; top: number; width: number; height: number } | null;
  getStoredInlineSelectionRanges(shapeIndex: number): ParagraphTextRange[] | null;
  getSelectedRangeFontSizePt(shapeIndex: number, ranges: ParagraphTextRange[]): number | null;
  applyRunStyle(change: RunStyleChange): void | Promise<void>;
  applyAlignment(align: ParagraphAlignment): void;
  flushActiveEditor(): void;
}

interface FontSizeOverride {
  contextKey: string;
  fontSizePt: number;
}

/**
 * Owns the floating text-format toolbar, top-bar font-family control, and the
 * shared toolbar popover used for font/color/table pickers. Extracted from
 * `NativePowerPointView`; it borrows text-editing state through {@link TextToolbarHost}.
 */
export class TextToolbarController {
  private textToolbarEl: HTMLElement | null = null;
  private textToolbarControls: TextToolbarControls | null = null;
  private textToolbarShapeIndex: number | null = null;
  private topFontButton: HTMLButtonElement | null = null;
  private topFontLabel: HTMLElement | null = null;
  private textColorValue = '000000';
  private textHighlightValue = 'FFFF00';
  private activeToolbarPopover: HTMLElement | null = null;
  private toolbarPopoverCleanup: (() => void) | null = null;
  private toolbarFormattingSnapshot: ToolbarFormattingSnapshot | null = null;
  private fontSizeOverride: FontSizeOverride | null = null;
  /** Latest font size requested while a prior apply is still in flight. */
  private pendingFontSizePt: number | null = null;
  private fontSizeApplyRunning = false;

  constructor(private readonly host: TextToolbarHost) {}

  getFormattingSnapshot(): ToolbarFormattingSnapshot | null {
    return this.toolbarFormattingSnapshot;
  }

  clearFormattingSnapshot(): void {
    this.toolbarFormattingSnapshot = null;
  }

  hasFormattingSnapshot(): boolean {
    return this.toolbarFormattingSnapshot !== null;
  }

  hasActivePopover(): boolean {
    return this.activeToolbarPopover !== null;
  }

  getToolbarShapeIndex(): number | null {
    return this.textToolbarShapeIndex;
  }

  /** Keep shape/selection context across toolbar popovers even if the inline editor closes. */
  preserveFormattingContext(): void {
    const snapshot = this.captureToolbarFormattingSnapshot();
    if (snapshot) {
      this.toolbarFormattingSnapshot = snapshot;
      debugLog('text-format', 'Preserved toolbar formatting context', {
        shapeIndex: snapshot.shapeIndex,
        rangeCount: snapshot.ranges?.length ?? 0,
        run: snapshot.run,
      });
    }
  }

  reset(): void {
    this.textToolbarEl = null;
    this.textToolbarControls = null;
    this.textToolbarShapeIndex = null;
    this.topFontButton = null;
    this.topFontLabel = null;
    this.fontSizeOverride = null;
    this.pendingFontSizePt = null;
    this.fontSizeApplyRunning = false;
  }

  wireFontFamilyButton(button: HTMLButtonElement, label: HTMLElement): void {
    this.topFontButton = button;
    this.topFontLabel = label;
    this.bindToolbarButton(button, () => this.openFontMenu(button));
    this.setTopFontControl('Font', false);
  }

  updateTextToolbar(): void {
    const context = this.host.getTextStyleContext();
    if (!context) {
      this.hideTextToolbar();
      return;
    }

    const controls = this.ensureTextToolbar();
    if (!controls || !this.textToolbarEl) return;

    // Position the toolbar only when it first appears for a shape. Subsequent
    // updates (e.g. flushing the inline editor after clicking the font-size box)
    // change the anchor from the caret line to the whole-shape box, which would
    // make the toolbar jump; keeping the spawn position avoids that.
    const wasVisible = this.textToolbarEl.hasClass('is-visible');
    const shapeChanged = this.textToolbarShapeIndex !== context.shapeIndex;
    this.textToolbarEl.addClass('is-visible');
    if (!wasVisible || shapeChanged) {
      this.positionTextToolbar(context.anchor);
      this.textToolbarShapeIndex = context.shapeIndex;
    }
    this.reflectTextToolbarState(context);
  }

  bindToolbarButton(button: HTMLElement, action: () => void): void {
    button.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();
      action();
    });
  }

  openToolbarPopover(anchor: HTMLElement, build: (popover: HTMLElement) => void): void {
    this.closeToolbarPopover();

    const popover = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-toolbar-popover native-powerpoint-light-surface',
      stopPointerDown: true
    });
    build(popover);

    positionPopoverBelow(popover, anchor);

    this.activeToolbarPopover = popover;
    this.toolbarPopoverCleanup = bindPopoverDismissHandlers({
      popover,
      anchor,
      onDismiss: () => this.closeToolbarPopover(),
      closeOnEscape: false
    });
  }

  closeToolbarPopover(): void {
    this.toolbarPopoverCleanup?.();
    this.toolbarPopoverCleanup = null;
    this.activeToolbarPopover?.remove();
    this.activeToolbarPopover = null;
  }

  private hideTextToolbar(): void {
    this.textToolbarEl?.removeClass('is-visible');
    this.textToolbarShapeIndex = null;
    this.closeToolbarPopover();
    this.host.currentRunStyle = null;
    this.fontSizeOverride = null;
    this.setTopFontControl('Font', false);
  }

  private setTopFontControl(label: string, enabled: boolean): void {
    this.topFontLabel?.setText(label);
    const button = this.topFontButton;
    if (!button) return;
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
  }

  private positionTextToolbar(anchor: { left: number; top: number; width: number; height: number }): void {
    const toolbar = this.textToolbarEl;
    if (!toolbar || !this.host.canvasPane) return;

    const toolbarHeight = toolbar.offsetHeight || 40;
    const gap = 8;
    let top = anchor.top - toolbarHeight - gap;
    if (top < this.host.canvasPane.scrollTop + 4) {
      top = anchor.top + anchor.height + gap;
    }

    const maxLeft = Math.max(0, this.host.canvasPane.scrollWidth - (toolbar.offsetWidth || 0) - 4);
    const left = Math.min(Math.max(anchor.left, 4), maxLeft);
    toolbar.setCssProps({ left: `${left}px`, top: `${Math.max(0, top)}px` });
  }

  private reflectTextToolbarState(context: TextStyleContext): void {
    const controls = this.textToolbarControls;
    if (!controls || !this.host.engine) return;

    const runTarget = context.run ?? this.getFirstRunTarget(context.shapeIndex);
    const style = runTarget
      ? this.host.engine.getRunStyle(this.host.currentSlide, context.shapeIndex, runTarget.paragraphIndex, runTarget.runIndex)
      : null;
    const selectedRanges = this.getActiveInlineSelectionRanges(context.shapeIndex);
    const hasSelectedText = selectedRanges?.some((range) => range.start !== range.end) ?? false;
    const selectedRangeFontSizePt = hasSelectedText
      ? this.host.getSelectedRangeFontSizePt(context.shapeIndex, selectedRanges ?? [])
      : null;
    let reflectedStyle = selectedRanges?.length && style
      ? {
          ...style,
          bold: this.host.engine.areRangesStyled(this.host.currentSlide, context.shapeIndex, selectedRanges, 'bold'),
          italic: this.host.engine.areRangesStyled(this.host.currentSlide, context.shapeIndex, selectedRanges, 'italic'),
          underline: this.host.engine.areRangesStyled(this.host.currentSlide, context.shapeIndex, selectedRanges, 'underline')
        }
      : style;
    const optimisticFontSize = this.getFontSizeOverride(context, selectedRanges);
    if (optimisticFontSize !== null && reflectedStyle) {
      reflectedStyle = { ...reflectedStyle, fontSizePt: optimisticFontSize };
    } else if (hasSelectedText && selectedRangeFontSizePt !== null && reflectedStyle) {
      reflectedStyle = { ...reflectedStyle, fontSizePt: selectedRangeFontSizePt };
    }
    this.host.currentRunStyle = reflectedStyle;

    controls.bold.toggleClass('is-active', Boolean(reflectedStyle?.bold));
    controls.italic.toggleClass('is-active', Boolean(reflectedStyle?.italic));
    controls.underline.toggleClass('is-active', Boolean(reflectedStyle?.underline));
    this.setTopFontControl(reflectedStyle?.fontFamily ?? this.getEffectiveFontFamily(context) ?? 'Font', true);

    if (activeDocument.activeElement !== controls.fontSizeInput) {
      const sizePt = optimisticFontSize
        ?? (hasSelectedText
          ? selectedRangeFontSizePt
          : (reflectedStyle?.fontSizePt ?? this.getEffectiveFontSizePt(context)));
      const nextValue = sizePt ? String(Math.round(sizePt)) : '';
      if (controls.fontSizeInput.value !== nextValue) {
        debugLog('text-format', 'Reflected font size in toolbar', {
          shapeIndex: context.shapeIndex,
          selectionRangeCount: selectedRanges?.length ?? 0,
          source: optimisticFontSize !== null
            ? 'selection-override'
            : hasSelectedText
              ? 'selected-ranges'
              : 'rendered-style',
          previous: controls.fontSizeInput.value,
          next: nextValue,
          runStyleFontSizePt: style?.fontSizePt ?? null,
          selectedRangeFontSizePt,
          optimisticFontSizePt: optimisticFontSize
        });
        controls.fontSizeInput.value = nextValue;
      }
    }

    if (reflectedStyle?.color) {
      this.textColorValue = reflectedStyle.color;
    }
    if (reflectedStyle?.highlight) {
      this.textHighlightValue = reflectedStyle.highlight;
    }
    controls.textColorBar.style.setProperty('--np-swatch-color', `#${reflectedStyle?.color ?? this.textColorValue}`);
    controls.highlightBar.style.setProperty('--np-swatch-color', reflectedStyle?.highlight ? `#${reflectedStyle.highlight}` : 'transparent');

    const alignment = reflectedStyle?.alignment ?? 'l';
    for (const align of ['l', 'ctr', 'r', 'just'] as ParagraphAlignment[]) {
      controls.alignButtons[align].toggleClass('is-active', alignment === align);
    }
  }

  private getFirstRunTarget(shapeIndex: number): RunTarget | null {
    const shape = this.host.svgEl?.querySelector(`g[data-ooxml-shape-idx="${shapeIndex}"]`);
    const run = shape?.querySelector('tspan[data-ooxml-run-idx]') ?? null;
    if (!run) return null;

    const paragraph = run.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paragraph?.getAttribute('data-ooxml-para-idx'));
    const runIndex = Number(run.getAttribute('data-ooxml-run-idx'));
    if (!Number.isFinite(paragraphIndex) || !Number.isFinite(runIndex)) return null;
    return { paragraphIndex, runIndex };
  }

  private getActiveInlineSelectionRanges(shapeIndex: number): ParagraphTextRange[] | null {
    return this.host.getStoredInlineSelectionRanges(shapeIndex)
      ?? (this.toolbarFormattingSnapshot?.shapeIndex === shapeIndex ? this.toolbarFormattingSnapshot.ranges : null);
  }

  private getSvgEmuPerUnit(): number | null {
    const scale = Number(this.host.svgEl?.getAttribute('data-ooxml-scale'));
    if (Number.isFinite(scale) && scale > 0) return scale;

    const cx = Number(this.host.svgEl?.getAttribute('data-ooxml-slide-cx'));
    const width = this.host.svgEl ? Number.parseFloat(this.host.svgEl.getAttribute('width') || '') : Number.NaN;
    if (Number.isFinite(cx) && Number.isFinite(width) && width > 0) return cx / width;

    return null;
  }

  private getRelevantTextRuns(context: TextStyleContext): SVGTSpanElement[] {
    if (!this.host.svgEl) return [];
    const shape = this.host.svgEl.querySelector(`g[data-ooxml-shape-idx="${context.shapeIndex}"]`);
    if (!shape) return [];

    const allRuns = Array.from(shape.querySelectorAll('tspan[data-ooxml-run-idx]')).filter(isSVGTSpanElement);
    const targetRun = context.run;
    if (!targetRun) return allRuns;

    return allRuns.filter((run) => {
      const para = run.closest('tspan[data-ooxml-para-idx]');
      return Number(run.getAttribute('data-ooxml-run-idx')) === targetRun.runIndex
        && Number(para?.getAttribute('data-ooxml-para-idx')) === targetRun.paragraphIndex;
    });
  }

  private getEffectiveFontSizePt(context: TextStyleContext): number | null {
    const emuPerUnit = this.getSvgEmuPerUnit();
    if (!emuPerUnit) return null;

    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    const EMU_PER_POINT = 12700;
    let detected: number | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const userUnits = Number.parseFloat(window.getComputedStyle(run).fontSize);
      if (!Number.isFinite(userUnits) || userUnits <= 0) continue;
      const rounded = Math.round((userUnits * emuPerUnit) / EMU_PER_POINT);
      if (detected === null) {
        detected = rounded;
      } else if (detected !== rounded) {
        return null;
      }
    }
    return detected;
  }

  private getEffectiveFontFamily(context: TextStyleContext): string | null {
    const runs = this.getRelevantTextRuns(context);
    if (runs.length === 0) return null;

    let detected: string | null = null;
    for (const run of runs) {
      if ((run.textContent || '').length === 0) continue;
      const family = parsePrimaryFontFamily(window.getComputedStyle(run).fontFamily);
      if (!family) continue;
      if (detected === null) {
        detected = family;
      } else if (detected !== family) {
        return null;
      }
    }
    return detected;
  }

  private ensureTextToolbar(): TextToolbarControls | null {
    if (this.textToolbarControls && this.textToolbarEl?.isConnected) {
      return this.textToolbarControls;
    }
    if (!this.host.canvasPane) return null;

    this.textToolbarEl?.remove();
    const toolbar = this.host.canvasPane.createDiv({ cls: PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS });
    toolbar.setAttribute('role', 'toolbar');
    toolbar.setAttribute('aria-label', this.host.t('powerpoint:accessibility.textFormatting'));
    toolbar.addEventListener('pointerdown', (event) => event.stopPropagation());

    const styleGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const bold = this.createTextToolbarButton(styleGroup, 'bold', 'Bold', () => this.toggleRunFlag('bold'));
    const italic = this.createTextToolbarButton(styleGroup, 'italic', 'Italic', () => this.toggleRunFlag('italic'));
    const underline = this.createTextToolbarButton(styleGroup, 'underline', 'Underline', () => this.toggleRunFlag('underline'));

    const sizeGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    this.createTextToolbarButton(sizeGroup, 'minus', 'Decrease font size', () => {
      void this.stepFontSize(-1);
    });
    const fontSizeInput = sizeGroup.createEl('input', {
      cls: 'native-powerpoint-text-toolbar-size',
      type: 'number',
      attr: {
        'aria-label': 'Font size',
        min: String(TEXT_TOOLBAR_MIN_FONT_SIZE),
        max: String(TEXT_TOOLBAR_MAX_FONT_SIZE)
      }
    });
    fontSizeInput.addEventListener('pointerdown', () => this.flushActiveEditorForToolbarInput(), true);
    fontSizeInput.addEventListener('focus', () => this.flushActiveEditorForToolbarInput());
    fontSizeInput.addEventListener('change', () => this.commitFontSizeInput());
    fontSizeInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.commitFontSizeInput();
      }
    });
    this.createTextToolbarButton(sizeGroup, 'plus', 'Increase font size', () => {
      void this.stepFontSize(1);
    });

    const colorGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const textColorButton = this.createTextToolbarSwatchButton(colorGroup, 'baseline', 'Text color');
    const textColorBar = textColorButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(textColorButton, () => {
      this.preserveFormattingContext();
      this.openColorPopover(textColorButton, this.textColorValue, false, (color) => {
        debugLog('text-format', 'setTextColor', { color });
        void this.host.applyRunStyle({ color });
      });
    });

    const highlightButton = this.createTextToolbarSwatchButton(colorGroup, 'highlighter', 'Highlight color');
    const highlightBar = highlightButton.createDiv({ cls: 'native-powerpoint-text-toolbar-swatch-bar' });
    this.bindToolbarButton(highlightButton, () => {
      this.preserveFormattingContext();
      this.openColorPopover(highlightButton, this.textHighlightValue, true, (color) => {
        debugLog('text-format', 'setHighlight', { color });
        void this.host.applyRunStyle({ highlight: color });
      });
    });

    const alignGroup = toolbar.createDiv({ cls: 'native-powerpoint-text-toolbar-group' });
    const alignButtons: Record<ParagraphAlignment, HTMLButtonElement> = {
      l: this.createTextToolbarButton(alignGroup, 'align-left', 'Align left', () => this.applyAlignment('l')),
      ctr: this.createTextToolbarButton(alignGroup, 'align-center', 'Align center', () => this.applyAlignment('ctr')),
      r: this.createTextToolbarButton(alignGroup, 'align-right', 'Align right', () => this.applyAlignment('r')),
      just: this.createTextToolbarButton(alignGroup, 'align-justify', 'Justify', () => this.applyAlignment('just'))
    };

    this.textToolbarEl = toolbar;
    this.textToolbarControls = {
      bold,
      italic,
      underline,
      fontSizeInput,
      textColorBar,
      highlightBar,
      alignButtons
    };
    return this.textToolbarControls;
  }

  private createTextToolbarButton(
    container: HTMLElement,
    icon: string,
    label: string,
    action: () => void
  ): HTMLButtonElement {
    const button = createToolbarIconButton(container, {
      className: ['native-powerpoint-toolbar-btn', 'native-powerpoint-text-toolbar-btn'],
      icon,
      label
    });
    this.bindToolbarButton(button, action);
    return button;
  }

  private createTextToolbarSwatchButton(container: HTMLElement, icon: string, label: string): HTMLButtonElement {
    return createToolbarIconButton(container, {
      className: ['native-powerpoint-toolbar-btn', 'native-powerpoint-text-toolbar-btn', 'native-powerpoint-text-toolbar-swatch'],
      icon,
      label,
      iconClassName: 'native-powerpoint-text-toolbar-swatch-icon'
    });
  }

  private toggleRunFlag(flag: 'bold' | 'italic' | 'underline'): void {
    const context = this.host.getTextStyleContext();
    if (context && this.host.engine) {
      const selectedRanges = this.getActiveInlineSelectionRanges(context.shapeIndex);
      if (selectedRanges?.length) {
        const next = !this.host.engine.areRangesStyled(this.host.currentSlide, context.shapeIndex, selectedRanges, flag);
        debugLog('text-format', 'toggleRunFlag', { flag, path: 'inline-ranges', next, shapeIndex: context.shapeIndex });
        void this.host.applyRunStyle({ [flag]: next });
        return;
      }
    }

    const editor = this.host.activeEditor;
    const target = this.host.activeTextStyleTarget;
    if (editor && target && this.host.engine) {
      const start = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      const end = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      if (start < end) {
        const next = !this.host.engine.isRangeStyled(
          this.host.currentSlide,
          target.shapeIndex,
          target.paragraphIndex,
          start,
          end,
          flag
        );
        debugLog('text-format', 'toggleRunFlag', {
          flag,
          path: 'editor-range',
          next,
          shapeIndex: target.shapeIndex,
          paragraphIndex: target.paragraphIndex,
          start,
          end
        });
        void this.host.applyRunStyle({ [flag]: next });
        return;
      }
    }

    const current = this.host.currentRunStyle?.[flag] ?? false;
    debugLog('text-format', 'toggleRunFlag', { flag, path: 'caret-or-shape', next: !current });
    void this.host.applyRunStyle({ [flag]: !current });
  }

  private applyAlignment(align: ParagraphAlignment): void {
    debugLog('text-format', 'setAlignment', { align });
    this.host.applyAlignment(align);
  }

  private stepFontSize(delta: number): Promise<void> {
    const input = this.textToolbarControls?.fontSizeInput;
    const inputValue = Number(input?.value);
    const context = this.host.getTextStyleContext();
    const selectedRanges = context ? this.getActiveInlineSelectionRanges(context.shapeIndex) : null;
    const optimisticFontSize = context ? this.getFontSizeOverride(context, selectedRanges) : null;
    const current = optimisticFontSize
      ?? (Number.isFinite(inputValue) && inputValue > 0
        ? inputValue
        : (this.host.currentRunStyle?.fontSizePt ?? 18));
    const next = Math.min(TEXT_TOOLBAR_MAX_FONT_SIZE, Math.max(TEXT_TOOLBAR_MIN_FONT_SIZE, Math.round(current) + delta));
    this.reflectOptimisticFontSize(next);
    debugLog('text-format', 'stepFontSize', { delta, current, next, inputValue, optimisticFontSize });
    if (next === current) return Promise.resolve();
    // Optimistic UI updates immediately; coalesce rapid −/+ into one mutation
    // flush so queued full-path applies do not stack per click.
    return this.enqueueFontSizeApply(next);
  }

  private commitFontSizeInput(): void {
    const input = this.textToolbarControls?.fontSizeInput;
    if (!input) return;

    const value = Number(input.value);
    if (!Number.isFinite(value)) return;

    const clamped = Math.min(
      TEXT_TOOLBAR_MAX_FONT_SIZE,
      Math.max(TEXT_TOOLBAR_MIN_FONT_SIZE, Math.round(value))
    );
    this.reflectOptimisticFontSize(clamped);
    debugLog('text-format', 'commitFontSizeInput', { value, clamped });
    void this.enqueueFontSizeApply(clamped);
  }

  /**
   * Apply the latest requested font size. While an apply is in flight, further
   * steps only update `pendingFontSizePt` and trigger one trailing flush —
   * so five rapid − clicks become at most two mutations, not five.
   */
  private async enqueueFontSizeApply(fontSizePt: number): Promise<void> {
    this.pendingFontSizePt = fontSizePt;
    if (this.fontSizeApplyRunning) {
      debugLog('text-format', 'fontSize apply coalesced', { pendingFontSizePt: fontSizePt });
      return;
    }
    this.fontSizeApplyRunning = true;
    try {
      while (this.pendingFontSizePt !== null) {
        const size = this.pendingFontSizePt;
        this.pendingFontSizePt = null;
        debugLog('text-format', 'fontSize apply flush', { fontSizePt: size });
        await Promise.resolve(this.host.applyRunStyle({ fontSizePt: size }));
      }
    } finally {
      this.fontSizeApplyRunning = false;
      if (this.pendingFontSizePt !== null) {
        void this.enqueueFontSizeApply(this.pendingFontSizePt);
      }
    }
  }

  private reflectOptimisticFontSize(fontSizePt: number): void {
    const input = this.textToolbarControls?.fontSizeInput;
    if (input) input.value = String(fontSizePt);
    const context = this.host.getTextStyleContext();
    if (context) {
      this.fontSizeOverride = {
        contextKey: this.getFontSizeContextKey(context, this.getActiveInlineSelectionRanges(context.shapeIndex)),
        fontSizePt
      };
    }
    if (this.host.currentRunStyle) {
      this.host.currentRunStyle = { ...this.host.currentRunStyle, fontSizePt };
    }
  }

  private getFontSizeOverride(
    context: TextStyleContext,
    selectedRanges: ParagraphTextRange[] | null
  ): number | null {
    const override = this.fontSizeOverride;
    if (!override) return null;
    if (override.contextKey === this.getFontSizeContextKey(context, selectedRanges)) {
      return override.fontSizePt;
    }
    this.fontSizeOverride = null;
    return null;
  }

  private getFontSizeContextKey(context: TextStyleContext, selectedRanges: ParagraphTextRange[] | null): string {
    const runKey = context.run ? `${context.run.paragraphIndex}:${context.run.runIndex}` : 'shape';
    const rangeKey = selectedRanges?.map((range) => `${range.paragraphIndex}:${range.start}-${range.end}`).join(',') ?? 'none';
    return `${context.shapeIndex}|${runKey}|${rangeKey}`;
  }

  private captureToolbarFormattingSnapshot(): ToolbarFormattingSnapshot | null {
    const context = this.host.getTextStyleContext();
    if (!context) return null;

    let ranges = this.host.getStoredInlineSelectionRanges(context.shapeIndex);
    if (!ranges && this.toolbarFormattingSnapshot?.shapeIndex === context.shapeIndex) {
      ranges = this.toolbarFormattingSnapshot.ranges;
    }
    const editor = this.host.activeEditor;
    const target = this.host.activeTextStyleTarget;
    if (!ranges && editor && target && target.shapeIndex === context.shapeIndex) {
      const start = Math.min(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      const end = Math.max(editor.selectionStart ?? 0, editor.selectionEnd ?? 0);
      ranges = [{ paragraphIndex: target.paragraphIndex, start, end }];
    }

    return { ...context, ranges };
  }

  flushActiveEditorForToolbarInput(): void {
    const snapshot = this.captureToolbarFormattingSnapshot();
    this.host.flushActiveEditor();
    this.toolbarFormattingSnapshot = snapshot;
  }

  private openFontMenu(anchor: HTMLElement): void {
    this.preserveFormattingContext();
    const fonts = [...TEXT_TOOLBAR_FONTS];
    const context = this.host.getTextStyleContext();
    const current = this.host.currentRunStyle?.fontFamily
      ?? (context ? this.getEffectiveFontFamily(context) : null);
    if (current && !fonts.includes(current)) {
      fonts.unshift(current);
    }

    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-font-menu');
      for (const font of fonts) {
        const item = createMenuItem(popover, {
          className: 'native-powerpoint-color-popover-item native-powerpoint-font-menu-item',
          text: font
        });
        item.style.setProperty('--np-font-family', font);
        if (current === font) {
          item.addClass('is-active');
        }
        this.bindToolbarButton(item, () => {
          this.closeToolbarPopover();
          const target = this.host.getTextStyleContext();
          logPptxAction('text-format', 'change-font', {
            font,
            slide: this.host.currentSlide,
            shapeIndex: target?.shapeIndex ?? null,
          });
          debugLog('text-format', 'setFontFamily', { font });
          void this.host.applyRunStyle({ fontFamily: font });
        });
      }
    });
  }

  private openColorPopover(
    anchor: HTMLElement,
    currentColor: string,
    allowNone: boolean,
    onPick: (color: string | null) => void
  ): void {
    this.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-color-popover');

      if (allowNone) {
        const noneButton = createMenuItem(popover, {
          className: 'native-powerpoint-color-popover-none',
          text: this.host.t('powerpoint:color.noColor')
        });
        this.bindToolbarButton(noneButton, () => {
          this.closeToolbarPopover();
          onPick(null);
        });
      }

      const grid = popover.createDiv({ cls: 'native-powerpoint-color-popover-grid' });
      for (const swatch of TEXT_TOOLBAR_SWATCHES) {
        const cell = grid.createEl('button', {
          cls: 'native-powerpoint-color-popover-swatch',
          attr: { 'aria-label': `#${swatch}` }
        });
        cell.style.setProperty('--np-swatch-color', `#${swatch}`);
        const fill = cell.createSpan({ cls: 'native-powerpoint-color-popover-swatch-fill' });
        fill.style.setProperty('--np-swatch-color', `#${swatch}`);
        if (swatch.toUpperCase() === currentColor.toUpperCase()) {
          cell.addClass('is-active');
        }
        this.bindToolbarButton(cell, () => {
          this.closeToolbarPopover();
          onPick(swatch);
        });
      }

      const customRow = popover.createDiv({ cls: 'native-powerpoint-color-popover-custom' });
      customRow.createSpan({ text: this.host.t('powerpoint:color.custom') });
      const customInput = customRow.createEl('input', {
        type: 'color',
        attr: {
          'aria-label': this.host.t('powerpoint:accessibility.customColor'),
          value: `#${currentColor}`
        }
      });
      customInput.value = `#${currentColor}`;
      customInput.addEventListener('pointerdown', () => this.flushActiveEditorForToolbarInput(), true);
      customInput.addEventListener('focus', () => this.flushActiveEditorForToolbarInput());
      customInput.addEventListener('change', () => {
        const picked = customInput.value.replace(/^#/, '').toUpperCase();
        this.closeToolbarPopover();
        onPick(picked);
      });
    });
  }
}
