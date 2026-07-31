import { type App, type TFile } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import { isElement, isNode } from '../domGuards';
import { debugLog, errorLog, logPptxAction } from '../logger';
import { createMenuItem, createMenuSection, createPopoverShell, positionPopoverBelow } from '../menuControls';
import type { InsertableShapeGeometry, ParagraphTextRange, PresentationEngine, TextBoxInsertOrigin } from '../PresentationEngine';
import {
  getImageMimeType,
  InsertTableModal,
  POWERPOINT_IMAGE_FILE_ACCEPT,
  VaultImageSuggestModal
} from '../PowerPointInsertModals';
import { normalizeImageForPowerPoint } from './heicToPng';
import {
  CHART_TEMPLATE_ENTRIES,
  INSERTABLE_CHART_TYPES,
  readRecentChartTypes,
  rememberRecentChartType,
  TEXT_BOX_INSET_EMU,
  type InsertableChartType,
  type ParagraphListStyle,
} from '../SlideInsertions';
import { cleanError } from './runtimeCompat';
import type { PresentationSession } from './session/PresentationSession';
import type { PptxCommand } from './commands/types';
import type { HistoryEntry, TextEditTarget } from './types';

const LEADING_MANUAL_BULLET = /^[\t \u00A0]*[•◦▪‣⁃][\t \u00A0]+/;

export interface InsertHost {
  readonly t: TranslateFn;
  readonly session: PresentationSession;
  readonly engine: PresentationEngine | null;
  readonly app: App;
  readonly layoutEl: HTMLElement | null;
  readonly selectedShapeIndex: number | null;
  readonly activeEditorTarget: SVGTextElement | SVGTSpanElement | null;
  currentSlide: number;
  ensureEditable(action: string): boolean;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderEditedShape(shapeIndex: number): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  syncCurrentThumbnailShape(shapeIndex: number): boolean;
  /** Drop filmstrip SVG reuse before a structural insert redraw. */
  invalidateCachedSlideRenders(indices: number | number[]): void;
  selectShape(shapeIndex: number): void;
  selectShapeForTextEditing(shapeIndex: number): void;
  startTextEditor(): void;
  createEditIconButton(container: HTMLElement, icon: string, label: string, onClick: () => void): HTMLButtonElement;
  getTextEditTarget(target: SVGTextElement | SVGTSpanElement | null): TextEditTarget | null;
  getListStyleTarget(): { shapeIndex: number; paragraphIndex: number; ranges: ParagraphTextRange[] } | null;
  finishInlineTextEditing(reason: string): Promise<void>;
  registerDocumentPointerDown(handler: (event: PointerEvent) => void, capture?: boolean): void;
  openToolbarPopover(anchor: HTMLElement, build: (popover: HTMLElement) => void): void;
  bindToolbarButton(button: HTMLElement, action: () => void): void;
  closeToolbarPopover(): void;
}

export class InsertController {
  private activeInsertMenu: HTMLElement | null = null;
  private insertTableButton: HTMLButtonElement | null = null;
  private insertChartButton: HTMLButtonElement | null = null;
  private imageFileInput: HTMLInputElement | null = null;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: InsertHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  private tb(suffix: string): string {
    return this.host.t(`powerpoint:toolbar.${suffix}`);
  }

  private getInsertEngine(
    action: string,
    failureMessage: string,
    context: Record<string, unknown> = {}
  ): PresentationEngine | null {
    const engine = this.host.engine;
    if (!engine) {
      errorLog('insert', failureMessage, {
        slide: this.host.currentSlide,
        ...context,
        reason: 'Presentation engine unavailable'
      });
      return null;
    }
    if (!this.host.ensureEditable(action)) {
      errorLog('insert', failureMessage, {
        slide: this.host.currentSlide,
        ...context,
        reason: 'Presentation is not editable'
      });
      return null;
    }
    return engine;
  }

  private async commitInsertedShape(
    historyLabel: string,
    command: PptxCommand,
    editImmediately = false
  ): Promise<number> {
    const history = await this.host.captureHistoryEntry(historyLabel);
    const shapeIndex = await this.host.session.applyCommand(command) as number;
    this.host.recordHistoryEntry(history);
    // Structural inserts update OOXML but leave the filmstrip thumbnail marked
    // rendered. Without invalidation, renderCurrentSlide clones that stale SVG
    // (source: thumbnail-cache) and selectShape cannot find the new shapeIndex.
    this.host.invalidateCachedSlideRenders(this.host.currentSlide);
    const rendered = await this.host.renderCurrentSlide(editImmediately);
    if (rendered) {
      if (editImmediately) {
        this.host.selectShapeForTextEditing(shapeIndex);
        this.host.startTextEditor();
      } else {
        this.host.selectShape(shapeIndex);
        await this.host.renderThumbnails();
      }
    }
    return shapeIndex;
  }

  registerMenus(): void {
    const closeMenus = (event: MouseEvent) => {
      const target = isNode(event.target) ? event.target : null;
      if (target && this.activeInsertMenu?.contains(target)) return;
      if (isElement(target) && target.closest('.native-powerpoint-insert-menu-anchor')) return;
      this.closeInsertMenus();
    };
    this.host.registerDocumentPointerDown(closeMenus, true);

    if (!this.host.layoutEl) return;
    const input = this.host.layoutEl.createEl('input', {
      type: 'file',
      cls: 'native-powerpoint-image-file-input'
    });
    input.accept = POWERPOINT_IMAGE_FILE_ACCEPT;
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      input.value = '';
      if (file) void this.insertImageFromLocalFile(file);
    });
    this.imageFileInput = input;
  }

  createToolbarGroup(toolbar: HTMLElement): void {
    const insertGroup = toolbar.createDiv({ cls: 'native-powerpoint-toolbar-group' });

    const imageButton = this.host.createEditIconButton(insertGroup, 'image', this.tb('insertImage'), () => {
      this.toggleInsertMenu(imageButton, [
        { label: this.tb('fromVault'), onClick: () => this.openVaultImagePicker() },
        { label: this.tb('uploadFile'), onClick: () => this.imageFileInput?.click() }
      ]);
    });

    const shapeButton = this.host.createEditIconButton(insertGroup, 'shapes', this.tb('insertShape'), () => {
      this.toggleInsertMenu(shapeButton, [
        { label: this.tb('rectangle'), onClick: () => void this.insertShape('rect') },
        { label: this.tb('ellipse'), onClick: () => void this.insertShape('ellipse') },
        { label: this.tb('roundedRectangle'), onClick: () => void this.insertShape('roundRect') },
        { label: this.tb('line'), onClick: () => void this.insertShape('line') },
        { label: this.tb('arrow'), onClick: () => void this.insertShape('rightArrow') }
      ]);
    });

    this.host.createEditIconButton(insertGroup, 'type', this.tb('insertTextBox'), () => void this.insertTextBox(true));
    const tableButton = this.host.createEditIconButton(insertGroup, 'table', this.tb('insertTable'), () =>
      this.openTableSizePicker(tableButton)
    );
    this.insertTableButton = tableButton;
    const chartButton = this.host.createEditIconButton(insertGroup, 'bar-chart-3', this.tb('insertChart'), () =>
      this.openChartTypePicker(chartButton)
    );
    this.insertChartButton = chartButton;
    this.host.createEditIconButton(insertGroup, 'list', this.tb('bulletedList'), () => void this.applyListStyle('bullet'));
    this.host.createEditIconButton(insertGroup, 'list-ordered', this.tb('numberedList'), () => void this.applyListStyle('number'));
  }

  toggleInsertMenu(
    anchor: HTMLButtonElement,
    items: { label: string; icon?: string; onClick: () => void }[]
  ): void {
    if (!anchor.dataset.menuId) {
      anchor.dataset.menuId = `insert-menu-${Math.random().toString(36).slice(2)}`;
    }

    // Only treat a repeat click as "toggle closed" when a menu is actually open
    // for this anchor. Without the explicit null check, a first click compares
    // two `undefined`s (no open menu, no anchor id yet) and wrongly closes.
    if (this.activeInsertMenu && this.activeInsertMenu.dataset.anchorId === anchor.dataset.menuId) {
      this.closeInsertMenus();
      return;
    }

    this.closeInsertMenus();
    anchor.classList.add('native-powerpoint-insert-menu-anchor');

    const menu = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-insert-menu native-powerpoint-light-surface'
    });
    menu.dataset.anchorId = anchor.dataset.menuId;
    for (const item of items) {
      createMenuItem(menu, {
        className: 'native-powerpoint-insert-menu-item',
        text: item.label,
        icon: item.icon ? { name: item.icon, className: 'native-powerpoint-insert-menu-icon' } : undefined,
        preventDefaultOnClick: true,
        stopClickPropagation: true,
        onClick: () => {
          this.closeInsertMenus();
          item.onClick();
        }
      });
    }

    positionPopoverBelow(menu, anchor);
    this.activeInsertMenu = menu;
  }

  openChartTypePicker(anchor?: HTMLButtonElement | null): void {
    const chartAnchor = anchor ?? this.insertChartButton;
    if (!chartAnchor) return;

    if (!chartAnchor.dataset.menuId) {
      chartAnchor.dataset.menuId = `insert-menu-${Math.random().toString(36).slice(2)}`;
    }
    if (this.activeInsertMenu && this.activeInsertMenu.dataset.anchorId === chartAnchor.dataset.menuId) {
      this.closeInsertMenus();
      return;
    }

    this.closeInsertMenus();
    chartAnchor.classList.add('native-powerpoint-insert-menu-anchor');
    logPptxAction('insert', 'open-chart-type-picker', {
      slide: this.host.currentSlide,
      recentCount: readRecentChartTypes().length,
    });

    const menu = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-insert-menu native-powerpoint-insert-chart-menu native-powerpoint-light-surface'
    });
    menu.dataset.anchorId = chartAnchor.dataset.menuId;

    const addChartItem = (label: string, icon: string, chartType: InsertableChartType) => {
      createMenuItem(menu, {
        className: 'native-powerpoint-insert-menu-item',
        text: label,
        icon: { name: icon, className: 'native-powerpoint-insert-menu-icon' },
        preventDefaultOnClick: true,
        stopClickPropagation: true,
        onClick: () => {
          this.closeInsertMenus();
          void this.insertChart(chartType);
        }
      });
    };

    const recent = readRecentChartTypes();
    if (recent.length > 0) {
      createMenuSection(menu, {
        className: 'native-powerpoint-insert-menu-section',
        text: this.tb('chartTypeRecent'),
      });
      for (const chartType of recent) {
        const entry = INSERTABLE_CHART_TYPES.find((candidate) => candidate.id === chartType);
        if (!entry) continue;
        addChartItem(this.tb(`chartType.${entry.labelKey}`), entry.icon, entry.id);
      }
    }

    createMenuSection(menu, {
      className: 'native-powerpoint-insert-menu-section',
      text: this.tb('chartTypeTemplates'),
    });
    for (const template of CHART_TEMPLATE_ENTRIES) {
      addChartItem(this.tb(`chartType.${template.labelKey}`), template.icon, template.chartType);
    }

    createMenuSection(menu, {
      className: 'native-powerpoint-insert-menu-section',
      text: this.tb('chartTypeAll'),
    });
    for (const entry of INSERTABLE_CHART_TYPES) {
      addChartItem(this.tb(`chartType.${entry.labelKey}`), entry.icon, entry.id);
    }

    positionPopoverBelow(menu, chartAnchor);
    this.activeInsertMenu = menu;
  }

  getInsertChartButton(): HTMLButtonElement | null {
    return this.insertChartButton;
  }

  closeInsertMenus(): void {
    this.activeInsertMenu?.remove();
    this.activeInsertMenu = null;
  }

  clickImageFileInput(): void {
    this.imageFileInput?.click();
  }

  getInsertTableButton(): HTMLButtonElement | null {
    return this.insertTableButton;
  }

  openVaultImagePicker(): void {
    new VaultImageSuggestModal(this.host.app, (file) => void this.insertImageFromVaultFile(file), this.host.t).open();
  }

  openInsertTableModal(): void {
    if (!this.host.ensureEditable('insert table')) return;
    new InsertTableModal(this.host.app, (rows, cols) => void this.insertTable(rows, cols), this.host.t).open();
  }

  // Google Slides-style size picker: a hover grid that matches the look of the
  // other toolbar popovers (color, font) instead of a separate modal dialog.
  openTableSizePicker(anchor: HTMLElement | null): void {
    if (!this.host.ensureEditable('insert table')) return;
    if (!anchor) {
      this.openInsertTableModal();
      return;
    }

    const cols = 10;
    const rows = 8;
    this.host.openToolbarPopover(anchor, (popover) => {
      popover.addClass('native-powerpoint-table-picker');

      const grid = popover.createDiv({ cls: 'native-powerpoint-table-picker-grid' });
      const insertTableLabel = this.tb('insertTable');
      const label = popover.createDiv({
        cls: 'native-powerpoint-table-picker-label',
        text: insertTableLabel
      });

      const cells: HTMLButtonElement[] = [];
      const highlight = (activeCols: number, activeRows: number): void => {
        cells.forEach((cell, index) => {
          const c = index % cols;
          const r = Math.floor(index / cols);
          cell.toggleClass('is-active', c < activeCols && r < activeRows);
        });
        label.setText(
          activeCols > 0 && activeRows > 0
            ? this.host.t('powerpoint:accessibility.tableCellSize', { columns: activeCols, rows: activeRows })
            : insertTableLabel
        );
      };

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = grid.createEl('button', {
            cls: 'native-powerpoint-table-picker-cell',
            attr: {
              'aria-label': this.host.t('powerpoint:accessibility.tableCellSize', {
                columns: c + 1,
                rows: r + 1
              })
            }
          });
          cell.addEventListener('pointerenter', () => highlight(c + 1, r + 1));
          this.host.bindToolbarButton(cell, () => {
            this.host.closeToolbarPopover();
            void this.insertTable(r + 1, c + 1);
          });
          cells.push(cell);
        }
      }

      grid.addEventListener('pointerleave', () => highlight(0, 0));
    });
  }

  async insertTextBox(editImmediately = false, origin?: TextBoxInsertOrigin): Promise<void> {
    if (!this.getInsertEngine('insert text box', 'PowerPoint text-box insertion failed')) return;

    try {
      const shapeIndex = await this.commitInsertedShape('Add text box', {
        type: 'insert-text-box',
        slideIndex: this.host.currentSlide,
        origin,
      }, editImmediately);
		debugLog('insert', 'Inserted PowerPoint text box', {
			slide: this.host.currentSlide,
			shapeIndex,
        requestedOrigin: origin ?? null,
			textInsetEmu: TEXT_BOX_INSET_EMU,
		});
    } catch (error) {
      errorLog('insert', 'PowerPoint text-box insertion failed', {
        slide: this.host.currentSlide,
        error
      });
      this.notice('powerpoint:notice.couldNotAddTextBox', { message: cleanError(error) });
    }
  }

  async insertImageFromVaultFile(file: TFile): Promise<void> {
    if (!this.getInsertEngine('insert image', 'PowerPoint vault image insertion failed', {
      file: file.path
    })) return;

    try {
      const rawBytes = new Uint8Array(await this.host.app.vault.readBinary(file));
      const normalized = await normalizeImageForPowerPoint(
        rawBytes,
        getImageMimeType(file.extension),
        file.extension,
      );
      const shapeIndex = await this.commitInsertedShape('Insert image', {
        type: 'insert-image',
        slideIndex: this.host.currentSlide,
        imageData: normalized.bytes,
        mimeType: normalized.mimeType,
      });
      debugLog('insert', 'Inserted PowerPoint image from vault', {
        slide: this.host.currentSlide,
        shapeIndex,
        file: file.path,
        bytes: normalized.bytes.byteLength,
        convertedFromHeic: normalized.convertedFromHeic,
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint vault image insertion failed', { file: file.path, error });
      this.notice('powerpoint:notice.couldNotInsertImage', { message: cleanError(error) });
    }
  }

  async insertImageFromLocalFile(file: File): Promise<void> {
    if (!this.getInsertEngine('insert image', 'PowerPoint local image insertion failed', {
      fileName: file.name,
      mimeType: file.type || null
    })) return;

    try {
      const rawBytes = new Uint8Array(await file.arrayBuffer());
      const normalized = await normalizeImageForPowerPoint(
        rawBytes,
        file.type || getImageMimeType(file.name.split('.').pop() ?? 'png'),
        file.name,
      );
      const shapeIndex = await this.commitInsertedShape('Insert image', {
        type: 'insert-image',
        slideIndex: this.host.currentSlide,
        imageData: normalized.bytes,
        mimeType: normalized.mimeType,
      });
      debugLog('insert', 'Inserted PowerPoint image from local file', {
        slide: this.host.currentSlide,
        shapeIndex,
        fileName: file.name,
        mimeType: normalized.mimeType,
        bytes: normalized.bytes.byteLength,
        convertedFromHeic: normalized.convertedFromHeic,
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint local image insertion failed', {
        fileName: file.name,
        mimeType: file.type || null,
        error
      });
      this.notice('powerpoint:notice.couldNotInsertImage', { message: cleanError(error) });
    }
  }

  async insertShape(geometry: InsertableShapeGeometry): Promise<void> {
    if (!this.getInsertEngine('insert shape', 'PowerPoint shape insertion failed', { geometry })) return;

    try {
      const shapeIndex = await this.commitInsertedShape('Insert shape', {
        type: 'insert-shape',
        slideIndex: this.host.currentSlide,
        geometry
      });
      debugLog('insert', 'Inserted PowerPoint shape', {
        slide: this.host.currentSlide,
        shapeIndex,
        geometry
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint shape insertion failed', { geometry, error });
      this.notice('powerpoint:notice.couldNotInsertShape', { message: cleanError(error) });
    }
  }

  async insertTable(rows: number, cols: number): Promise<void> {
    if (!this.getInsertEngine('insert table', 'PowerPoint table insertion failed', {
      rows,
      columns: cols
    })) return;

    try {
      const shapeIndex = await this.commitInsertedShape('Insert table', {
        type: 'insert-table',
        slideIndex: this.host.currentSlide,
        rows,
        cols
      });
      debugLog('insert', 'Inserted PowerPoint table', {
        slide: this.host.currentSlide,
        shapeIndex,
        rows,
        columns: cols
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint table insertion failed', { rows, columns: cols, error });
      this.notice('powerpoint:notice.couldNotInsertTable', { message: cleanError(error) });
    }
  }

  async insertChart(chartType: InsertableChartType = 'column'): Promise<void> {
    if (!this.getInsertEngine('insert chart', 'PowerPoint chart insertion failed', { chartType })) return;

    try {
      const shapeIndex = await this.commitInsertedShape('Insert chart', {
        type: 'insert-chart',
        slideIndex: this.host.currentSlide,
        chartType,
      });
      rememberRecentChartType(chartType);
      debugLog('insert', 'Inserted PowerPoint chart', {
        slide: this.host.currentSlide,
        shapeIndex,
        chartType,
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint chart insertion failed', {
        slide: this.host.currentSlide,
        chartType,
        error,
      });
      this.notice('powerpoint:notice.couldNotInsertChart', { message: cleanError(error) });
    }
  }

  async applyListStyle(style: ParagraphListStyle): Promise<void> {
    const engine = this.getInsertEngine('format text', 'PowerPoint list-style update failed', { style });
    if (!engine) return;

    // Resolve before awaiting an inline commit. That commit intentionally clears
    // the live SVG target, which otherwise makes a paragraph-2 selection fall
    // back to paragraph 0 of the selected text box.
    const listTarget = this.host.getListStyleTarget();
    const shapeIndex = listTarget?.shapeIndex ?? this.host.selectedShapeIndex;
    if (shapeIndex === null) {
      errorLog('insert', 'PowerPoint list-style update failed', {
        slide: this.host.currentSlide,
        style,
        reason: 'No text box selected'
      });
      this.notice('powerpoint:notice.selectTextBoxFirst');
      return;
    }

    const paragraphIndex = listTarget?.paragraphIndex ?? 0;
    const ranges = listTarget?.ranges ?? [];
    const selectedParagraphIndexes = [...new Set(ranges.map((candidate) => candidate.paragraphIndex))];
    const targetParagraphIndexes = selectedParagraphIndexes.length > 0
      ? selectedParagraphIndexes
      : [paragraphIndex];
    const rawCurrentStyles = targetParagraphIndexes.map((index) => (
      engine.getParagraphListStyle(this.host.currentSlide, shapeIndex, index)
    ));
    const hasLeadingManualBullets = targetParagraphIndexes.map((index) => {
      return LEADING_MANUAL_BULLET.test(
        engine.getParagraphRunText(this.host.currentSlide, shapeIndex, index) ?? '',
      );
    });
    // Imported bullets can be literal text with an explicit `buNone`. Treat
    // those visible markers as bullets for toggle resolution, so one click
    // removes them instead of silently converting them into native bullets.
    const currentStyles = rawCurrentStyles.map((candidate, index) => (
      candidate === 'none' && hasLeadingManualBullets[index] ? 'bullet' : candidate
    ));
    const currentStyle = currentStyles[0] ?? null;
    const stripLeadingManualBullet = hasLeadingManualBullets.some(Boolean);
    const resolvedStyle = currentStyles.length > 0 && currentStyles.every((candidate) => candidate === style)
      ? 'none'
      : style;
    // A list marker belongs to a paragraph, not its selected text fragment.
    // Removing a list from a cross-paragraph drag must therefore expand every
    // target back to its full run text; otherwise the first partial range splits
    // and leaves its prefix bullet behind.
    const removeWholeParagraphLists = resolvedStyle === 'none' && ranges.length > 0;
    const commandRanges = removeWholeParagraphLists
      ? targetParagraphIndexes.map((index) => ({
        paragraphIndex: index,
        start: 0,
        end: (engine.getParagraphRunText(this.host.currentSlide, shapeIndex, index) ?? '').length,
      })).filter((candidate) => candidate.end > candidate.start)
      : ranges;
    const range = commandRanges.length === 1 ? commandRanges[0] ?? null : null;
    logPptxAction('text-format', 'apply-list-style', {
      slide: this.host.currentSlide,
      shapeIndex,
      paragraphIndex,
      style: resolvedStyle,
      requestedStyle: style,
      currentStyle,
      currentStyles,
      rawCurrentStyles,
      hasLeadingManualBullets,
      toggled: resolvedStyle === 'none',
      stripLeadingManualBullet,
      strategy: removeWholeParagraphLists
        ? 'whole-paragraph-ranges'
        : commandRanges.length > 1 ? 'selected-text-ranges' : range ? 'selected-text-range' : 'whole-paragraph',
      range: range ? { start: range.start, end: range.end } : null,
      selectedRangeCount: ranges.length,
      appliedRangeCount: commandRanges.length,
      selectedParagraphIndexes,
    });
    try {
      await this.host.finishInlineTextEditing('before-list-formatting');
      const history = await this.host.captureHistoryEntry(
        resolvedStyle === 'bullet' ? 'Bulleted list' : resolvedStyle === 'number' ? 'Numbered list' : 'Remove list'
      );
      await this.host.session.applyCommand({
        slideIndex: this.host.currentSlide,
        shapeIndex,
        style: resolvedStyle,
        stripLeadingManualBullet,
        ...(commandRanges.length > 1
          ? { type: 'apply-list-style-ranges' as const, ranges: commandRanges }
          : range
            ? { type: 'apply-list-style-range' as const, range }
            : { type: 'apply-list-style' as const, paragraphIndex }),
      });
      const appliedStyles = targetParagraphIndexes.map((index) => (
        engine.getParagraphListStyle(this.host.currentSlide, shapeIndex, index)
      ));
      debugLog('text-format', 'Read back PowerPoint list style mutation', {
        slide: this.host.currentSlide,
        shapeIndex,
        targetParagraphIndexes,
        expectedStyle: resolvedStyle,
        appliedStyles,
        matchesExpected: appliedStyles.every((candidate) => candidate === resolvedStyle),
      });
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderEditedShape(shapeIndex);
      // A list style changes one shape. Rebuilding the poster filmstrip renders
      // the entire slide synchronously; clone the canonical edited group instead.
      const thumbnailSynced = rendered && this.host.syncCurrentThumbnailShape(shapeIndex);
      debugLog('text-format', 'Applied PowerPoint list style', {
        slide: this.host.currentSlide,
        shapeIndex,
        paragraphIndex,
        style: resolvedStyle,
        requestedStyle: style,
        currentStyle,
        currentStyles,
        rawCurrentStyles,
        hasLeadingManualBullets,
        toggled: resolvedStyle === 'none',
        stripLeadingManualBullet,
        strategy: removeWholeParagraphLists
          ? 'whole-paragraph-ranges'
          : commandRanges.length > 1 ? 'selected-text-ranges' : range ? 'selected-text-range' : 'whole-paragraph',
        range: range ? { start: range.start, end: range.end } : null,
        selectedRangeCount: ranges.length,
        appliedRangeCount: commandRanges.length,
        selectedParagraphIndexes,
        rendered,
        thumbnailSynced,
      });
    } catch (error) {
      errorLog('insert', 'PowerPoint list-style update failed', {
        slide: this.host.currentSlide,
        shapeIndex,
        paragraphIndex,
        style,
        error
      });
      this.notice('powerpoint:notice.couldNotUpdateListStyle', { message: cleanError(error) });
    }
  }
}
