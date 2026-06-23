import { Notice, setIcon } from 'obsidian';

import type { PresentationEngine } from '../PresentationEngine';
import type { ShapeTransform } from 'pptx-svg';
import { isNode, isSVGTextElement } from '../domGuards';
import { debugLog, errorLog } from '../logger';
import { cleanError } from './runtimeCompat';
import { normalizeSearchText } from './textUtils';
import { getShapeIndex } from './svgUtils';
import type { HistoryEntry, PowerPointFindMatch, SvgInlineSelectionBox } from './types';

/**
 * The slice of `NativePowerPointView` that the find/replace subsystem reaches
 * back into. Keeping this surface explicit is the regression boundary for the
 * extraction: as long as these members behave as they did when the methods
 * lived on the view, find/replace behavior is unchanged.
 *
 * The view supplies this via an adapter object (not by implementing the
 * interface directly) so its own members can stay `private`.
 */
export interface FindReplaceHost {
  readonly engine: PresentationEngine | null;
  readonly isLoading: boolean;
  currentSlide: number;
  selectedShapeIndex: number | null;
  selectedTransform: ShapeTransform | null;
  readonly activeEditor: HTMLTextAreaElement | null;
  readonly svgEl: SVGSVGElement | null;
  ensureEditable(action: string): boolean;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  getShapeTextParagraphs(shape: Element): (SVGTextElement | SVGTSpanElement)[];
  getParagraphLeafText(element: SVGTextElement | SVGTSpanElement): string;
  getSvgInlineSelectionBoxes(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number
  ): SvgInlineSelectionBox[];
  formatSvgNumber(value: number): string;
}

/**
 * Owns the floating find/replace panel, the in-slide match highlighting, and
 * text replacement. Extracted verbatim from `NativePowerPointView`; it borrows
 * shared editor state through {@link FindReplaceHost}.
 */
export class FindReplaceController {
  private findMatches: PowerPointFindMatch[] = [];
  private currentFindMatchIndex = 0;
  private findHighlightRects: SVGRectElement[] = [];

  private findPanelEl: HTMLElement | null = null;
  private findInputEl: HTMLInputElement | null = null;
  private findStatusEl: HTMLElement | null = null;
  private findReplaceInputEl: HTMLInputElement | null = null;
  private findReplaceToggleEl: HTMLButtonElement | null = null;
  private findButtonEl: HTMLButtonElement | null = null;
  private findPanelDismissHandler: ((event: Event) => void) | null = null;
  private findPanelRepositionHandler: (() => void) | null = null;
  private isFindReplaceMode = false;

  constructor(private readonly host: FindReplaceHost) {}

  /**
   * Builds the floating find/replace panel and anchors it to the toolbar's
   * search button. Must be called once the toolbar button exists.
   */
  createPanel(anchorButton: HTMLButtonElement): void {
    this.findButtonEl = anchorButton;

    // Mounted on <body> (not inside the editor layout) so the fixed-position
    // dropdown is positioned relative to the viewport. Ancestors in the editor
    // tree use CSS transforms for zoom, which would otherwise make a
    // position:fixed child resolve against the transformed box and render in
    // the wrong place.
    this.findPanelEl?.remove();
    const panel = activeDocument.body.createDiv({ cls: 'native-powerpoint-find-panel native-powerpoint-find-panel-floating native-powerpoint-light-surface' });
    this.findPanelEl = panel;

    const findRow = panel.createDiv({ cls: 'native-powerpoint-find-row' });

    const toggleButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn native-powerpoint-find-replace-toggle',
      attr: { 'aria-label': 'Toggle replace' }
    });
    setIcon(toggleButton, 'chevron-right');
    this.findReplaceToggleEl = toggleButton;

    const input = findRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'search',
      attr: {
        'aria-label': 'Find text in presentation',
        placeholder: 'Find text'
      }
    });
    this.findInputEl = input;

    this.findStatusEl = findRow.createDiv({ cls: 'native-powerpoint-find-status', text: 'No search' });

    const previousButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Previous match' }
    });
    setIcon(previousButton, 'chevron-up');

    const nextButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Next match' }
    });
    setIcon(nextButton, 'chevron-down');

    const closeButton = findRow.createEl('button', {
      cls: 'native-powerpoint-find-btn',
      attr: { 'aria-label': 'Close find' }
    });
    setIcon(closeButton, 'x');

    const replaceRow = panel.createDiv({ cls: 'native-powerpoint-find-replace-row' });

    const replaceInput = replaceRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'text',
      attr: {
        'aria-label': 'Replacement text',
        placeholder: 'Replace with'
      }
    });
    this.findReplaceInputEl = replaceInput;

    const replaceButton = replaceRow.createEl('button', {
      cls: 'native-powerpoint-find-replace-btn',
      text: 'Replace',
      attr: { 'aria-label': 'Replace current match' }
    });

    const replaceAllButton = replaceRow.createEl('button', {
      cls: 'native-powerpoint-find-replace-btn',
      text: 'Replace all',
      attr: { 'aria-label': 'Replace all matches' }
    });

    input.addEventListener('input', () => {
      void this.refreshFindMatches({ reveal: true });
    });
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeFindPanel();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        void this.moveFindMatch(event.shiftKey ? -1 : 1);
      }
    });
    replaceInput.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.closeFindPanel();
      } else if (event.key === 'Enter') {
        event.preventDefault();
        if (event.shiftKey) {
          void this.replaceAllMatches();
        } else {
          void this.replaceCurrentMatch();
        }
      }
    });
    toggleButton.addEventListener('click', () => this.setFindReplaceMode(!this.isFindReplaceMode));
    previousButton.addEventListener('click', () => void this.moveFindMatch(-1));
    nextButton.addEventListener('click', () => void this.moveFindMatch(1));
    closeButton.addEventListener('click', () => this.closeFindPanel());
    replaceButton.addEventListener('click', () => void this.replaceCurrentMatch());
    replaceAllButton.addEventListener('click', () => void this.replaceAllMatches());

    this.setFindReplaceMode(false);
  }

  private setFindReplaceMode(enabled: boolean): void {
    this.isFindReplaceMode = enabled;
    this.findPanelEl?.toggleClass('is-replace-mode', enabled);
    if (this.findReplaceToggleEl) {
      setIcon(this.findReplaceToggleEl, enabled ? 'chevron-down' : 'chevron-right');
      this.findReplaceToggleEl.setAttribute('aria-expanded', String(enabled));
    }
  }

  open(options: { replace?: boolean } = {}): void {
    if (!this.host.engine || this.host.isLoading) {
      new Notice('Open a loaded PowerPoint file to search it.');
      return;
    }

    this.findPanelEl?.addClass('is-open');
    debugLog('search', 'Opened PowerPoint find panel', {
      replaceMode: options.replace === true,
      slide: this.host.currentSlide
    });
    if (options.replace) {
      this.setFindReplaceMode(true);
    }
    this.findButtonEl?.addClass('is-active');
    const seedText = this.getSelectedFindSeedText();
    if (seedText && this.findInputEl && !this.findInputEl.value.trim()) {
      this.findInputEl.value = seedText;
    }

    this.attachFindPanelDismissHandlers();
    void this.refreshFindMatches({ reveal: Boolean(this.findInputEl?.value.trim()) });
    window.requestAnimationFrame(() => {
      this.positionFindPanel();
      this.findInputEl?.focus();
      this.findInputEl?.select();
    });
  }

  toggle(options: { replace?: boolean } = {}): void {
    if (this.findPanelEl?.hasClass('is-open')) {
      this.closeFindPanel();
      return;
    }
    this.open(options);
  }

  /** Tears down the panel and its global listeners. Call from the view's onClose. */
  dispose(): void {
    this.detachFindPanelDismissHandlers();
    this.findPanelEl?.remove();
    this.findPanelEl = null;
  }

  /** Clears match state and panel inputs when a presentation is (un)loaded. */
  reset(): void {
    this.findMatches = [];
    this.currentFindMatchIndex = 0;
    if (this.findInputEl) this.findInputEl.value = '';
    if (this.findReplaceInputEl) this.findReplaceInputEl.value = '';
    this.setFindReplaceMode(false);
    this.closeFindPanel();
    this.updateFindStatus();
  }

  /** Re-applies in-slide match highlighting after the current slide re-renders. */
  refreshHighlight(): void {
    this.applyFindHighlight();
  }

  private positionFindPanel(): void {
    const panel = this.findPanelEl;
    const anchor = this.findButtonEl;
    if (!panel || !anchor || !panel.hasClass('is-open')) return;

    const rect = anchor.getBoundingClientRect();
    const panelWidth = panel.offsetWidth || 320;
    const gap = 6;
    const margin = 8;
    let left = rect.right - panelWidth;
    left = Math.max(margin, Math.min(left, window.innerWidth - panelWidth - margin));
    const top = Math.min(rect.bottom + gap, window.innerHeight - panel.offsetHeight - margin);
    panel.style.left = `${Math.round(left)}px`;
    panel.style.top = `${Math.round(Math.max(margin, top))}px`;
  }

  private attachFindPanelDismissHandlers(): void {
    if (!this.findPanelDismissHandler) {
      this.findPanelDismissHandler = (event: Event) => {
        const target = event.target;
        if (!isNode(target)) return;
        if (this.findPanelEl?.contains(target)) return;
        if (this.findButtonEl?.contains(target)) return;
        this.closeFindPanel();
      };
      activeDocument.addEventListener('pointerdown', this.findPanelDismissHandler, true);
    }
    if (!this.findPanelRepositionHandler) {
      this.findPanelRepositionHandler = () => this.positionFindPanel();
      window.addEventListener('resize', this.findPanelRepositionHandler);
      window.addEventListener('scroll', this.findPanelRepositionHandler, true);
    }
  }

  private detachFindPanelDismissHandlers(): void {
    if (this.findPanelDismissHandler) {
      activeDocument.removeEventListener('pointerdown', this.findPanelDismissHandler, true);
      this.findPanelDismissHandler = null;
    }
    if (this.findPanelRepositionHandler) {
      window.removeEventListener('resize', this.findPanelRepositionHandler);
      window.removeEventListener('scroll', this.findPanelRepositionHandler, true);
      this.findPanelRepositionHandler = null;
    }
  }

  private async replaceCurrentMatch(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('replace text')) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findInputEl?.focus();
      return;
    }

    if (this.findMatches.length === 0) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) {
        new Notice('No matches to replace.');
        return;
      }
    }

    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match || match.shapeIndex === null) {
      new Notice('Select a match to replace, or use Replace all.');
      return;
    }

    const replacement = this.findReplaceInputEl?.value ?? '';
    try {
      const history = await this.host.captureHistoryEntry('Replace text');
      const count = await this.host.engine.replaceText(query, replacement, {
        slideIndex: match.slideIndex,
        shapeIndex: match.shapeIndex
      });
      if (count === 0) {
        new Notice('No matches to replace.');
        return;
      }
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.host.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      debugLog('search', 'Replaced current PowerPoint find match', {
        queryLength: query.length,
        replacementLength: replacement.length,
        replacedCount: count,
        slide: match.slideIndex,
        shapeIndex: match.shapeIndex
      });
      new Notice(count === 1 ? 'Replaced 1 match.' : `Replaced ${count} matches.`);
    } catch (error) {
      errorLog('search', 'PowerPoint replace-current failed', {
        queryLength: query.length,
        replacementLength: replacement.length,
        error
      });
      new Notice(`Could not replace text: ${cleanError(error)}`);
    }
  }

  private async replaceAllMatches(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('replace text')) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findInputEl?.focus();
      return;
    }

    const replacement = this.findReplaceInputEl?.value ?? '';
    try {
      const history = await this.host.captureHistoryEntry('Replace all text');
      const count = await this.host.engine.replaceText(query, replacement);
      if (count === 0) {
        new Notice('No matches to replace.');
        return;
      }
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.host.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      debugLog('search', 'Replaced all PowerPoint find matches', {
        queryLength: query.length,
        replacementLength: replacement.length,
        replacedCount: count
      });
      new Notice(count === 1 ? 'Replaced 1 match.' : `Replaced ${count} matches.`);
    } catch (error) {
      errorLog('search', 'PowerPoint replace-all failed', {
        queryLength: query.length,
        replacementLength: replacement.length,
        error
      });
      new Notice(`Could not replace text: ${cleanError(error)}`);
    }
  }

  private closeFindPanel(): void {
    this.findPanelEl?.removeClass('is-open');
    this.findButtonEl?.removeClass('is-active');
    this.detachFindPanelDismissHandlers();
    this.clearFindHighlight();
  }

  private getSelectedFindSeedText(): string {
    const activeEditor = this.host.activeEditor;
    if (
      activeEditor
      && activeEditor.selectionStart !== null
      && activeEditor.selectionEnd !== null
      && activeEditor.selectionStart !== activeEditor.selectionEnd
    ) {
      return normalizeSearchText(activeEditor.value.slice(activeEditor.selectionStart, activeEditor.selectionEnd));
    }

    return '';
  }

  private async refreshFindMatches(options: { reveal?: boolean } = {}): Promise<void> {
    const query = this.findInputEl?.value.trim() ?? '';
    this.clearFindHighlight();

    if (!this.host.engine || !query) {
      this.findMatches = [];
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      debugLog('search', 'PowerPoint find results cleared', {
        queryLength: query.length,
        hasEngine: Boolean(this.host.engine)
      });
      return;
    }

    this.findMatches = this.collectFindMatches(query);
    if (this.findMatches.length === 0) {
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      debugLog('search', 'PowerPoint find completed', {
        queryLength: query.length,
        matchCount: 0,
        slideCount: this.host.engine.slideCount
      });
      return;
    }

    const currentSlideMatchIndex = this.findMatches.findIndex((match) => match.slideIndex >= this.host.currentSlide);
    this.currentFindMatchIndex = currentSlideMatchIndex === -1 ? 0 : currentSlideMatchIndex;

    if (options.reveal) {
      await this.revealCurrentFindMatch();
    } else {
      this.applyFindHighlight();
      this.updateFindStatus();
    }
    debugLog('search', 'PowerPoint find completed', {
      queryLength: query.length,
      matchCount: this.findMatches.length,
      slideCount: this.host.engine.slideCount,
      selectedMatch: this.currentFindMatchIndex + 1,
      reveal: options.reveal === true
    });
  }

  private collectFindMatches(query: string): PowerPointFindMatch[] {
    const engine = this.host.engine;
    if (!engine) return [];

    const queryLower = query.toLocaleLowerCase();
    const matches: PowerPointFindMatch[] = [];
    const parser = new DOMParser();

    for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
      const { svg } = engine.renderSlide(slideIndex);
      const slideDocument = parser.parseFromString(svg, 'image/svg+xml');
      const shapeElements = Array.from(slideDocument.querySelectorAll('g[data-ooxml-shape-idx]'));
      let foundShapeMatch = false;

      for (const shape of shapeElements) {
        const shapeIndex = getShapeIndex(shape);
        if (shapeIndex === null) continue;

        const text = normalizeSearchText(shape.textContent ?? '');
        if (!text || !text.toLocaleLowerCase().includes(queryLower)) continue;

        foundShapeMatch = true;
        matches.push({ slideIndex, shapeIndex, text });
      }

      if (!foundShapeMatch) {
        const slideText = normalizeSearchText(slideDocument.documentElement.textContent ?? '');
        if (slideText && slideText.toLocaleLowerCase().includes(queryLower)) {
          matches.push({ slideIndex, shapeIndex: null, text: slideText });
        }
      }
    }

    return matches;
  }

  private async moveFindMatch(direction: -1 | 1): Promise<void> {
    if (!this.findInputEl?.value.trim()) {
      this.open();
      return;
    }

    if (this.findMatches.length === 0) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) return;
    }

    this.currentFindMatchIndex = (this.currentFindMatchIndex + direction + this.findMatches.length) % this.findMatches.length;
    await this.revealCurrentFindMatch();
  }

  private async revealCurrentFindMatch(): Promise<void> {
    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match) {
      this.updateFindStatus();
      return;
    }

    if (match.slideIndex !== this.host.currentSlide) {
      this.host.currentSlide = match.slideIndex;
      this.host.selectedShapeIndex = null;
      this.host.selectedTransform = null;
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        await this.host.renderThumbnails();
      }
    }

    // Intentionally do not select the matched shape: a selection would draw a
    // box/outline around the whole text frame. Find should only highlight the
    // matched characters themselves.
    this.applyFindHighlight();
    this.updateFindStatus();
  }

  private clearFindHighlight(): void {
    for (const rect of this.findHighlightRects) {
      rect.remove();
    }
    this.findHighlightRects = [];
    this.host.svgEl?.querySelectorAll('.native-powerpoint-find-current').forEach((element) => {
      element.removeClass('native-powerpoint-find-current');
    });
  }

  private applyFindHighlight(): void {
    this.clearFindHighlight();
    if (!this.host.svgEl) return;

    const query = this.findInputEl?.value ?? '';
    const trimmed = query.trim();
    if (!trimmed) return;

    const queryLower = trimmed.toLocaleLowerCase();
    const currentMatch = this.findMatches[this.currentFindMatchIndex];
    const currentShapeIndex = currentMatch && currentMatch.slideIndex === this.host.currentSlide
      ? currentMatch.shapeIndex
      : null;

    for (const shape of Array.from(this.host.svgEl.querySelectorAll('g[data-ooxml-shape-idx]'))) {
      const shapeIndex = getShapeIndex(shape);
      const isCurrent = shapeIndex !== null && shapeIndex === currentShapeIndex;
      this.highlightFindOccurrencesInShape(shape, queryLower, isCurrent);
    }
  }

  private highlightFindOccurrencesInShape(shape: Element, queryLower: string, isCurrent: boolean): void {
    if (!queryLower) return;
    for (const paragraph of this.host.getShapeTextParagraphs(shape)) {
      const text = this.host.getParagraphLeafText(paragraph);
      if (!text) continue;

      const lower = text.toLocaleLowerCase();
      let from = 0;
      while (from <= lower.length) {
        const index = lower.indexOf(queryLower, from);
        if (index === -1) break;
        this.renderFindHighlightRects(paragraph, index, index + queryLower.length, isCurrent);
        from = index + Math.max(1, queryLower.length);
      }
    }
  }

  private renderFindHighlightRects(
    element: SVGTextElement | SVGTSpanElement,
    start: number,
    end: number,
    isCurrent: boolean
  ): void {
    const boxes = this.host.getSvgInlineSelectionBoxes(element, start, end);
    const textElement = element.closest('text');
    const parent = textElement?.parentNode;
    if (!isSVGTextElement(textElement) || !parent) return;

    for (const box of boxes) {
      const rect = activeDocument.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.classList.add('native-powerpoint-find-highlight');
      if (isCurrent) rect.classList.add('is-current');
      rect.setAttribute('x', this.host.formatSvgNumber(box.x));
      rect.setAttribute('y', this.host.formatSvgNumber(box.y));
      rect.setAttribute('width', this.host.formatSvgNumber(box.width));
      rect.setAttribute('height', this.host.formatSvgNumber(box.height));
      rect.setAttribute('rx', '1');
      parent.insertBefore(rect, textElement);
      this.findHighlightRects.push(rect);
    }
  }

  private updateFindStatus(): void {
    if (!this.findStatusEl) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findStatusEl.setText('No search');
      this.findStatusEl.removeAttribute('title');
      return;
    }

    if (this.findMatches.length === 0) {
      this.findStatusEl.setText('No matches');
      this.findStatusEl.removeAttribute('title');
      return;
    }

    const match = this.findMatches[this.currentFindMatchIndex];
    const slideLabel = match ? `Slide ${match.slideIndex + 1}` : 'Slide';
    this.findStatusEl.setText(`${this.currentFindMatchIndex + 1} / ${this.findMatches.length} | ${slideLabel}`);
    if (match) {
      this.findStatusEl.setAttribute('title', match.text);
    }
  }
}
