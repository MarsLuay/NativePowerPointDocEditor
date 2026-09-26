import { Component, setIcon } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import type { PresentationEngine } from '../PresentationEngine';
import { isNode, isSVGTextElement } from '../domGuards';
import { debugLog, errorLog } from '../logger';
import { formatFindResultStatus, wrapMatchIndex } from '../find/findReplaceShell';
import { createMenuItem, createPopoverShell, createToolbarIconButton } from '../menuControls';
import { cleanError } from './runtimeCompat';
import {
  collectFindMatchesFromSearchIndex,
  createFindSearchIndexSlideFromOoxml,
  type PowerPointFindSearchIndexSlide,
} from './findSearchIndex';
import { normalizeSearchText } from './textUtils';
import { getShapeIndex } from './svgUtils';
import type { PresentationSession } from './session/PresentationSession';
import type { HistoryEntry, PowerPointFindMatch, SvgInlineSelectionBox } from './types';

const FIND_INDEX_YIELD_BUDGET_MS = 8;
const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';

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
  readonly t: TranslateFn;
  readonly session: PresentationSession;
  readonly engine: PresentationEngine | null;
  readonly isLoading: boolean;
  readonly currentSlide: number;
  readonly activeEditor: HTMLTextAreaElement | null;
  readonly svgEl: SVGSVGElement | null;
  ensureEditable(action: string): boolean;
  clearSelection(): void;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  navigateToSlide(index: number, reason: string): Promise<void>;
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
  private findMatchesQuery = '';
  private currentFindMatchIndex = 0;
  private findHighlightRects: SVGRectElement[] = [];

  private findPanelEl: HTMLElement | null = null;
  private findInputEl: HTMLInputElement | null = null;
  private findStatusEl: HTMLElement | null = null;
  private findReplaceInputEl: HTMLInputElement | null = null;
  private findReplaceToggleEl: HTMLButtonElement | null = null;
  private findButtonEl: HTMLButtonElement | null = null;
  private findPanelDomScope: Component | null = null;
  private findPanelDismissHandler: ((event: Event) => void) | null = null;
  private findPanelRepositionHandler: (() => void) | null = null;
  private findRefreshRequestId = 0;
  private findSearchIndex: PowerPointFindSearchIndexSlide[] | null = null;
  private findSearchIndexEngine: PresentationEngine | null = null;
  private findSearchIndexBuild: {
    engine: PresentationEngine;
    generation: number;
    promise: Promise<PowerPointFindSearchIndexSlide[] | null>;
  } | null = null;
  private findSearchIndexGeneration = 0;
  private isFindReplaceMode = false;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: FindReplaceHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  /**
   * Builds the floating find/replace panel and anchors it to the toolbar's
   * search button. Must be called once the toolbar button exists.
   */
  createPanel(anchorButton: HTMLButtonElement): void {
    debugLog('search', 'PowerPoint find panel setup started', { op: 'create-panel' });
    this.findButtonEl = anchorButton;

    // Mounted on <body> (not inside the editor layout) so the fixed-position
    // dropdown is positioned relative to the viewport. Ancestors in the editor
    // tree use CSS transforms for zoom, which would otherwise make a
    // position:fixed child resolve against the transformed box and render in
    // the wrong place.
    this.findPanelEl?.remove();
    const panel = createPopoverShell(activeDocument.body, {
      className: 'native-powerpoint-find-panel native-powerpoint-find-panel-floating native-powerpoint-light-surface'
    });
    this.findPanelEl = panel;

    const findRow = panel.createDiv({ cls: 'native-powerpoint-find-row' });

    const toggleButton = createToolbarIconButton(findRow, {
      className: ['native-powerpoint-find-btn', 'native-powerpoint-find-replace-toggle'],
      icon: 'chevron-right',
      label: this.host.t('powerpoint:find.toggleReplace')
    });
    this.findReplaceToggleEl = toggleButton;

    const input = findRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'search',
      attr: {
        'aria-label': this.host.t('powerpoint:find.ariaLabel'),
        placeholder: this.host.t('powerpoint:find.placeholder')
      }
    });
    this.findInputEl = input;

    this.findStatusEl = findRow.createDiv({ cls: 'native-powerpoint-find-status', text: this.host.t('powerpoint:find.noSearch') });

    const previousButton = createToolbarIconButton(findRow, {
      className: 'native-powerpoint-find-btn',
      icon: 'chevron-up',
      label: this.host.t('powerpoint:find.previousMatch')
    });

    const nextButton = createToolbarIconButton(findRow, {
      className: 'native-powerpoint-find-btn',
      icon: 'chevron-down',
      label: this.host.t('powerpoint:find.nextMatch')
    });

    const closeButton = createToolbarIconButton(findRow, {
      className: 'native-powerpoint-find-btn',
      icon: 'x',
      label: this.host.t('powerpoint:find.closeFind')
    });

    const replaceRow = panel.createDiv({ cls: 'native-powerpoint-find-replace-row' });

    const replaceInput = replaceRow.createEl('input', {
      cls: 'native-powerpoint-find-input',
      type: 'text',
      attr: {
        'aria-label': this.host.t('powerpoint:find.replacementAriaLabel'),
        placeholder: this.host.t('powerpoint:find.replacePlaceholder')
      }
    });
    this.findReplaceInputEl = replaceInput;

    const replaceButton = createMenuItem(replaceRow, {
      className: 'native-powerpoint-find-replace-btn',
      text: this.host.t('powerpoint:find.replace'),
      ariaLabel: this.host.t('powerpoint:find.replaceCurrentMatch')
    });

    const replaceAllButton = createMenuItem(replaceRow, {
      className: 'native-powerpoint-find-replace-btn',
      text: this.host.t('powerpoint:find.replaceAll'),
      ariaLabel: this.host.t('powerpoint:find.replaceAllMatches')
    });

    input.addEventListener('input', () => this.scheduleFindRefresh());
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
      this.notice('powerpoint:notice.openLoadedToSearch');
      return;
    }

    this.findPanelEl?.addClass('is-open');
    debugLog('search', 'PowerPoint find panel open started', {
      op: 'open',
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
    debugLog('search', 'PowerPoint find panel toggle started', {
      op: 'toggle',
      replaceMode: options.replace === true,
      wasOpen: this.findPanelEl?.hasClass('is-open') === true
    });
    if (this.findPanelEl?.hasClass('is-open')) {
      this.closeFindPanel();
      return;
    }
    this.open(options);
  }

  /** Tears down the panel and its global listeners. Call from the view's onClose. */
  dispose(): void {
    debugLog('search', 'PowerPoint find panel dispose started', { op: 'dispose' });
    this.beginFindRefreshRequest();
    this.invalidateFindSearchIndex('dispose');
    this.detachFindPanelDismissHandlers();
    this.findPanelEl?.remove();
    this.findPanelEl = null;
  }

  /** Clears match state and panel inputs when a presentation is (un)loaded. */
  reset(): void {
    debugLog('search', 'PowerPoint find reset started', { op: 'reset' });
    this.beginFindRefreshRequest();
    this.invalidateFindSearchIndex('reset');
    this.findMatches = [];
    this.findMatchesQuery = '';
    this.currentFindMatchIndex = 0;
    if (this.findInputEl) this.findInputEl.value = '';
    if (this.findReplaceInputEl) this.findReplaceInputEl.value = '';
    this.setFindReplaceMode(false);
    this.closeFindPanel();
    this.updateFindStatus();
  }

  /** Re-applies in-slide match highlighting after the current slide re-renders. */
  refreshHighlight(): void {
    this.invalidateFindSearchIndex('slide-render');
    const panelOpen = this.isFindPanelOpen();
    debugLog('search', 'PowerPoint find highlight refresh started', {
      op: 'refresh-highlight',
      matchCount: this.findMatches.length,
      panelOpen,
    });
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
		panel.setCssProps({
			left: `${Math.round(left)}px`,
			top: `${Math.round(Math.max(margin, top))}px`,
		});
  }

  private attachFindPanelDismissHandlers(): void {
    if (this.findPanelDomScope) {
      return;
    }

    this.findPanelDismissHandler = (event: Event) => {
      const target = event.target;
      if (!isNode(target)) return;
      if (this.findPanelEl?.contains(target)) return;
      if (this.findButtonEl?.contains(target)) return;
      this.closeFindPanel();
    };
    this.findPanelRepositionHandler = () => this.positionFindPanel();
    this.findPanelDomScope = new Component();
    this.findPanelDomScope.registerDomEvent(activeDocument, 'pointerdown', this.findPanelDismissHandler, true);
    this.findPanelDomScope.registerDomEvent(window, 'resize', this.findPanelRepositionHandler);
    this.findPanelDomScope.registerDomEvent(window, 'scroll', this.findPanelRepositionHandler, true);
    this.findPanelDomScope.load();
  }

  private detachFindPanelDismissHandlers(): void {
    this.findPanelDomScope?.unload();
    this.findPanelDomScope = null;
    this.findPanelDismissHandler = null;
    this.findPanelRepositionHandler = null;
  }

  private async replaceCurrentMatch(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('replace text')) return;

    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.findInputEl?.focus();
      return;
    }

    if (this.findMatches.length === 0 || this.findMatchesQuery !== query) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) {
        this.notice('powerpoint:notice.noMatchesToReplace');
        return;
      }
    }

    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match || match.shapeIndex === null) {
      this.notice('powerpoint:notice.selectMatchToReplace');
      return;
    }

    const replacement = this.findReplaceInputEl?.value ?? '';
    try {
      debugLog('search', 'PowerPoint replace-current started', {
        op: 'replace-current',
        queryLength: query.length,
        replacementLength: replacement.length,
        slide: match.slideIndex,
        shapeIndex: match.shapeIndex
      });
      const history = await this.host.captureHistoryEntry('Replace text');
      const count = await this.host.session.applyCommand({
        type: 'replace-text',
        query,
        replacement,
        slideIndex: match.slideIndex,
        shapeIndex: match.shapeIndex
      }) as number;
      if (count === 0) {
        this.notice('powerpoint:notice.noMatchesToReplace');
        return;
      }
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.host.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      debugLog('search', 'Replaced current PowerPoint find match', {
        op: 'replace-current',
        queryLength: query.length,
        replacementLength: replacement.length,
        replacedCount: count,
        slide: match.slideIndex,
        shapeIndex: match.shapeIndex
      });
      this.notice('powerpoint:notice.replacedMatches', { count });
    } catch (error) {
      errorLog('search', 'PowerPoint replace-current failed', {
        op: 'replace-current',
        queryLength: query.length,
        replacementLength: replacement.length,
        error
      });
      this.notice('powerpoint:notice.couldNotReplaceText', { message: cleanError(error) });
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
      debugLog('search', 'PowerPoint replace-all started', {
        op: 'replace-all',
        queryLength: query.length,
        replacementLength: replacement.length
      });
      const history = await this.host.captureHistoryEntry('Replace all text');
      const count = await this.host.session.applyCommand({
        type: 'replace-text',
        query,
        replacement
      }) as number;
      if (count === 0) {
        this.notice('powerpoint:notice.noMatchesToReplace');
        return;
      }
      this.host.recordHistoryEntry(history);
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.host.renderThumbnails();
      await this.refreshFindMatches({ reveal: true });
      debugLog('search', 'Replaced all PowerPoint find matches', {
        op: 'replace-all',
        queryLength: query.length,
        replacementLength: replacement.length,
        replacedCount: count
      });
      this.notice('powerpoint:notice.replacedMatches', { count });
    } catch (error) {
      errorLog('search', 'PowerPoint replace-all failed', {
        op: 'replace-all',
        queryLength: query.length,
        replacementLength: replacement.length,
        error
      });
      this.notice('powerpoint:notice.couldNotReplaceText', { message: cleanError(error) });
    }
  }

  private closeFindPanel(): void {
    const wasOpen = this.isFindPanelOpen();
    this.beginFindRefreshRequest();
    this.findPanelEl?.removeClass('is-open');
    this.findButtonEl?.removeClass('is-active');
    this.detachFindPanelDismissHandlers();
    this.clearFindHighlight();
    debugLog('search', 'PowerPoint find panel closed', {
      op: 'close',
      wasOpen,
      matchCount: this.findMatches.length,
      slide: this.host.currentSlide,
    });
  }

  private isFindPanelOpen(): boolean {
    return this.findPanelEl?.hasClass('is-open') === true;
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

  private scheduleFindRefresh(): void {
    const query = this.findInputEl?.value.trim() ?? '';
    const requestId = this.beginFindRefreshRequest();
    this.clearFindHighlight();
    if (!query) {
      void this.refreshFindMatches({ requestId });
      return;
    }

    this.findStatusEl?.setText(this.host.t('powerpoint:find.searching'));
    void this.refreshFindMatches({ requestId });
  }

  private beginFindRefreshRequest(): number {
    this.findRefreshRequestId += 1;
    return this.findRefreshRequestId;
  }

  private async refreshFindMatches(options: { reveal?: boolean; requestId?: number } = {}): Promise<void> {
    const requestId = options.requestId ?? this.beginFindRefreshRequest();
    const query = this.findInputEl?.value.trim() ?? '';
    this.clearFindHighlight();

    if (!this.host.engine || !query) {
      this.findMatches = [];
      this.findMatchesQuery = '';
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      debugLog('search', 'PowerPoint find results cleared', {
        op: 'refresh-matches',
        queryLength: query.length,
        hasEngine: Boolean(this.host.engine)
      });
      return;
    }

    const startedAt = performance.now();
    let searchIndex: PowerPointFindSearchIndexSlide[] | null;
    try {
      searchIndex = await this.getFindSearchIndex();
      debugLog('search', 'PowerPoint find index received by refresh', {
        op: 'refresh-matches',
        queryLength: query.length,
        hasSearchIndex: Boolean(searchIndex),
        requestId,
        activeRequestId: this.findRefreshRequestId,
      });
    } catch (error) {
      if (requestId !== this.findRefreshRequestId) return;
      this.findMatches = [];
      this.findMatchesQuery = '';
      this.currentFindMatchIndex = 0;
      this.updateFindStatus();
      errorLog('search', 'PowerPoint find index build failed', {
        op: 'build-index',
        queryLength: query.length,
        error,
      });
      return;
    }
    try {
      const activeQuery = this.findInputEl?.value.trim() ?? '';
      const panelOpen = this.findPanelEl?.hasClass('is-open') === true;
      if (query !== activeQuery || !panelOpen || !searchIndex) {
        debugLog('search', 'PowerPoint find result refresh skipped', {
          op: 'refresh-matches',
          queryLength: query.length,
          activeQueryLength: activeQuery.length,
          panelOpen,
          hasSearchIndex: Boolean(searchIndex),
          requestId,
          activeRequestId: this.findRefreshRequestId,
        });
        return;
      }

      this.findMatches = collectFindMatchesFromSearchIndex(searchIndex, query);
      this.findMatchesQuery = query;
      const searchMs = Math.round(performance.now() - startedAt);
      if (this.findMatches.length === 0) {
        this.currentFindMatchIndex = 0;
        this.updateFindStatus();
        debugLog('search', 'PowerPoint find completed', {
          op: 'refresh-matches',
          queryLength: query.length,
          matchCount: 0,
          slideCount: this.host.engine.slideCount,
          searchMs
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
        op: 'refresh-matches',
        queryLength: query.length,
        matchCount: this.findMatches.length,
        slideCount: this.host.engine.slideCount,
        selectedMatch: this.currentFindMatchIndex + 1,
        reveal: options.reveal === true,
        searchMs
      });
    } catch (error) {
      errorLog('search', 'PowerPoint find refresh failed', {
        op: 'refresh-matches',
        queryLength: query.length,
        requestId,
        error,
      });
    }
  }

  private async getFindSearchIndex(): Promise<PowerPointFindSearchIndexSlide[] | null> {
    const engine = this.host.engine;
    if (!engine) return null;
    if (this.findSearchIndex && this.findSearchIndexEngine === engine) {
      return this.findSearchIndex;
    }
    const generation = this.findSearchIndexGeneration;
    if (
      this.findSearchIndexBuild
      && this.findSearchIndexBuild.engine === engine
      && this.findSearchIndexBuild.generation === generation
    ) {
      return this.findSearchIndexBuild.promise;
    }

    this.findSearchIndexEngine = engine;
    const promise = this.buildFindSearchIndex(engine, generation);
    this.findSearchIndexBuild = { engine, generation, promise };
    const clearBuild = () => {
      if (this.findSearchIndexBuild?.promise === promise) {
        this.findSearchIndexBuild = null;
      }
    };
    void promise.then(clearBuild).catch(clearBuild);
    return promise;
  }

  private async buildFindSearchIndex(
    engine: PresentationEngine,
    generation: number,
  ): Promise<PowerPointFindSearchIndexSlide[] | null> {
    const startedAt = performance.now();
    const searchIndex: PowerPointFindSearchIndexSlide[] = [];
    let shapeCount = 0;
    let yieldedAt = startedAt;
    debugLog('search', 'PowerPoint find index build started', {
      op: 'build-index',
      slideCount: engine.slideCount,
    });

    for (let slideIndex = 0; slideIndex < engine.slideCount; slideIndex++) {
      if (generation !== this.findSearchIndexGeneration || engine !== this.host.engine) {
        return null;
      }
      const indexSlide = createFindSearchIndexSlideFromOoxml(slideIndex, engine.getSlideXml(slideIndex));
      shapeCount += indexSlide.shapeMatches.length;
      searchIndex.push(indexSlide);

      if (performance.now() - yieldedAt >= FIND_INDEX_YIELD_BUDGET_MS) {
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
        yieldedAt = performance.now();
      }
    }

    if (generation !== this.findSearchIndexGeneration || engine !== this.host.engine) {
      return null;
    }
    this.findSearchIndex = searchIndex;
    this.findSearchIndexEngine = engine;
    debugLog('search', 'PowerPoint find index build completed', {
      op: 'build-index',
      slideCount: searchIndex.length,
      shapeCount,
      ms: Math.round(performance.now() - startedAt),
    });
    return searchIndex;
  }

  private invalidateFindSearchIndex(reason: string): void {
    const hadIndex = this.findSearchIndex !== null || this.findSearchIndexBuild !== null;
    this.findSearchIndexGeneration += 1;
    this.findSearchIndex = null;
    this.findSearchIndexEngine = null;
    if (hadIndex) {
      debugLog('search', 'PowerPoint find index invalidated', { op: 'invalidate-index', reason });
    }
  }

  private async moveFindMatch(direction: -1 | 1): Promise<void> {
    const query = this.findInputEl?.value.trim() ?? '';
    if (!query) {
      this.open();
      return;
    }

    if (this.findMatches.length === 0 || this.findMatchesQuery !== query) {
      await this.refreshFindMatches();
      if (this.findMatches.length === 0) return;
    }

    this.currentFindMatchIndex = wrapMatchIndex(this.currentFindMatchIndex, direction, this.findMatches.length);
    await this.revealCurrentFindMatch();
  }

  private async revealCurrentFindMatch(): Promise<void> {
    const match = this.findMatches[this.currentFindMatchIndex];
    if (!match) {
      this.updateFindStatus();
      return;
    }

    if (match.slideIndex !== this.host.currentSlide) {
      await this.host.navigateToSlide(match.slideIndex, 'find-match');
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
    if (!this.isFindPanelOpen() || !this.host.svgEl) {
      debugLog('search', 'PowerPoint find highlight skipped', {
        op: 'apply-highlight',
        reason: this.isFindPanelOpen() ? 'missing-slide' : 'panel-closed',
      });
      return;
    }

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
    const ownerDocument = textElement.ownerDocument;

    for (const box of boxes) {
      // Obsidian's createSvg helper appends to its receiver. The text can live
      // in an SVG/XML document, where appending a second root throws
      // HierarchyRequestError. Create the rect detached, then insert it below.
      const rect = ownerDocument.createElementNS(SVG_NAMESPACE, 'rect');
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
    const match = query && this.findMatches.length > 0 ? this.findMatches[this.currentFindMatchIndex] : undefined;
    const slideLabel = match
      ? this.host.t('powerpoint:find.slideLabel', { slideNumber: match.slideIndex + 1 })
      : '';

    this.findStatusEl.setText(
      formatFindResultStatus(query, this.currentFindMatchIndex, this.findMatches.length, {
        noSearch: this.host.t('powerpoint:find.noSearch'),
        noMatches: this.host.t('powerpoint:find.noMatches'),
        resultCount: (current, total) =>
          this.host.t('powerpoint:find.resultCount', { current, total, slide: slideLabel })
      })
    );

    if (match) {
      this.findStatusEl.setAttribute('title', match.text);
    } else {
      this.findStatusEl.removeAttribute('title');
    }
  }
}
