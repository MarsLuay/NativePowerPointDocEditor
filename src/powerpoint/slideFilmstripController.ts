import { Menu } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import type { PresentationEngine, SlideLayoutKind } from '../PresentationEngine';
import { createSvgElementFromString, type SvgSecurityIssue } from '../SvgSecurity';
import type { ShapeTransform } from 'pptx-svg';
import { debugLog, errorLog, warnLog } from '../logger';
import { cleanError } from './runtimeCompat';
import { normalizeSvgForDisplay } from './svgUtils';
import type { HistoryEntry } from './types';
import { scheduleIdleWork } from '../idleSchedule';
import { isHTMLElement } from '../domGuards';
import {
	priorityThumbnailIndices,
	remainingThumbnailIndices,
	shouldUseLazyThumbnails,
	sortThumbnailIndicesByProximity,
	THUMBNAIL_IDLE_BATCH_SIZE,
} from './thumbnailLazyRender';

/**
 * The slice of `NativePowerPointView` that the slide filmstrip subsystem reaches
 * back into. Keeping this surface explicit is the regression boundary for the
 * extraction: as long as these members behave as they did when the methods
 * lived on the view, filmstrip behavior is unchanged.
 */
export interface SlideFilmstripHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  readonly thumbnailContainer: HTMLElement | null;
  readonly isLoading: boolean;
  currentSlide: number;
  lastInteractionRegion: 'canvas' | 'thumbnails';
  selectedShapeIndex: number | null;
  selectedTransform: ShapeTransform | null;
  slideRenderGeneration: number;
  isNavigatingSlide: boolean;
  canEdit(): boolean;
  ensureEditable(action: string): boolean;
  finishInlineTextEditing(reason: string): Promise<void>;
  captureHistoryEntry(label: string): Promise<HistoryEntry>;
  recordHistoryEntry(entry: HistoryEntry): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean, expectedGeneration?: number): Promise<boolean>;
  clearSelection(options?: { skipTextCommit?: boolean }): void;
  renderInspector(): void;
  prepareSvgForRender(
    svg: string,
    isThumbnail?: boolean
  ): { svg: string; issues: SvgSecurityIssue[]; allowed: boolean };
  createNativeMenu(): Menu;
}

/**
 * Owns the slide thumbnail filmstrip, slide navigation, multi-slide selection,
 * and slide CRUD/reorder operations. Extracted verbatim from
 * `NativePowerPointView`; it borrows shared editor state through
 * {@link SlideFilmstripHost}.
 */
export class SlideFilmstripController {
  readonly selectedSlideIndices = new Set<number>();
  private thumbnailDragIndex: number | null = null;
  private slideNavigationPromise: Promise<void> = Promise.resolve();
  private readonly notice: TranslateNoticeFn;
  private pendingThumbnailIndices = new Set<number>();
  private thumbnailRefreshScheduled = false;
  private cancelThumbnailRefresh: (() => void) | null = null;
  private thumbnailRenderGeneration = 0;
  private thumbnailObserver: IntersectionObserver | null = null;
  private cancelIdleThumbnailFill: (() => void) | null = null;
  private renderedThumbnailIndices = new Set<number>();

  constructor(private readonly host: SlideFilmstripHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  dispose(): void {
    this.cancelThumbnailRefresh?.();
    this.cancelThumbnailRefresh = null;
    this.thumbnailRefreshScheduled = false;
    this.pendingThumbnailIndices.clear();
    this.disconnectThumbnailObserver();
    this.cancelIdleThumbnailFill?.();
    this.cancelIdleThumbnailFill = null;
    this.renderedThumbnailIndices.clear();
    this.thumbnailRenderGeneration += 1;
  }

  /**
   * Re-render one filmstrip thumbnail. Used after layout edits so we do not
   * rebuild every slide preview on each object move.
   */
  async renderThumbnailAt(index: number): Promise<void> {
    if (!this.host.engine || !this.host.thumbnailContainer) return;
    if (index < 0 || index >= this.host.engine.slideCount) return;

    const items = this.host.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail');
    const item = items.item(index);
    if (!isHTMLElement(item)) return;

    const started = performance.now();
    const preview = item.querySelector('.native-powerpoint-thumbnail-preview');
    if (!isHTMLElement(preview)) return;

    preview.empty();
    try {
      const safeSvg = this.host.prepareSvgForRender(this.host.engine.renderSlide(index).svg, true);
      const thumbnailSvg = createSvgElementFromString(safeSvg.svg, preview.ownerDocument);
      if (!thumbnailSvg) {
        throw new Error('Could not read thumbnail SVG.');
      }
      preview.appendChild(thumbnailSvg);
    } catch {
      preview.createDiv({
        cls: 'native-powerpoint-thumbnail-error',
        text: this.host.t('powerpoint:loading.thumbnailError')
      });
    }

    const thumbnailSvg = preview.querySelector('svg');
    if (thumbnailSvg) {
      this.host.engine.applyFontFidelity(thumbnailSvg);
      this.host.engine.formatChartAxisLabels(thumbnailSvg, index);
      normalizeSvgForDisplay(thumbnailSvg);
      thumbnailSvg.addClass('native-powerpoint-thumbnail-svg');
    }

    debugLog('render', 'PowerPoint renderThumbnailAt complete', {
      index,
      ms: Math.round(performance.now() - started)
    });
    this.markThumbnailRendered(index);
  }

  /**
   * Coalesce single-slide thumbnail refreshes onto idle time so drag commits
   * never block on rebuilding the full filmstrip.
   */
  scheduleThumbnailRefresh(indices: number | number[]): void {
    const list = Array.isArray(indices) ? indices : [indices];
    for (const index of list) {
      if (Number.isInteger(index) && index >= 0) {
        this.pendingThumbnailIndices.add(index);
      }
    }
    if (this.thumbnailRefreshScheduled) return;

    this.thumbnailRefreshScheduled = true;
    this.cancelThumbnailRefresh?.();
    this.cancelThumbnailRefresh = scheduleIdleWork(() => {
      this.thumbnailRefreshScheduled = false;
      const toRefresh = [...this.pendingThumbnailIndices];
      this.pendingThumbnailIndices.clear();
      void this.refreshThumbnailsAt(toRefresh);
    }, { timeout: 2000 });
  }

  private async refreshThumbnailsAt(indices: number[]): Promise<void> {
    if (indices.length === 0) return;

    const started = performance.now();
    const sorted = [...indices].sort((left, right) => left - right);
    for (const index of sorted) {
      await this.renderThumbnailAt(index);
    }
    debugLog('render', 'PowerPoint thumbnail refresh completed', {
      indices: sorted,
      ms: Math.round(performance.now() - started)
    });
  }

  async renderThumbnails(options?: { preferLazy?: boolean }): Promise<void> {
    if (!this.host.engine || !this.host.thumbnailContainer) return;

    const thumbnailStarted = performance.now();
    const engine = this.host.engine;
    const thumbnailContainer = this.host.thumbnailContainer;
    const slideCount = engine.slideCount;
    const generation = ++this.thumbnailRenderGeneration;
    const preferLazy = options?.preferLazy === true;
    // Structural edits (delete/move/duplicate) prefer lazy so the UI responds
    // after the active thumb instead of blocking on a full filmstrip rebuild.
    const lazy = preferLazy || shouldUseLazyThumbnails(slideCount);
    debugLog('render', 'PowerPoint renderThumbnails start', { slideCount, lazy, preferLazy });

    const { cx, cy } = await engine.getSlideSizeEmu();
    if (
      generation !== this.thumbnailRenderGeneration
      || engine !== this.host.engine
      || thumbnailContainer !== this.host.thumbnailContainer
    ) {
      return;
    }
    const aspectRatio = `${cx} / ${cy}`;
    thumbnailContainer.style.setProperty('--native-powerpoint-thumbnail-aspect-ratio', aspectRatio);
    debugLog('render', 'PowerPoint thumbnail aspect ratio resolved', { cx, cy, aspectRatio });

    this.disconnectThumbnailObserver();
    this.cancelIdleThumbnailFill?.();
    this.cancelIdleThumbnailFill = null;
    this.renderedThumbnailIndices.clear();
    thumbnailContainer.empty();

    for (let index = 0; index < slideCount; index += 1) {
      this.appendThumbnailShell(index, !lazy);
    }

    if (!lazy) {
      for (let index = 0; index < slideCount; index += 1) {
        if (generation !== this.thumbnailRenderGeneration) return;
        await this.renderThumbnailAt(index);
      }
    } else {
		// Render only the active slide before returning control to the editor. The
		// rest of the filmstrip fills through the observer/idle path below.
		const priority = priorityThumbnailIndices(this.host.currentSlide, slideCount, 0);
      await this.renderThumbnailBatch(priority, generation);
      if (generation !== this.thumbnailRenderGeneration) return;
      this.setupThumbnailObserver(generation);
      this.scheduleIdleThumbnailFill(generation);
    }

    const thumbnailMs = Math.round(performance.now() - thumbnailStarted);
    debugLog('render', 'PowerPoint renderThumbnails complete', {
      slideCount,
      lazy,
      preferLazy,
      renderedCount: this.renderedThumbnailIndices.size,
      ms: thumbnailMs
    });
    if (!lazy && thumbnailMs > 1500) {
      warnLog('render', 'slow PowerPoint renderThumbnails', { slideCount, ms: thumbnailMs });
    }
  }

  private appendThumbnailShell(index: number, renderImmediately: boolean): void {
    if (!this.host.thumbnailContainer) return;

    const item = this.host.thumbnailContainer.createDiv({ cls: 'native-powerpoint-thumbnail' });
    item.dataset.slideIndex = String(index);
    if (index === this.host.currentSlide) item.addClass('active');
    if (this.selectedSlideIndices.has(index)) item.addClass('is-selected');

    const preview = item.createDiv({
      cls: renderImmediately
        ? 'native-powerpoint-thumbnail-preview'
        : 'native-powerpoint-thumbnail-preview is-pending'
    });
    preview.dataset.thumbnailRendered = 'false';

    const numberEl = item.createDiv({ cls: 'native-powerpoint-thumbnail-number' });
    numberEl.textContent = String(index + 1);
    item.addEventListener('click', (event) => {
      this.host.lastInteractionRegion = 'thumbnails';
      if (event.shiftKey) {
        this.selectSlideRange(this.host.currentSlide, index);
      } else if (event.metaKey || event.ctrlKey) {
        this.toggleSlideSelection(index);
      } else {
        this.selectedSlideIndices.clear();
        this.selectedSlideIndices.add(index);
        this.navigateToSlide(index, 'thumbnail-click');
      }
    });
    item.addEventListener('contextmenu', (event) => {
      event.preventDefault();
      this.showSlideContextMenu(event, index);
    });
    this.registerThumbnailDrag(item, index);
  }

  private async renderThumbnailBatch(indices: number[], generation: number): Promise<void> {
    const sorted = [...indices].sort((left, right) => left - right);
    for (const index of sorted) {
      if (generation !== this.thumbnailRenderGeneration) return;
      if (this.renderedThumbnailIndices.has(index)) continue;
      await this.renderThumbnailAt(index);
    }
  }

  private markThumbnailRendered(index: number): void {
    this.renderedThumbnailIndices.add(index);
    const items = this.host.thumbnailContainer?.querySelectorAll('.native-powerpoint-thumbnail');
    const item = items?.item(index);
    const preview = item?.querySelector('.native-powerpoint-thumbnail-preview');
    if (isHTMLElement(preview)) {
      preview.removeClass('is-pending');
      preview.dataset.thumbnailRendered = 'true';
      this.thumbnailObserver?.unobserve(preview);
    }
  }

  private disconnectThumbnailObserver(): void {
    this.thumbnailObserver?.disconnect();
    this.thumbnailObserver = null;
  }

  private setupThumbnailObserver(generation: number): void {
    if (!this.host.thumbnailContainer || typeof IntersectionObserver === 'undefined') return;

    this.disconnectThumbnailObserver();
    this.thumbnailObserver = new IntersectionObserver((entries) => {
      if (generation !== this.thumbnailRenderGeneration) return;

      const toRender: number[] = [];
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const preview = entry.target;
        if (!isHTMLElement(preview)) continue;
        if (preview.dataset.thumbnailRendered === 'true') continue;
        const item = preview.closest('.native-powerpoint-thumbnail');
        const index = Number(item?.getAttribute('data-slide-index'));
        if (!Number.isInteger(index) || this.renderedThumbnailIndices.has(index)) continue;
        toRender.push(index);
      }

      if (toRender.length > 0) {
        void this.renderThumbnailBatch(toRender, generation);
      }
    }, {
      root: this.host.thumbnailContainer,
      rootMargin: '240px 0px',
    });

    this.host.thumbnailContainer
      .querySelectorAll('.native-powerpoint-thumbnail-preview')
      .forEach((preview) => {
        if (isHTMLElement(preview) && preview.dataset.thumbnailRendered !== 'true') {
          this.thumbnailObserver?.observe(preview);
        }
      });
  }

  private scheduleIdleThumbnailFill(generation: number): void {
    const fillBatch = () => {
      if (generation !== this.thumbnailRenderGeneration || !this.host.engine) return;

      const slideCount = this.host.engine.slideCount;
      const remaining = sortThumbnailIndicesByProximity(
        remainingThumbnailIndices(slideCount, this.renderedThumbnailIndices),
        this.host.currentSlide,
      );
      if (remaining.length === 0) return;

      const batch = remaining.slice(0, THUMBNAIL_IDLE_BATCH_SIZE);
		void this.renderThumbnailBatch(batch, generation)
			.then(() => {
				if (generation !== this.thumbnailRenderGeneration) return;
				if (remaining.length > batch.length) {
					this.cancelIdleThumbnailFill = scheduleIdleWork(fillBatch, { timeout: 3000 });
				}
			})
			.catch((error) => {
				errorLog('render', 'PowerPoint idle thumbnail render failed', {
					error: cleanError(error),
				});
			});
    };

    this.cancelIdleThumbnailFill = scheduleIdleWork(fillBatch, { timeout: 1500 });
  }

  private registerThumbnailDrag(item: HTMLElement, index: number): void {
    item.draggable = this.host.canEdit();
    item.dataset.slideIndex = String(index);

    item.addEventListener('dragstart', (event) => {
      if (!this.host.canEdit()) {
        event.preventDefault();
        return;
      }
      this.thumbnailDragIndex = index;
      item.addClass('is-dragging');
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(index));
      }
    });

    item.addEventListener('dragend', () => {
      this.thumbnailDragIndex = null;
      this.clearThumbnailDropIndicators();
      item.removeClass('is-dragging');
    });

    item.addEventListener('dragover', (event) => {
      if (this.thumbnailDragIndex === null) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
      const after = this.isPointerInLowerHalf(event, item);
      this.clearThumbnailDropIndicators();
      item.addClass(after ? 'drop-after' : 'drop-before');
    });

    item.addEventListener('dragleave', () => {
      item.removeClass('drop-before');
      item.removeClass('drop-after');
    });

    item.addEventListener('drop', (event) => {
      event.preventDefault();
      const fromIndex = this.thumbnailDragIndex;
      this.thumbnailDragIndex = null;
      this.clearThumbnailDropIndicators();
      if (fromIndex === null) return;

      const after = this.isPointerInLowerHalf(event, item);
      let toIndex = after ? index + 1 : index;
      if (fromIndex < toIndex) toIndex -= 1;
      void this.reorderSlideByDrag(fromIndex, toIndex);
    });
  }

  private isPointerInLowerHalf(event: DragEvent, item: HTMLElement): boolean {
    const rect = item.getBoundingClientRect();
    return event.clientY > rect.top + rect.height / 2;
  }

  private clearThumbnailDropIndicators(): void {
    this.host.thumbnailContainer?.querySelectorAll('.drop-before, .drop-after').forEach((element) => {
      element.classList.remove('drop-before', 'drop-after');
    });
  }

  navigateToSlide(index: number, reason: string): void {
    const run = () => this.goToSlide(index, reason);
	this.slideNavigationPromise = this.slideNavigationPromise
		.then(run, run)
		.catch((error) => {
			errorLog('slide', 'PowerPoint slide navigation failed', {
				index,
				reason,
				error: cleanError(error),
			});
		});
	void this.slideNavigationPromise;
  }

  private updateThumbnailActiveState(): void {
    if (!this.host.thumbnailContainer) return;

    const items = this.host.thumbnailContainer.querySelectorAll('.native-powerpoint-thumbnail');
    items.forEach((item, index) => {
      item.toggleClass('active', index === this.host.currentSlide);
      item.toggleClass('is-selected', this.selectedSlideIndices.has(index));
    });
  }

  private async goToSlide(index: number, reason = 'unknown'): Promise<void> {
    if (!this.host.engine || this.host.isLoading) return;
    if (index < 0 || index >= this.host.engine.slideCount) return;

    const fromSlide = this.host.currentSlide;
    const focusesFilmstrip =
      reason === 'thumbnail-click'
      || reason.startsWith('keyboard-');
    if (focusesFilmstrip) {
      this.host.lastInteractionRegion = 'thumbnails';
    }

    if (index === fromSlide) {
      // Re-clicking the active thumbnail must still drop canvas selection so
      // Delete targets the slide (not a stale shape index set).
      await this.host.finishInlineTextEditing(`slide-navigation:${reason}`);
      this.host.clearSelection({ skipTextCommit: true });
      this.updateThumbnailActiveState();
      debugLog('slide', 'goToSlide skipped (already active)', {
        index,
        reason,
        clearedShapeSelection: true,
        interactionRegion: this.host.lastInteractionRegion,
      });
      return;
    }

    const generation = ++this.host.slideRenderGeneration;
    const navigationStarted = performance.now();
    this.host.isNavigatingSlide = true;
    debugLog('slide', 'goToSlide start', { from: fromSlide, to: index, reason, generation });

    try {
      await this.host.finishInlineTextEditing(`slide-navigation:${reason}`);
      if (generation !== this.host.slideRenderGeneration) {
        debugLog('slide', 'goToSlide aborted (superseded)', { from: fromSlide, to: index, generation, reason });
        return;
      }

      this.host.currentSlide = index;
      // Clear the full multi-select set — nulling only selectedShapeIndex left
      // selectedShapeIndices populated and made Delete a no-op.
      this.host.clearSelection();
      const rendered = await this.host.renderCurrentSlide(false, generation);
      if (generation !== this.host.slideRenderGeneration) {
        debugLog('slide', 'goToSlide render discarded (superseded)', { index, generation, reason });
        return;
      }

      if (rendered) {
        this.updateThumbnailActiveState();
        if (this.host.engine && shouldUseLazyThumbnails(this.host.engine.slideCount)) {
          this.scheduleThumbnailRefresh(
            priorityThumbnailIndices(index, this.host.engine.slideCount),
          );
        }
        this.host.renderInspector();
      }

      const navigationMs = Math.round(performance.now() - navigationStarted);
      debugLog('slide', 'goToSlide complete', { index, reason, generation, ms: navigationMs });
      if (navigationMs > 1000) {
        warnLog('slide', 'slow goToSlide', { from: fromSlide, to: index, reason, generation, ms: navigationMs });
      }
    } catch (error) {
      errorLog('slide', 'goToSlide failed', { from: fromSlide, to: index, reason, generation, error });
      this.notice('powerpoint:notice.couldNotOpenSlide', { slideNumber: index + 1, message: cleanError(error) });
    } finally {
      if (generation === this.host.slideRenderGeneration) {
        this.host.isNavigatingSlide = false;
      }
    }
  }

  async addSlide(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('add slide')) return;

    try {
      debugLog('slide', 'Add slide started', { afterSlide: this.host.currentSlide });
      const history = await this.host.captureHistoryEntry('Add slide');
      const result = await this.host.engine.addSlide(this.host.currentSlide);
      this.host.currentSlide = result.slideIndex;
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
      debugLog('slide', 'Add slide completed', { slide: this.host.currentSlide, slideCount: this.host.engine.slideCount });
    } catch (error) {
      errorLog('slide', 'Add slide failed', { slide: this.host.currentSlide, error });
      this.notice('powerpoint:notice.couldNotAddSlide', { message: cleanError(error) });
    }
  }

  async deleteSlide(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('delete slide')) return;

    try {
      debugLog('slide', 'Delete slide started', { slide: this.host.currentSlide });
      const history = await this.host.captureHistoryEntry('Delete slide');
      const result = await this.host.engine.deleteSlide(this.host.currentSlide);
      this.host.currentSlide = result.slideIndex;
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Delete slide completed', { slide: this.host.currentSlide, slideCount: this.host.engine.slideCount });
    } catch (error) {
      errorLog('slide', 'Delete slide failed', { slide: this.host.currentSlide, error });
      this.notice('powerpoint:notice.couldNotDeleteSlide', { message: cleanError(error) });
    }
  }

  async deleteSelectedSlides(): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('delete slides')) return;

    const targets = Array.from(this.selectedSlideIndices).sort((a, b) => a - b);
    if (targets.length === 0) return;
    if (targets.length >= this.host.engine.slideCount) {
      this.notice('powerpoint:notice.cannotDeleteEverySlide');
      return;
    }

    try {
      debugLog('slide', 'Delete selected slides started', { targets, slideCount: this.host.engine.slideCount });
      const history = await this.host.captureHistoryEntry('Delete slides');
      const result = await this.host.engine.deleteSlides(targets);
      this.host.currentSlide = result.slideIndex;
      this.clearSlideSelection();
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Delete selected slides completed', {
        deletedCount: targets.length,
        slide: this.host.currentSlide,
        slideCount: this.host.engine.slideCount
      });
    } catch (error) {
      errorLog('slide', 'Delete selected slides failed', { targets, error });
      this.notice('powerpoint:notice.couldNotDeleteSlides', { message: cleanError(error) });
    }
  }

  async moveSlide(direction: -1 | 1): Promise<void> {
    await this.moveSlideAt(this.host.currentSlide, direction);
  }

  async moveSlideAt(index: number, direction: -1 | 1): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('move slide')) return;
    if (index < 0 || index >= this.host.engine.slideCount) return;

    try {
      debugLog('slide', 'Move slide started', { index, direction });
      const history = await this.host.captureHistoryEntry('Move slide');
      const result = await this.host.engine.moveSlide(index, direction);
      if (result.slideIndex === index) return;

      this.host.currentSlide = result.slideIndex;
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Move slide completed', { from: index, to: this.host.currentSlide });
    } catch (error) {
      errorLog('slide', 'Move slide failed', { index, direction, error });
      this.notice('powerpoint:notice.couldNotMoveSlide', { message: cleanError(error) });
    }
  }

  async addSlideWithLayout(layout: SlideLayoutKind): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('add slide')) return;

    try {
      debugLog('slide', 'Add slide with layout started', { afterSlide: this.host.currentSlide, layout });
      const history = await this.host.captureHistoryEntry('New slide');
      const result = await this.host.engine.addSlideWithLayout(this.host.currentSlide, layout);
      this.host.currentSlide = result.slideIndex;
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails();
      debugLog('slide', 'Add slide with layout completed', {
        slide: this.host.currentSlide,
        slideCount: this.host.engine.slideCount,
        layout
      });
    } catch (error) {
      errorLog('slide', 'Add slide with layout failed', { layout, error });
      this.notice('powerpoint:notice.couldNotAddSlide', { message: cleanError(error) });
    }
  }

  async duplicateSlide(targetIndex: number = this.host.currentSlide): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('duplicate slide')) return;
    if (targetIndex < 0 || targetIndex >= this.host.engine.slideCount) return;

    try {
      debugLog('slide', 'Duplicate slide started', { targetIndex });
      const history = await this.host.captureHistoryEntry('Duplicate slide');
      const result = await this.host.engine.duplicateSlide(targetIndex);
      this.host.currentSlide = result.slideIndex;
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Duplicate slide completed', {
        sourceSlide: targetIndex,
        slide: this.host.currentSlide,
        slideCount: this.host.engine.slideCount
      });
    } catch (error) {
      errorLog('slide', 'Duplicate slide failed', { targetIndex, error });
      this.notice('powerpoint:notice.couldNotDuplicateSlide', { message: cleanError(error) });
    }
  }

  async deleteSlideAt(targetIndex: number): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('delete slide')) return;
    if (targetIndex < 0 || targetIndex >= this.host.engine.slideCount) return;

    try {
      debugLog('slide', 'Delete slide at index started', { targetIndex });
      const history = await this.host.captureHistoryEntry('Delete slide');
      const result = await this.host.engine.deleteSlide(targetIndex);
      this.host.currentSlide = result.slideIndex;
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Delete slide at index completed', {
        deletedSlide: targetIndex,
        slide: this.host.currentSlide,
        slideCount: this.host.engine.slideCount
      });
    } catch (error) {
      errorLog('slide', 'Delete slide at index failed', { targetIndex, error });
      this.notice('powerpoint:notice.couldNotDeleteSlide', { message: cleanError(error) });
    }
  }

  async reorderSlideByDrag(fromIndex: number, toIndex: number): Promise<void> {
    if (!this.host.engine) return;
    if (!this.host.ensureEditable('reorder slides')) return;

    const slideCount = this.host.engine.slideCount;
    if (
      fromIndex === toIndex ||
      fromIndex < 0 ||
      fromIndex >= slideCount ||
      toIndex < 0 ||
      toIndex >= slideCount
    ) {
      return;
    }

    const order = Array.from({ length: slideCount }, (_, index) => index);
    const [moved] = order.splice(fromIndex, 1);
    if (moved === undefined) return;
    order.splice(toIndex, 0, moved);

    try {
      debugLog('slide', 'Reorder slides started', { fromIndex, toIndex, slideCount });
      const history = await this.host.captureHistoryEntry('Reorder slides');
      await this.host.engine.reorderSlides(order);
      this.host.currentSlide = toIndex;
      this.host.clearSelection();
      this.host.recordHistoryEntry(history);
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.renderThumbnails({ preferLazy: true });
      debugLog('slide', 'Reorder slides completed', { fromIndex, toIndex });
    } catch (error) {
      errorLog('slide', 'Reorder slides failed', { fromIndex, toIndex, error });
      this.notice('powerpoint:notice.couldNotReorderSlides', { message: cleanError(error) });
    }
  }

  showSlideContextMenu(event: MouseEvent, index: number): void {
    if (!this.host.engine) return;

    const menu = this.host.createNativeMenu();
    menu.addItem((item) =>
      item
        .setTitle('New slide')
        .setIcon('plus')
        .onClick(() => {
          this.host.currentSlide = index;
          void this.addSlideWithLayout('blank');
        })
    );
    menu.addItem((item) =>
      item
        .setTitle('Duplicate slide')
        .setIcon('files')
        .onClick(() => void this.duplicateSlide(index))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Move up')
        .setIcon('arrow-up')
        .setDisabled(index <= 0)
        .onClick(() => void this.moveSlideAt(index, -1))
    );
    menu.addItem((item) =>
      item
        .setTitle('Move down')
        .setIcon('arrow-down')
        .setDisabled(!this.host.engine || index >= this.host.engine.slideCount - 1)
        .onClick(() => void this.moveSlideAt(index, 1))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Delete slide')
        .setIcon('trash-2')
        .setDisabled(!this.host.engine || this.host.engine.slideCount <= 1)
        .onClick(() => void this.deleteSlideAt(index))
    );
    menu.showAtMouseEvent(event);
  }

  selectAllSlides(): void {
    if (!this.host.engine) return;
    const count = this.host.engine.slideCount;
    if (count === 0) return;
    this.selectedSlideIndices.clear();
    for (let index = 0; index < count; index += 1) {
      this.selectedSlideIndices.add(index);
    }
    this.applySlideSelectionClasses();
  }

  clearSlideSelection(): void {
    if (this.selectedSlideIndices.size === 0) return;
    this.selectedSlideIndices.clear();
    this.applySlideSelectionClasses();
  }

  private toggleSlideSelection(index: number): void {
    if (this.selectedSlideIndices.has(index)) {
      this.selectedSlideIndices.delete(index);
    } else {
      this.selectedSlideIndices.add(index);
    }
    this.applySlideSelectionClasses();
  }

  private selectSlideRange(anchor: number, index: number): void {
    const start = Math.min(anchor, index);
    const end = Math.max(anchor, index);
    this.selectedSlideIndices.clear();
    for (let slide = start; slide <= end; slide += 1) {
      this.selectedSlideIndices.add(slide);
    }
    this.applySlideSelectionClasses();
  }

  private applySlideSelectionClasses(): void {
    this.host.thumbnailContainer?.querySelectorAll('.native-powerpoint-thumbnail').forEach((thumbnail, index) => {
      thumbnail.classList.toggle('is-selected', this.selectedSlideIndices.has(index));
    });
  }
}
