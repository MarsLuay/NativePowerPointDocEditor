import { Platform } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import type { PresentationEngine } from '../PresentationEngine';
import { HISTORY_LIMIT } from './constants';
import { debugLog, errorLog } from '../logger';
import { cleanError } from './runtimeCompat';
import type { HistoryEntry, HistorySlideXmlEntry, HistoryTransformChange } from './types';

/**
 * The slice of `NativePowerPointView` that the undo/redo subsystem reaches back
 * into. As with {@link FindReplaceHost}, this is the regression boundary for
 * the extraction; the view supplies it via an adapter object so its own members
 * can stay `private`.
 */
export interface HistoryHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  currentSlide: number;
  readonly activeEditor: HTMLTextAreaElement | null;
  /** Active in-place text edits keep a short, separate undo stack. */
  canUndoInlineEdit?(): boolean;
  canRedoInlineEdit?(): boolean;
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  clearAutosave(): void;
  clearDragState(): void;
  clearSelection(): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderThumbnails(): Promise<void>;
  scheduleThumbnailRefresh(indices: number | number[]): void;
}

/**
 * Owns the undo/redo stacks and the snapshot capture/restore flow. Extracted
 * verbatim from `NativePowerPointView`; it borrows shared editor state through
 * {@link HistoryHost}.
 *
 * The view keeps thin `captureHistoryEntry`/`recordHistoryEntry` forwarders so
 * its ~30 mutation call sites are unchanged.
 */
export class HistoryController {
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private isRestoringHistory = false;

  private undoButton: HTMLButtonElement | null = null;
  private redoButton: HTMLButtonElement | null = null;
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: HistoryHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  /** True while a restore is in flight; mutations and history should be inert. */
  get isRestoring(): boolean {
    return this.isRestoringHistory;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Registers the toolbar buttons so availability/labels can be kept in sync. */
  attachButtons(undoButton: HTMLButtonElement, redoButton: HTMLButtonElement): void {
    debugLog('history', 'PowerPoint history button attachment started', { op: 'attach-buttons' });
    this.undoButton = undoButton;
    this.redoButton = redoButton;
  }

  async capture(label: string): Promise<HistoryEntry> {
    if (!this.host.engine) {
      throw new Error('Open a loaded PowerPoint file first.');
    }

    debugLog('history', 'PowerPoint history snapshot capture started', {
      op: 'capture-snapshot',
      label,
      slide: this.host.currentSlide
    });
    try {
      const buffer = await this.host.engine.snapshotAuthoritativePackage();
      debugLog('history', 'Captured PowerPoint history snapshot', {
        op: 'capture-snapshot',
        label,
        slide: this.host.currentSlide,
        bytes: buffer.byteLength
      });
      return {
        kind: 'snapshot',
        buffer,
        currentSlide: this.host.currentSlide,
        label
      };
    } catch (error) {
      errorLog('history', 'PowerPoint history snapshot capture failed', {
        op: 'capture-snapshot',
        label,
        error
      });
      throw error;
    }
  }

  /**
   * Records an object transform as a before/after delta instead of a full-deck
   * snapshot. This keeps drag/resize commits off the expensive
   * `engine.export()` path, which is the main cause of move lag on large decks.
   */
  captureTransform(slideIndex: number, changes: HistoryTransformChange[], label: string): HistoryEntry {
    debugLog('history', 'PowerPoint transform history capture started', {
      op: 'capture-transform',
      label,
      slide: slideIndex,
      changeCount: changes.length
    });
    return {
      kind: 'transform',
      slideIndex,
      changes: changes.map((change) => ({ ...change })),
      currentSlide: this.host.currentSlide,
      label
    };
  }

  captureSlideXml(slideIndex: number, label: string): HistorySlideXmlEntry {
    if (!this.host.engine) {
      throw new Error('Open a loaded PowerPoint file first.');
    }

    debugLog('history', 'PowerPoint slide XML history capture started', {
      op: 'capture-slide-xml',
      label,
      slide: slideIndex
    });
    try {
      const beforeXml = this.host.engine.getSlideXml(slideIndex);
      debugLog('history', 'Captured PowerPoint slide XML history base', {
        op: 'capture-slide-xml',
        label,
        slide: slideIndex,
        chars: beforeXml.length
      });
      return {
        kind: 'slideXml',
        slideIndex,
        beforeXml,
        afterXml: beforeXml,
        currentSlide: this.host.currentSlide,
        label
      };
    } catch (error) {
      errorLog('history', 'PowerPoint slide XML history capture failed', {
        op: 'capture-slide-xml',
        label,
        slide: slideIndex,
        error
      });
      throw error;
    }
  }

  completeSlideXml(entry: HistorySlideXmlEntry): HistorySlideXmlEntry {
    if (!this.host.engine) {
      throw new Error('Open a loaded PowerPoint file first.');
    }

    debugLog('history', 'PowerPoint slide XML history completion started', {
      op: 'complete-slide-xml',
      label: entry.label,
      slide: entry.slideIndex
    });
    try {
      const afterXml = this.host.engine.getSlideXml(entry.slideIndex);
      debugLog('history', 'Captured PowerPoint slide XML history result', {
        op: 'complete-slide-xml',
        label: entry.label,
        slide: entry.slideIndex,
        beforeChars: entry.beforeXml.length,
        afterChars: afterXml.length,
        changed: entry.beforeXml !== afterXml
      });
      return {
        ...entry,
        afterXml
      };
    } catch (error) {
      errorLog('history', 'PowerPoint slide XML history completion failed', {
        op: 'complete-slide-xml',
        label: entry.label,
        slide: entry.slideIndex,
        error
      });
      throw error;
    }
  }

  record(entry: HistoryEntry): void {
    debugLog('history', 'PowerPoint history entry record started', {
      op: 'record',
      label: entry.label
    });
    this.undoStack.push(entry);
    if (this.undoStack.length > HISTORY_LIMIT) {
      this.undoStack.shift();
    }
    this.redoStack = [];
    debugLog('history', 'Recorded PowerPoint history entry', {
      op: 'record',
      label: entry.label,
      undoDepth: this.undoStack.length,
      redoDepth: this.redoStack.length
    });
    this.updateAvailability();
  }

  clear(): void {
    debugLog('history', 'PowerPoint history clear started', { op: 'clear' });
    this.undoStack = [];
    this.redoStack = [];
    this.isRestoringHistory = false;
    debugLog('history', 'Cleared PowerPoint history', { op: 'clear' });
    this.updateAvailability();
  }

  async undo(): Promise<void> {
    await this.restoreHistoryEntry(this.undoStack, this.redoStack, 'undo');
  }

  async redo(): Promise<void> {
    await this.restoreHistoryEntry(this.redoStack, this.undoStack, 'redo');
  }

  private async restoreHistoryEntry(
    source: HistoryEntry[],
    destination: HistoryEntry[],
    action: 'undo' | 'redo'
  ): Promise<void> {
    if (!this.host.engine) {
      debugLog('history', `PowerPoint ${action} skipped`, { op: action, reason: 'no-engine' });
      return;
    }
    if (this.isRestoringHistory) {
      debugLog('history', `PowerPoint ${action} skipped`, { op: action, reason: 'restoring' });
      return;
    }
    if (source.length === 0) {
      debugLog('history', `PowerPoint ${action} skipped`, { op: action, reason: 'empty-stack' });
      return;
    }
    if (!this.host.ensureEditable(action)) {
      debugLog('history', `PowerPoint ${action} skipped`, { op: action, reason: 'not-editable' });
      return;
    }
    if (this.host.activeEditor) {
      debugLog('history', `PowerPoint ${action} skipped`, {
        op: action,
        reason: 'active-inline-editor',
        sourceDepth: source.length,
        destinationDepth: destination.length,
      });
      this.host.activeEditor.blur();
      return;
    }

    const entry = source[source.length - 1];
    if (!entry) return;

    this.isRestoringHistory = true;
    this.host.clearAutosave();
    this.host.clearDragState();
    this.updateAvailability();
    debugLog('history', `PowerPoint ${action} started`, {
      op: action,
      label: entry.label,
      sourceDepth: source.length,
      destinationDepth: destination.length
    });

    try {
      if (entry.kind === 'transform') {
        for (const change of entry.changes) {
          await this.host.engine.updateShapeTransform(
            entry.slideIndex,
            change.shapeIndex,
            action === 'undo' ? change.before : change.after
          );
        }
        source.pop();
        destination.push(entry);
        if (destination.length > HISTORY_LIMIT) {
          destination.shift();
        }
        this.host.currentSlide = Math.max(0, Math.min(entry.slideIndex, this.host.engine.slideCount - 1));
      } else if (entry.kind === 'slideXml') {
        await this.host.engine.restoreSlideXml(
          entry.slideIndex,
          action === 'undo' ? entry.beforeXml : entry.afterXml
        );
        source.pop();
        destination.push(entry);
        if (destination.length > HISTORY_LIMIT) {
          destination.shift();
        }
        this.host.currentSlide = Math.max(0, Math.min(entry.slideIndex, this.host.engine.slideCount - 1));
      } else {
        const current = await this.capture(entry.label);
        await this.host.engine.restoreSnapshot(entry.buffer);
        source.pop();
        destination.push(current);
        if (destination.length > HISTORY_LIMIT) {
          destination.shift();
        }
        this.host.currentSlide = Math.max(0, Math.min(entry.currentSlide, this.host.engine.slideCount - 1));
      }

      this.host.clearSelection();
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) {
        if (entry.kind === 'transform' || entry.kind === 'slideXml') {
          this.host.scheduleThumbnailRefresh(entry.slideIndex);
        } else {
          await this.host.renderThumbnails();
        }
      }
      debugLog('history', `PowerPoint ${action} completed`, {
        op: action,
        label: entry.label,
        slide: this.host.currentSlide,
        sourceDepth: source.length,
        destinationDepth: destination.length
      });
    } catch (error) {
      errorLog('history', `PowerPoint ${action} failed`, {
        op: action,
        label: entry.label,
        error
      });
      this.notice('powerpoint:notice.couldNotHistoryAction', { action, message: cleanError(error) });
    } finally {
      this.isRestoringHistory = false;
      this.updateAvailability();
    }
  }

  updateAvailability(): void {
    const canUseHistory = this.host.canEdit() && !this.isRestoringHistory;
    const modifier = Platform?.isMacOS === true ? 'Cmd' : 'Ctrl';
    const hasInlineUndo = this.host.canUndoInlineEdit?.() === true;
    const hasInlineRedo = this.host.canRedoInlineEdit?.() === true;
    this.updateHistoryButton(
      this.undoButton,
      'Undo',
      `${modifier}+Z`,
      canUseHistory && (hasInlineUndo || this.undoStack.length > 0),
      hasInlineUndo,
    );
    this.updateHistoryButton(
      this.redoButton,
      'Redo',
      `${modifier}+Shift+Z`,
      canUseHistory && (hasInlineRedo || this.redoStack.length > 0),
      hasInlineRedo,
    );
  }

  private updateHistoryButton(
    button: HTMLButtonElement | null,
    label: string,
    shortcut: string,
    enabled: boolean,
    usesInlineHistory = false,
  ): void {
    if (!button) return;

    const nextEntry = usesInlineHistory ? null : label === 'Undo'
      ? this.undoStack[this.undoStack.length - 1]
      : this.redoStack[this.redoStack.length - 1];
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
    button.setAttribute('aria-label', nextEntry ? `${label} ${nextEntry.label} (${shortcut})` : `${label} (${shortcut})`);
  }
}
