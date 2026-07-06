import { Platform } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import type { PresentationEngine } from '../PresentationEngine';
import { HISTORY_LIMIT } from './constants';
import { debugLog, errorLog } from '../logger';
import { cleanError } from './runtimeCompat';
import type { HistoryEntry } from './types';

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
  ensureEditable(action: string): boolean;
  canEdit(): boolean;
  clearAutosave(): void;
  clearDragState(): void;
  clearSelection(): void;
  markDirty(): void;
  renderCurrentSlide(keepSelection?: boolean): Promise<boolean>;
  renderThumbnails(): Promise<void>;
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
    this.undoButton = undoButton;
    this.redoButton = redoButton;
  }

  async capture(label: string): Promise<HistoryEntry> {
    if (!this.host.engine) {
      throw new Error('Open a loaded PowerPoint file first.');
    }

    const buffer = await this.host.engine.export();
    debugLog('history', 'Captured PowerPoint history snapshot', {
      label,
      slide: this.host.currentSlide,
      bytes: buffer.byteLength
    });
    return {
      buffer,
      currentSlide: this.host.currentSlide,
      label
    };
  }

  record(entry: HistoryEntry): void {
    this.undoStack.push(entry);
    if (this.undoStack.length > HISTORY_LIMIT) {
      this.undoStack.shift();
    }
    this.redoStack = [];
    debugLog('history', 'Recorded PowerPoint history entry', {
      label: entry.label,
      undoDepth: this.undoStack.length,
      redoDepth: this.redoStack.length
    });
    this.updateAvailability();
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.isRestoringHistory = false;
    debugLog('history', 'Cleared PowerPoint history');
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
    if (!this.host.engine || this.isRestoringHistory || source.length === 0) return;
    if (!this.host.ensureEditable(action)) return;
    if (this.host.activeEditor) {
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
      label: entry.label,
      sourceDepth: source.length,
      destinationDepth: destination.length
    });

    try {
      const current = await this.capture(entry.label);
      await this.host.engine.restoreSnapshot(entry.buffer);
      source.pop();
      destination.push(current);
      if (destination.length > HISTORY_LIMIT) {
        destination.shift();
      }

      this.host.currentSlide = Math.max(0, Math.min(entry.currentSlide, this.host.engine.slideCount - 1));
      this.host.clearSelection();
      this.host.markDirty();
      const rendered = await this.host.renderCurrentSlide();
      if (rendered) await this.host.renderThumbnails();
      debugLog('history', `PowerPoint ${action} completed`, {
        label: entry.label,
        slide: this.host.currentSlide,
        sourceDepth: source.length,
        destinationDepth: destination.length
      });
    } catch (error) {
      errorLog('history', `PowerPoint ${action} failed`, {
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
    this.updateHistoryButton(this.undoButton, 'Undo', `${modifier}+Z`, canUseHistory && this.undoStack.length > 0);
    this.updateHistoryButton(this.redoButton, 'Redo', `${modifier}+Shift+Z`, canUseHistory && this.redoStack.length > 0);
  }

  private updateHistoryButton(
    button: HTMLButtonElement | null,
    label: string,
    shortcut: string,
    enabled: boolean
  ): void {
    if (!button) return;

    const nextEntry = label === 'Undo'
      ? this.undoStack[this.undoStack.length - 1]
      : this.redoStack[this.redoStack.length - 1];
    button.disabled = !enabled;
    button.toggleClass('is-disabled', !enabled);
    button.setAttribute('aria-disabled', String(!enabled));
    button.setAttribute('aria-label', nextEntry ? `${label} ${nextEntry.label} (${shortcut})` : `${label} (${shortcut})`);
  }
}
