import { SaveController, type SaveHost, type SaveStateStore } from '../saveController';
import type { SaveState } from '../types';
import { debugLog } from '../../logger';
import type { PptxCommand } from '../commands/types';

export interface PresentationSelectionSnapshot {
  readonly shapeIndexes: readonly number[];
}

export interface PresentationSessionSnapshot {
  readonly currentSlide: number;
  readonly selection: PresentationSelectionSnapshot;
  readonly dirty: boolean;
  readonly editVersion: number;
  readonly saveState: SaveState;
}

export type PresentationSessionEvent =
  | { type: 'selection'; snapshot: PresentationSessionSnapshot }
  | { type: 'slide'; snapshot: PresentationSessionSnapshot }
  | { type: 'save'; snapshot: PresentationSessionSnapshot }
  | { type: 'history'; action: 'undo' | 'redo'; snapshot: PresentationSessionSnapshot }
  | { type: 'command'; command: PptxCommand; snapshot: PresentationSessionSnapshot };

export type PresentationSessionListener = (event: PresentationSessionEvent) => void;

/** History boundary so the session can own its public undo/redo API. */
export interface HistoryPort {
  undo(): boolean | void;
  redo(): boolean | void;
}

/** Phase-2 mutation boundary for applying non-noop presentation commands. */
export interface MutationExecutor {
  execute(command: Exclude<PptxCommand, { type: 'noop' }>): Promise<unknown>;
}

export interface PresentationSessionOptions {
  readonly history?: HistoryPort;
  readonly mutationExecutor?: MutationExecutor;
}

const NOOP_HISTORY_PORT: HistoryPort = {
  undo: () => false,
  redo: () => false
};

/**
 * Owns presentation-wide state that is shared by controllers and the view.
 * Engine/package extraction intentionally remains separate.
 */
export class PresentationSession implements SaveStateStore {
  private currentSlideIndex = 0;
  private selectedShapeIndexes: readonly number[] = [];
  private isDirty = false;
  private version = 0;
  private saveStatus: SaveState = 'idle';
  private readonly listeners = new Set<PresentationSessionListener>();
  private readonly history: HistoryPort;
  private readonly mutationExecutor?: MutationExecutor;
  readonly saveController: SaveController;

  constructor(saveHost: SaveHost, options: PresentationSessionOptions = {}) {
    this.saveController = new SaveController(saveHost, this);
    this.history = options.history ?? NOOP_HISTORY_PORT;
    this.mutationExecutor = options.mutationExecutor;
  }

  get currentSlide(): number {
    return this.currentSlideIndex;
  }

  setCurrentSlide(index: number): void {
    if (index === this.currentSlideIndex) return;
    this.currentSlideIndex = index;
    this.emit('slide');
  }

  get selection(): PresentationSelectionSnapshot {
    return { shapeIndexes: [...this.selectedShapeIndexes] };
  }

  selectShapes(shapeIndexes: readonly number[]): void {
    const next = [...new Set(shapeIndexes)];
    if (next.some((index) => !Number.isSafeInteger(index) || index < 0)) {
      throw new RangeError('Selected shape indexes must be non-negative safe integers.');
    }
    if (this.selectedShapeIndexes.length === next.length
      && this.selectedShapeIndexes.every((index, position) => index === next[position])) {
      return;
    }
    this.selectedShapeIndexes = next;
    this.emit('selection');
  }

  clearSelection(): void {
    if (this.selectedShapeIndexes.length === 0) return;
    this.selectedShapeIndexes = [];
    this.emit('selection');
  }

  get dirty(): boolean {
    return this.isDirty;
  }

  get editVersion(): number {
    return this.version;
  }

  get saveState(): SaveState {
    return this.saveStatus;
  }

  async save(source: 'manual' | 'autosave' = 'manual'): Promise<boolean> {
    return this.saveController.save(source);
  }

  setSaveStatus(state: SaveState): void {
    this.saveController.setState(state);
  }

  clearAutosave(): void {
    this.saveController.clearAutosave();
  }

  async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    return this.saveController.preserveUnsavedChangesForTeardown(reason);
  }

  /** Records a document mutation and lets the save controller schedule persistence. */
  applyEdit(): void {
    this.saveController.markDirty();
  }

  undo(): boolean {
    const applied = this.history.undo() !== false;
    this.emitHistory('undo');
    return applied;
  }

  redo(): boolean {
    const applied = this.history.redo() !== false;
    this.emitHistory('redo');
    return applied;
  }

  /** Applies a mutation transaction and marks the session dirty only on commit. */
  async applyCommand(command: PptxCommand): Promise<unknown> {
    this.emitCommand(command);
    if (command.type === 'noop') return;
    if (!this.mutationExecutor) {
      debugLog('mutate', 'PowerPoint command recorded without mutation executor', {
        operation: command.type
      });
      return;
    }
    const result = await this.mutationExecutor.execute(command);
    const unchangedRunFormatting = result === false && (
      command.type === 'set-run-style'
      || command.type === 'set-run-style-range'
      || command.type === 'set-run-style-ranges'
    );
    if (unchangedRunFormatting) {
      debugLog('mutate', 'PowerPoint run formatting made no document change', {
        operation: command.type,
      });
      return result;
    }
    if (command.type === 'reorder-shapes' && result === null) {
      debugLog('mutate', 'PowerPoint overlap-aware reorder made no structural change', {
        slide: command.slideIndex,
        shapeIndexes: command.shapeIndexes,
        mode: command.mode,
      });
      return result;
    }
    this.applyEdit();
    return result;
  }

  markDirty(): void {
    this.isDirty = true;
    this.version += 1;
    this.setSaveState('dirty');
  }

  clearDirty(): void {
    if (!this.isDirty) return;
    this.isDirty = false;
    this.emit('save');
  }

  /** Test-only: set edit version without marking dirty side effects. */
  setEditVersionForTests(value: number): void {
    this.version = value;
  }

  setSaveState(state: SaveState): void {
    this.saveStatus = state;
    this.emit('save');
  }

  reset(): void {
    const hadSelection = this.selectedShapeIndexes.length > 0;
    this.saveController.reset();
    this.currentSlideIndex = 0;
    this.selectedShapeIndexes = [];
    this.isDirty = false;
    this.version = 0;
    this.saveStatus = 'idle';
    if (hadSelection) this.emit('selection');
    this.emit('slide');
    this.emit('save');
  }

  subscribe(listener: PresentationSessionListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(type: 'selection' | 'slide' | 'save'): void {
    const snapshot = this.snapshot();
    for (const listener of this.listeners) listener({ type, snapshot });
  }

  private emitHistory(action: 'undo' | 'redo'): void {
    const snapshot = this.snapshot();
    for (const listener of this.listeners) listener({ type: 'history', action, snapshot });
  }

  private emitCommand(command: PptxCommand): void {
    const snapshot = this.snapshot();
    for (const listener of this.listeners) listener({ type: 'command', command, snapshot });
  }

  snapshot(): PresentationSessionSnapshot {
    return {
      currentSlide: this.currentSlideIndex,
      selection: this.selection,
      dirty: this.isDirty,
      editVersion: this.version,
      saveState: this.saveStatus
    };
  }
}
