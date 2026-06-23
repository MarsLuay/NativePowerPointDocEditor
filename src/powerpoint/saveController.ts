import { type App, Notice, TFile, normalizePath } from 'obsidian';

import {
  inspectPowerPointPackage,
  summarizePackageMessages,
  validatePowerPointExport,
  validatePowerPointExportContents,
  type PowerPointPackageInspection
} from '../PowerPointPackage';
import { PresentationEngine } from '../PresentationEngine';
import {
  isEditablePowerPointExtension,
  isModernPowerPointExtension
} from './extensions';
import { cleanError, flushUi } from './runtimeCompat';
import type { SaveState } from './types';

export interface SaveStatusElements {
  statusEl: HTMLElement;
  statusLabelEl: HTMLElement;
  statusProgressEl: HTMLElement | null;
}

/**
 * The slice of `NativePowerPointView` that the save/autosave/recovery subsystem
 * reaches back into. The view supplies it via an adapter object so its own
 * members can stay `private`.
 */
export interface SaveHost {
  readonly app: App;
  readonly engine: PresentationEngine | null;
  readonly loadedFile: TFile | null;
  readonly file: TFile | null;
  sourcePackage: PowerPointPackageInspection | null;
  sourceBuffer: ArrayBuffer | null;
  readonly viewOnlyReason: string;
  autosaveEnabled(): boolean;
  ensureEditable(action: string): boolean;
  getSaveStatusElements(): SaveStatusElements | null;
  isSameLoadedPresentation(engine: PresentationEngine, file: TFile): boolean;
  validateExportBeforeSave(
    output: ArrayBuffer,
    engine: PresentationEngine,
    sourcePackage: PowerPointPackageInspection,
    sourceBuffer: ArrayBuffer,
    onProgress?: (message: string) => void
  ): Promise<PowerPointPackageInspection>;
  saveCurrentPresentation(): Promise<boolean>;
}

/**
 * Owns save/autosave/recovery state and the phased export validation pipeline.
 * Extracted from `NativePowerPointView`; status DOM creation stays on the view.
 */
export class SaveController {
  saveState: SaveState = 'idle';
  isDirty = false;
  editVersion = 0;
  saveTimer: number | null = null;
  savePromise: Promise<void> = Promise.resolve();
  lastSaveError: string | null = null;

  constructor(private readonly host: SaveHost) {}

  shouldPaintSaveProgress(): boolean {
    return this.host.getSaveStatusElements()?.statusLabelEl !== null;
  }

  markDirty(): void {
    this.isDirty = true;
    this.editVersion++;
    this.setSaveState('dirty');
    if (this.host.autosaveEnabled()) {
      this.scheduleAutosave();
    }
  }

  scheduleAutosave(delayMs = 1500): void {
    this.clearAutosave();
    if (!this.host.autosaveEnabled()) return;

    this.saveTimer = window.setTimeout(() => {
      this.saveTimer = null;
      void this.host.saveCurrentPresentation();
    }, delayMs);
  }

  clearAutosave(): void {
    if (this.saveTimer !== null) {
      window.clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
  }

  resetLoadedState(): void {
    this.clearAutosave();
    this.isDirty = false;
    this.editVersion = 0;
    this.lastSaveError = null;
    this.savePromise = Promise.resolve();
  }

  async saveCurrentPresentation(): Promise<boolean> {
    const file = this.host.loadedFile || this.host.file;
    const engine = this.host.engine;
    const sourcePackage = this.host.sourcePackage;
    const sourceBuffer = this.host.sourceBuffer;
    if (!file || !engine || !sourcePackage || !sourceBuffer || !isModernPowerPointExtension(file.extension)) {
      new Notice('Open a modern PowerPoint file to save it.');
      return false;
    }

    if (!isEditablePowerPointExtension(file.extension)) {
      new Notice(this.host.viewOnlyReason || 'This PowerPoint format is view-only in Native PowerPoint.');
      return false;
    }

    if (!this.host.ensureEditable('save changes')) {
      return false;
    }

    const targetVersion = this.editVersion;
    this.clearAutosave();
    this.setSaveState('saving', 'Exporting...');

    const run = async () => {
      if (this.shouldPaintSaveProgress()) await flushUi();
      const output = await engine.export();
      const exportedPackage = await this.host.validateExportBeforeSave(
        output,
        engine,
        sourcePackage,
        sourceBuffer,
        (message) => this.setSaveProgress(message)
      );
      this.setSaveProgress('Writing to vault...');
      if (this.shouldPaintSaveProgress()) await flushUi();
      await this.host.app.vault.modifyBinary(file, output);

      if (this.host.isSameLoadedPresentation(engine, file)) {
        this.host.sourcePackage = exportedPackage;
        this.host.sourceBuffer = output;

        if (this.editVersion === targetVersion) {
          this.isDirty = false;
          this.lastSaveError = null;
          this.setSaveState('saved');
        } else {
          this.setSaveState('dirty');
          this.scheduleAutosave();
        }
      }
    };

    this.savePromise = this.savePromise.then(run, run);

    try {
      await this.savePromise;
      return true;
    } catch (error) {
      const message = cleanError(error);
      this.lastSaveError = message;
      this.setSaveState('failed');
      new Notice(`Could not save ${file.name}: ${message}`);
      if (this.isDirty && this.host.autosaveEnabled()) {
        this.scheduleAutosave(5000);
      }
      return false;
    }
  }

  async validateExportBeforeSave(
    output: ArrayBuffer,
    engine = this.host.engine,
    sourcePackage = this.host.sourcePackage,
    sourceBuffer = this.host.sourceBuffer,
    onProgress?: (message: string) => void
  ): Promise<PowerPointPackageInspection> {
    if (!engine || !sourcePackage || !sourceBuffer) {
      throw new Error('Cannot verify the PowerPoint package before saving.');
    }

    onProgress?.('Checking package...');
    if (this.shouldPaintSaveProgress()) await flushUi();
    const exportedPackage = inspectPowerPointPackage(output);
    const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount);
    if (!validation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);
    }

    onProgress?.('Verifying contents...');
    if (this.shouldPaintSaveProgress()) await flushUi();
    const contentValidation = await validatePowerPointExportContents(sourceBuffer, output);
    if (!contentValidation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(contentValidation.errors)}`);
    }

    onProgress?.('Round-trip check...');
    if (this.shouldPaintSaveProgress()) await flushUi();
    await PresentationEngine.validateRoundTrip(output, engine.slideCount);
    return exportedPackage;
  }

  async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    this.clearAutosave();
    await this.savePromise.catch(() => undefined);

    if (!this.isDirty || !this.host.engine || !this.host.loadedFile) {
      return true;
    }

    if (this.host.autosaveEnabled()) {
      const saved = await this.host.saveCurrentPresentation();
      if (saved) return true;
    }

    return this.writeRecoveryCopy(reason);
  }

  async writeRecoveryCopy(reason: string): Promise<boolean> {
    const file = this.host.loadedFile;
    const engine = this.host.engine;
    const sourcePackage = this.host.sourcePackage;
    const sourceBuffer = this.host.sourceBuffer;
    if (!file || !engine || !sourcePackage || !sourceBuffer) {
      new Notice('Could not create a Native PowerPoint recovery copy because the open presentation is unavailable.');
      return false;
    }

    try {
      this.setSaveState('saving', 'Exporting recovery...');
      if (this.shouldPaintSaveProgress()) await flushUi();
      const output = await engine.export();
      let isValidated = true;
      let validationError = '';

      try {
        await this.host.validateExportBeforeSave(
          output,
          engine,
          sourcePackage,
          sourceBuffer,
          (message) => this.setSaveProgress(message)
        );
      } catch (error) {
        isValidated = false;
        validationError = cleanError(error);
      }

      this.setSaveProgress('Writing recovery...');
      if (this.shouldPaintSaveProgress()) await flushUi();

      const recoveryPath = this.getAvailableRecoveryPath(file, isValidated);
      await this.host.app.vault.createBinary(recoveryPath, output);
      this.isDirty = false;
      this.setSaveState('recovered');

      if (isValidated) {
        new Notice(`Unsaved edits were not written to ${file.name}. Recovery copy created while ${reason}: ${recoveryPath}`, 10000);
      } else {
        new Notice(
          `Save validation failed, so ${file.name} was not overwritten. An unvalidated recovery export was created at ${recoveryPath}. ${validationError}`,
          15000
        );
      }

      return true;
    } catch (error) {
      this.setSaveState('failed');
      new Notice(
        `Could not preserve unsaved edits from ${file.name} while ${reason}: ${cleanError(error)}. The original file was not overwritten.`,
        15000
      );
      return false;
    }
  }

  getAvailableRecoveryPath(file: TFile, isValidated: boolean): string {
    const slashIndex = file.path.lastIndexOf('/');
    const parentPath = slashIndex === -1 ? '' : file.path.slice(0, slashIndex);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const recoveryType = isValidated ? 'recovery' : 'unvalidated recovery';
    const baseName = `${file.basename} (Native PowerPoint ${recoveryType} ${timestamp})`;
    let sequence = 0;

    while (true) {
      const suffix = sequence === 0 ? '' : ` ${sequence + 1}`;
      const fileName = `${baseName}${suffix}.${file.extension}`;
      const candidate = normalizePath(parentPath ? `${parentPath}/${fileName}` : fileName);
      if (!this.host.app.vault.getAbstractFileByPath(candidate)) {
        return candidate;
      }
      sequence++;
    }
  }

  setSaveProgress(message: string): void {
    if (this.saveState !== 'saving') return;
    this.host.getSaveStatusElements()?.statusLabelEl.setText(message);
  }

  setSaveState(state: SaveState, savingMessage = 'Saving...'): void {
    this.saveState = state;
    const status = this.host.getSaveStatusElements();
    if (!status) return;

    const { statusEl, statusLabelEl, statusProgressEl } = status;

    const labels: Record<SaveState, string> = {
      idle: 'Ready',
      dirty: 'Unsaved',
      saving: savingMessage,
      saved: 'Saved',
      failed: 'Save failed',
      recovered: 'Recovery saved',
      'view-only': 'View-only'
    };

    statusLabelEl.setText(labels[state]);
    statusEl.dataset.state = state;
    statusEl.setAttribute('aria-busy', state === 'saving' ? 'true' : 'false');
    statusProgressEl?.toggleClass('is-active', state === 'saving');
    statusEl.toggleClass('is-clickable', state === 'failed');
    if (state === 'failed') {
      statusEl.setAttribute('role', 'button');
      statusEl.setAttribute('tabindex', '0');
      statusEl.setAttribute('aria-label', 'Save failed. Click to retry.');
      if (this.lastSaveError) {
        statusEl.title = `${this.lastSaveError} — click to retry`;
      } else {
        statusEl.title = 'Click to retry save';
      }
    } else {
      statusEl.setAttribute('role', 'status');
      statusEl.removeAttribute('tabindex');
      statusEl.removeAttribute('aria-label');
      statusEl.removeAttribute('title');
    }
  }
}
