import { type App, TFile, normalizePath } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import {
  inspectPowerPointPackage,
  summarizePackageMessages,
  validatePowerPointExport,
  validatePowerPointExportContents,
  type PowerPointPackageInspection
} from '../PowerPointPackage';
import { debugLog, errorLog, warnLog } from '../logger';
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
  readonly t: TranslateFn;
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
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: SaveHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  shouldPaintSaveProgress(): boolean {
    return this.host.getSaveStatusElements()?.statusLabelEl !== null;
  }

  markDirty(): void {
    this.isDirty = true;
    this.editVersion++;
    this.setSaveState('dirty');
    debugLog('save', 'PowerPoint marked dirty', {
      editVersion: this.editVersion,
      autosaveEnabled: this.host.autosaveEnabled()
    });
    if (this.host.autosaveEnabled()) {
      this.scheduleAutosave();
    }
  }

  scheduleAutosave(delayMs = 1500): void {
    this.clearAutosave();
    if (!this.host.autosaveEnabled()) return;

    debugLog('save', 'PowerPoint autosave scheduled', {
      delayMs,
      editVersion: this.editVersion
    });
    this.saveTimer = window.setTimeout(() => {
      this.saveTimer = null;
      debugLog('save', 'PowerPoint autosave started', { editVersion: this.editVersion });
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
      this.notice('powerpoint:notice.openModernToSave');
      return false;
    }

    if (!isEditablePowerPointExtension(file.extension)) {
      if (this.host.viewOnlyReason) {
        this.notice('powerpoint:notice.viewOnlyReason', { reason: this.host.viewOnlyReason });
      } else {
        this.notice('powerpoint:notice.viewOnlyDefault');
      }
      return false;
    }

    if (!this.host.ensureEditable('save changes')) {
      return false;
    }

    const targetVersion = this.editVersion;
    const saveStartedAt = performance.now();
    this.clearAutosave();
    this.setSaveState('saving', this.host.t('powerpoint:save.exporting'));
    debugLog('save', 'PowerPoint save started', {
      file: file.path,
      targetVersion,
      slideCount: engine.slideCount,
      sourceBytes: sourceBuffer.byteLength
    });

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
      this.setSaveProgress(this.host.t('powerpoint:save.writingToVault'));
      if (this.shouldPaintSaveProgress()) await flushUi();
      await this.host.app.vault.modifyBinary(file, output);
      debugLog('save', 'PowerPoint vault write completed', {
        file: file.path,
        bytes: output.byteLength,
        targetVersion
      });

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
      debugLog('save', 'PowerPoint save completed', {
        file: file.path,
        targetVersion,
        currentVersion: this.editVersion,
        dirty: this.isDirty,
        ms: Math.round(performance.now() - saveStartedAt)
      });
      return true;
    } catch (error) {
      const message = cleanError(error);
      this.lastSaveError = message;
      this.setSaveState('failed');
      errorLog('save', 'PowerPoint save failed', {
        file: file.path,
        targetVersion,
        error
      });
      this.notice('powerpoint:notice.couldNotSave', { fileName: file.name, message });
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

    onProgress?.(this.host.t('powerpoint:save.checkingPackage'));
    debugLog('save', 'PowerPoint save validation started', {
      slideCount: engine.slideCount,
      sourceBytes: sourceBuffer.byteLength,
      outputBytes: output.byteLength
    });
    if (this.shouldPaintSaveProgress()) await flushUi();
    const exportedPackage = inspectPowerPointPackage(output);
    const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount);
    if (!validation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);
    }

    onProgress?.(this.host.t('powerpoint:save.verifyingContents'));
    if (this.shouldPaintSaveProgress()) await flushUi();
    const contentValidation = await validatePowerPointExportContents(sourceBuffer, output);
    if (!contentValidation.ok) {
      throw new Error(`Export validation failed: ${summarizePackageMessages(contentValidation.errors)}`);
    }

    onProgress?.(this.host.t('powerpoint:save.roundTripCheck'));
    if (this.shouldPaintSaveProgress()) await flushUi();
    await PresentationEngine.validateRoundTrip(output, engine.slideCount);
    debugLog('save', 'PowerPoint save validation completed', {
      slideCount: engine.slideCount,
      outputBytes: output.byteLength
    });
    return exportedPackage;
  }

  async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    this.clearAutosave();
    await this.savePromise.catch(() => undefined);

    if (!this.isDirty || !this.host.engine || !this.host.loadedFile) {
      debugLog('save', 'PowerPoint teardown preservation skipped', {
        reason,
        dirty: this.isDirty,
        hasEngine: Boolean(this.host.engine),
        hasLoadedFile: Boolean(this.host.loadedFile)
      });
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
      this.notice('powerpoint:notice.recoveryCopyUnavailable');
      return false;
    }

    try {
      debugLog('save', 'PowerPoint recovery export started', {
        file: file.path,
        reason,
        editVersion: this.editVersion
      });
      this.setSaveState('saving', this.host.t('powerpoint:save.exportingRecovery'));
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
        warnLog('save', 'PowerPoint recovery export validation failed', {
          file: file.path,
          reason,
          error
        });
      }

      this.setSaveProgress(this.host.t('powerpoint:save.writingRecovery'));
      if (this.shouldPaintSaveProgress()) await flushUi();

      const recoveryPath = this.getAvailableRecoveryPath(file, isValidated);
      await this.host.app.vault.createBinary(recoveryPath, output);
      this.isDirty = false;
      this.setSaveState('recovered');
      debugLog('save', 'PowerPoint recovery copy written', {
        file: file.path,
        recoveryPath,
        reason,
        validated: isValidated,
        bytes: output.byteLength
      });

      if (isValidated) {
        this.notice(
          'powerpoint:notice.recoveryCopyCreated',
          { fileName: file.name, reason, recoveryPath },
          10000
        );
      } else {
        this.notice(
          'powerpoint:notice.saveValidationFailedRecovery',
          { fileName: file.name, recoveryPath, validationError },
          15000
        );
      }

      return true;
    } catch (error) {
      this.setSaveState('failed');
      errorLog('save', 'PowerPoint recovery copy failed', {
        file: file.path,
        reason,
        error
      });
      this.notice(
        'powerpoint:notice.recoveryCopyFailed',
        { fileName: file.name, reason, message: cleanError(error) },
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

  setSaveState(state: SaveState, savingMessage = this.host.t('powerpoint:save.saving')): void {
    this.saveState = state;
    const status = this.host.getSaveStatusElements();
    if (!status) return;

    const { statusEl, statusLabelEl, statusProgressEl } = status;

    const labels: Record<SaveState, string> = {
      idle: this.host.t('powerpoint:save.ready'),
      dirty: this.host.t('powerpoint:save.unsaved'),
      saving: savingMessage,
      saved: this.host.t('powerpoint:save.saved'),
      failed: this.host.t('powerpoint:save.failed'),
      recovered: this.host.t('powerpoint:save.recovered'),
      'view-only': this.host.t('powerpoint:save.viewOnly')
    };

    statusLabelEl.setText(labels[state]);
    statusEl.dataset.state = state;
    statusEl.setAttribute('aria-busy', state === 'saving' ? 'true' : 'false');
    statusProgressEl?.toggleClass('is-active', state === 'saving');
    statusEl.toggleClass('is-clickable', state === 'failed');
    if (state === 'failed') {
      statusEl.setAttribute('role', 'button');
      statusEl.setAttribute('tabindex', '0');
      statusEl.setAttribute('aria-label', this.host.t('powerpoint:save.failedAria'));
      if (this.lastSaveError) {
        statusEl.title = this.host.t('powerpoint:save.retryTitleWithError', { error: this.lastSaveError });
      } else {
        statusEl.title = this.host.t('powerpoint:save.retryTitle');
      }
    } else {
      statusEl.setAttribute('role', 'status');
      statusEl.removeAttribute('tabindex');
      statusEl.removeAttribute('aria-label');
      statusEl.removeAttribute('title');
    }
  }
}
