import { type App, TFile, normalizePath } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';
import { PresentationEngine } from '../PresentationEngine';
import {
  inspectPowerPointPackage,
  summarizePackageMessages,
  validatePowerPointExport,
  validatePowerPointExportContents,
  type PowerPointPackageInspection
} from '../PowerPointPackage';
import { applySaveStatusPresentation, getSaveStatusLabel } from '../save/saveStatus';
import {
  DocumentSaveCoordinator
} from '../save/DocumentSaveCoordinator';
import type { NativePowerPointSettings } from '../settings';
import { debugLog, errorLog, warnLog } from '../logger';
import { isEditablePowerPointExtension, isModernPowerPointExtension } from './extensions';
import { cleanError } from './runtimeCompat';
import type { SaveState } from './types';

export interface SaveHost {
  readonly t: TranslateFn;
  readonly app: App;
  readonly statusEl: HTMLElement | null;
  getSettings(): NativePowerPointSettings;
  getFile(): TFile | null;
  getLoadedFile(): TFile | null;
  getEngine(): PresentationEngine | null;
  getSourcePackage(): PowerPointPackageInspection | null;
  getSourceBuffer(): ArrayBuffer | null;
  setSource(packageInfo: PowerPointPackageInspection, buffer: ArrayBuffer): void;
  isCurrentPresentation(engine: PresentationEngine, file: TFile): boolean;
  ensureEditable(action: string): boolean;
  getViewOnlyReason(): string;
}

/** State boundary supplied by {@link PresentationSession}. */
export interface SaveStateStore {
  readonly dirty: boolean;
  readonly editVersion: number;
  readonly saveState: SaveState;
  markDirty(): void;
  clearDirty(): void;
  setSaveState(state: SaveState): void;
  setEditVersionForTests?(value: number): void;
}

interface PowerPointSaveContext {
  file: TFile;
  engine: PresentationEngine;
  sourcePackage: PowerPointPackageInspection;
  sourceBuffer: ArrayBuffer;
}

function getDeletionValidationAllowances(engine: PresentationEngine): {
  allowedMarkerRemovals: Record<string, number>;
  allowedUnknownElementRemovals: Record<string, number>;
  allowedPartRemovals: ReadonlySet<string>;
} {
  const deletionAware = engine as Partial<Pick<PresentationEngine,
    | 'getProtectedSlideMarkerRemovalAllowance'
    | 'getUnknownSlideElementRemovalAllowance'
    | 'getPrunedPackageParts'>>;
  return {
    allowedMarkerRemovals: deletionAware.getProtectedSlideMarkerRemovalAllowance?.() ?? {},
    allowedUnknownElementRemovals: deletionAware.getUnknownSlideElementRemovalAllowance?.() ?? {},
    allowedPartRemovals: deletionAware.getPrunedPackageParts?.() ?? new Set<string>(),
  };
}

function clearDeletionValidationAllowances(engine: PresentationEngine): void {
  const deletionAware = engine as Partial<Pick<PresentationEngine, 'clearProtectedSlideMarkerRemovalAllowance'>>;
  deletionAware.clearProtectedSlideMarkerRemovalAllowance?.();
}

/** Owns save serialization, autosave, validation, and recovery copies. */
export class SaveController {
  private readonly notice: TranslateNoticeFn;
  private readonly coordinator: DocumentSaveCoordinator<
    PowerPointSaveContext,
    ArrayBuffer,
    ArrayBuffer,
    PowerPointPackageInspection,
    'manual' | 'autosave'
  >;
  private editBurstCount = 0;
  private editBurstTimer: number | null = null;
  private saveError: string | null = null;
  /** Optional test hook to bypass package validation. */
  validateExportOverride: ((
    output: ArrayBuffer,
    engine: PresentationEngine,
    sourcePackage: PowerPointPackageInspection,
    sourceBuffer: ArrayBuffer,
    options?: { skipRoundTrip?: boolean }
  ) => Promise<PowerPointPackageInspection>) | null = null;

  constructor(
    private readonly host: SaveHost,
    private readonly session: SaveStateStore
  ) {
    this.notice = createTranslateNotice(this.host.t);
    this.coordinator = new DocumentSaveCoordinator({
      adapter: {
        serialize: async ({ engine }) => engine.export(),
        prepareForWrite: async (output) => output,
        validate: (output, context, request) => this.validateExport(
          output,
          context.engine,
          context.sourcePackage,
          context.sourceBuffer,
          { skipRoundTrip: request.source === 'autosave' }
        ),
        persist: async (output, exportedPackage, { file, engine }) => {
          await this.host.app.vault.modifyBinary(file, output);
          debugLog('save', 'PowerPoint vault write completed', { file: file.path, bytes: output.byteLength });
          clearDeletionValidationAllowances(engine);
          if (this.host.isCurrentPresentation(engine, file)) this.host.setSource(exportedPackage, output);
        }
      },
      getContext: () => this.getSaveContext(),
      autosave: {
        enabled: () => this.host.getSettings().autosaveEnabled,
        delayMs: () => this.editBurstCount > 1 ? 4000 : 1500,
        source: 'autosave'
      },
      onStateChange: (state, error) => {
        if (error) {
          this.saveError = cleanError(error);
          const file = this.host.getFile();
          errorLog('save', 'PowerPoint save failed', { file: file?.path ?? null, error });
          if (file) this.notice('powerpoint:notice.couldNotSave', { fileName: file.name, message: this.saveError });
        } else if (state === 'clean') {
          this.saveError = null;
          this.session.clearDirty();
        }
        this.setState(state === 'clean' ? 'saved' : state);
      },
      onAutosaveScheduled: (delayMs, editVersion) => debugLog('save', 'PowerPoint autosave scheduled', {
        file: this.host.getFile()?.path ?? null, delayMs, editVersion
      }),
      onAutosaveStarted: (editVersion) => debugLog('save', 'PowerPoint autosave started', {
        file: this.host.getFile()?.path ?? null, editVersion
      }),
      runAutosave: () => {
        void this.save('autosave');
      }
    });
  }

  get dirty(): boolean {
    return this.session.dirty;
  }

  /** Compatibility alias used by lifecycle tests. */
  get isDirty(): boolean {
    return this.session.dirty;
  }

  set isDirty(value: boolean) {
    if (value) {
      if (!this.session.dirty) this.session.markDirty();
      return;
    }
    this.session.clearDirty();
  }

  get editVersion(): number {
    return this.session.editVersion;
  }

  set editVersion(value: number) {
    this.session.setEditVersionForTests?.(value);
  }

  get state(): SaveState {
    return this.session.saveState;
  }

  get saveState(): SaveState {
    return this.session.saveState;
  }

  get lastSaveError(): string | null {
    return this.saveError;
  }

  get savePromise(): Promise<void> {
    return this.coordinator.waitForIdle();
  }

  set savePromise(promise: Promise<void>) {
    this.coordinator.setActiveSaveForTests(promise);
  }

  setState(state: SaveState): void {
    this.session.setSaveState(state);
    const statusEl = this.host.statusEl;
    if (!statusEl) return;

    applySaveStatusPresentation(statusEl, {
      state,
      label: getSaveStatusLabel(this.host.t, state),
      failedAriaLabel: this.host.t('powerpoint:save.failedAria'),
      failedTitle: this.saveError
        ? this.host.t('powerpoint:save.retryTitleWithError', { error: this.saveError })
        : this.host.t('powerpoint:save.retryTitle')
    });
  }

  markDirty(): void {
    this.session.markDirty();
    this.editBurstCount++;
    if (this.editBurstTimer === null) {
      this.editBurstTimer = window.setTimeout(() => {
        this.editBurstCount = 0;
        this.editBurstTimer = null;
      }, 5000);
    }
    this.coordinator.markDirty();
    debugLog('save', 'PowerPoint marked dirty', {
      file: this.host.getFile()?.path ?? null,
      editVersion: this.session.editVersion,
      autosaveEnabled: this.host.getSettings().autosaveEnabled,
      editBurstCount: this.editBurstCount
    });
  }

  clearAutosave(): void {
    this.coordinator.clearAutosave();
  }

  async save(source: 'manual' | 'autosave' = 'manual'): Promise<boolean> {
    const context = this.getSaveContext();
    if (!context || !isModernPowerPointExtension(context.file.extension)) {
      this.notice('powerpoint:notice.openModernToSave');
      return false;
    }

    if (!isEditablePowerPointExtension(context.file.extension)) {
      const reason = this.host.getViewOnlyReason();
      this.notice(reason ? 'powerpoint:notice.viewOnlyReason' : 'powerpoint:notice.viewOnlyDefault', reason ? { reason } : undefined);
      return false;
    }

    if (!this.host.ensureEditable('save changes')) return false;

    const saveStartedAt = performance.now();
    debugLog('save', 'PowerPoint save started', {
      file: context.file.path,
      targetVersion: this.coordinator.version,
      source,
      sourceBytes: context.sourceBuffer.byteLength,
      slideCount: context.engine.slideCount
    });

    const saved = await this.coordinator.save(source);
    debugLog('save', 'PowerPoint save completed', {
      file: context.file.path,
      currentVersion: this.coordinator.version,
      dirty: this.session.dirty,
      saved,
      ms: Math.round(performance.now() - saveStartedAt)
    });
    return saved;
  }

  async preserveUnsavedChangesForTeardown(reason: string): Promise<boolean> {
    this.clearAutosave();
    await this.coordinator.waitForIdle();

    const engine = this.host.getEngine();
    const loadedFile = this.host.getLoadedFile();
    if (!this.session.dirty || !engine || !loadedFile) {
      debugLog('save', 'PowerPoint teardown preservation skipped', {
        reason,
        dirty: this.session.dirty,
        hasEngine: Boolean(engine),
        hasLoadedFile: Boolean(loadedFile)
      });
      return true;
    }

    if (this.host.getSettings().autosaveEnabled && await this.save()) return true;
    return this.writeRecoveryCopy(reason);
  }

  reset(): void {
    this.clearAutosave();
    if (this.editBurstTimer !== null) {
      window.clearTimeout(this.editBurstTimer);
      this.editBurstTimer = null;
    }
    this.editBurstCount = 0;
    this.saveError = null;
    this.coordinator.reset();
  }

  private getSaveContext(): PowerPointSaveContext | null {
    const file = this.host.getFile();
    const engine = this.host.getEngine();
    const sourcePackage = this.host.getSourcePackage();
    const sourceBuffer = this.host.getSourceBuffer();
    return file && engine && sourcePackage && sourceBuffer
      ? { file, engine, sourcePackage, sourceBuffer }
      : null;
  }

  private async validateExport(
    output: ArrayBuffer,
    engine: PresentationEngine,
    sourcePackage: PowerPointPackageInspection,
    sourceBuffer: ArrayBuffer,
    options: { skipRoundTrip?: boolean } = {}
  ): Promise<PowerPointPackageInspection> {
    if (this.validateExportOverride) {
      return this.validateExportOverride(output, engine, sourcePackage, sourceBuffer, options);
    }
    debugLog('save', 'PowerPoint save validation started', {
      slideCount: engine.slideCount,
      sourceBytes: sourceBuffer.byteLength,
      outputBytes: output.byteLength
    });
    const exportedPackage = inspectPowerPointPackage(output);
    const deletionAllowances = getDeletionValidationAllowances(engine);
    const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount, {
      allowedPartRemovals: deletionAllowances.allowedPartRemovals,
    });
    if (!validation.ok) throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);

    const contentValidation = await validatePowerPointExportContents(
      sourceBuffer,
      output,
      deletionAllowances,
    );
    if (!contentValidation.ok) throw new Error(`Export validation failed: ${summarizePackageMessages(contentValidation.errors)}`);
    if (!options.skipRoundTrip) await PresentationEngine.validateRoundTrip(output, engine.slideCount);

    debugLog('save', 'PowerPoint save validation completed', { slideCount: engine.slideCount, outputBytes: output.byteLength });
    return exportedPackage;
  }

  private async writeRecoveryCopy(reason: string): Promise<boolean> {
    const file = this.host.getLoadedFile();
    const engine = this.host.getEngine();
    const sourcePackage = this.host.getSourcePackage();
    const sourceBuffer = this.host.getSourceBuffer();
    if (!file || !engine || !sourcePackage || !sourceBuffer) {
      this.notice('powerpoint:notice.recoveryCopyUnavailable');
      return false;
    }

    try {
      debugLog('save', 'PowerPoint recovery export started', { file: file.path, reason, editVersion: this.session.editVersion });
      const output = await engine.export();
      let isValidated = true;
      let validationError = '';
      try {
        await this.validateExport(output, engine, sourcePackage, sourceBuffer);
      } catch (error) {
        isValidated = false;
        validationError = cleanError(error);
        warnLog('save', 'PowerPoint recovery export validation failed', { file: file.path, reason, error });
      }

      const recoveryPath = this.getAvailableRecoveryPath(file, isValidated);
      await this.host.app.vault.createBinary(recoveryPath, output);
      this.session.clearDirty();
      this.setState('recovered');
      debugLog('save', 'PowerPoint recovery copy written', {
        file: file.path, recoveryPath, reason, validated: isValidated, bytes: output.byteLength
      });
      this.notice(
        isValidated ? 'powerpoint:notice.recoveryCopyCreated' : 'powerpoint:notice.saveValidationFailedRecovery',
        isValidated
          ? { fileName: file.name, reason, recoveryPath }
          : { fileName: file.name, recoveryPath, validationError },
        isValidated ? 10000 : 15000
      );
      return true;
    } catch (error) {
      this.setState('failed');
      errorLog('save', 'PowerPoint recovery copy failed', { file: file.path, reason, error });
      this.notice('powerpoint:notice.recoveryCopyFailed', { fileName: file.name, reason, message: cleanError(error) }, 15000);
      return false;
    }
  }

  private getAvailableRecoveryPath(file: TFile, isValidated: boolean): string {
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
      if (!this.host.app.vault.getAbstractFileByPath(candidate)) return candidate;
      sequence++;
    }
  }
}
