import { type App, type Menu, TFile, normalizePath } from 'obsidian';

import type { TranslateFn, TranslateNoticeFn } from '../i18n/translate';
import { createTranslateNotice } from '../i18n/translate';

import {
  getVaultFolderPrefix,
  resolveArtifactConflict,
  sanitizeArtifactBaseName,
  writeVaultBinaryArtifact,
  type ArtifactConflictChoice,
  type ArtifactWriteTarget
} from '../export/artifactPaths';
import type { PresentationEngine } from '../PresentationEngine';
import { debugLog, errorLog } from '../logger';
import { exportSlideToPng, exportSlidesToPdf, exportSlidesToPngZip } from '../PowerPointExport';
import { cleanError } from './runtimeCompat';

/**
 * The slice of `NativePowerPointView` that the export subsystem reaches back
 * into. As with the other controllers, this is the regression boundary; the
 * view supplies it via an adapter object so its own members stay `private`.
 */
export interface ExportHost {
  readonly t: TranslateFn;
  readonly engine: PresentationEngine | null;
  readonly currentSlide: number;
  readonly ownerDocument: Document;
  readonly app: App;
  /** The on-disk file backing the deck (the loaded file, falling back to the view's file). */
  readonly sourceFile: TFile | null;
  buildSlideSvgElement(index: number): SVGSVGElement | null;
  collectSvgElements(indices: number[]): SVGSVGElement[];
  createNativeMenu(): Menu;
}

/**
 * Owns the "export deck" UI and the PNG/PDF/zip artifact pipeline. The actual
 * rasterization lives in `PowerPointExport.ts`; this controller orchestrates
 * slide selection, the export menu, and writing the artifact into the vault.
 * Extracted verbatim from `NativePowerPointView`.
 */
export class ExportController {
  private readonly notice: TranslateNoticeFn;

  constructor(private readonly host: ExportHost) {
    this.notice = createTranslateNotice(this.host.t);
  }

  openMenu(anchor: HTMLElement): void {
    const engine = this.host.engine;
    if (!engine || engine.slideCount === 0) {
      this.notice('powerpoint:notice.openToExport');
      return;
    }

    debugLog('export', 'PowerPoint export menu open started', {
      op: 'open-menu',
      slideCount: engine.slideCount
    });
    const menu = this.host.createNativeMenu();
    menu.addItem((item) =>
      item
        .setTitle(this.host.t('powerpoint:export.currentSlidePng'))
        .setIcon('image')
        .onClick(() => void this.exportCurrentSlideAsPng())
    );
    menu.addItem((item) =>
      item
        .setTitle(this.host.t('powerpoint:export.currentSlidePdf'))
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(true))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle(this.host.t('powerpoint:export.wholeDeckPdf'))
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(false))
    );
    menu.addItem((item) =>
      item
        .setTitle(this.host.t('powerpoint:export.wholeDeckPngsZip'))
        .setIcon('file-archive')
        .onClick(() => void this.exportDeckAsPngZip())
    );

    const rect = anchor.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
  }

  async exportCurrentSlideAsPng(): Promise<void> {
    if (!this.host.engine) return;

    try {
      debugLog('export', 'PowerPoint PNG export started', {
        op: 'current-slide-png',
        slide: this.host.currentSlide
      });
      const element = this.host.buildSlideSvgElement(this.host.currentSlide);
      if (!element) {
        throw new Error('This slide could not be rendered for export.');
      }

      const bytes = await exportSlideToPng(element, this.host.ownerDocument);
      await this.saveExportArtifact(
        `${this.getExportBaseName()}-slide-${this.host.currentSlide + 1}`,
        'png',
        bytes
      );
      debugLog('export', 'PowerPoint PNG export completed', {
        op: 'current-slide-png',
        slide: this.host.currentSlide,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PNG export failed', {
        op: 'current-slide-png',
        slide: this.host.currentSlide,
        error
      });
      this.notice('powerpoint:notice.couldNotExportSlidePng', { message: cleanError(error) });
    }
  }

  async exportDeckAsPdf(currentSlideOnly: boolean): Promise<void> {
    const engine = this.host.engine;
    if (!engine) return;

    try {
      debugLog('export', 'PowerPoint PDF export started', {
        op: currentSlideOnly ? 'current-slide-pdf' : 'deck-pdf',
        currentSlideOnly,
        slideCount: currentSlideOnly ? 1 : engine.slideCount
      });
      const indices = currentSlideOnly
        ? [this.host.currentSlide]
        : Array.from({ length: engine.slideCount }, (_, index) => index);
      const elements = this.host.collectSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      this.notice(currentSlideOnly ? 'powerpoint:notice.exportingSlideToPdf' : 'powerpoint:notice.exportingDeckToPdf');
      const bytes = await exportSlidesToPdf(elements, this.host.ownerDocument);
      const baseName = currentSlideOnly
        ? `${this.getExportBaseName()}-slide-${this.host.currentSlide + 1}`
        : this.getExportBaseName();
      await this.saveExportArtifact(baseName, 'pdf', bytes);
      debugLog('export', 'PowerPoint PDF export completed', {
        op: currentSlideOnly ? 'current-slide-pdf' : 'deck-pdf',
        currentSlideOnly,
        slideCount: indices.length,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PDF export failed', {
        op: currentSlideOnly ? 'current-slide-pdf' : 'deck-pdf',
        currentSlideOnly,
        error
      });
      this.notice('powerpoint:notice.couldNotExportPdf', { message: cleanError(error) });
    }
  }

  async exportDeckAsPngZip(): Promise<void> {
    const engine = this.host.engine;
    if (!engine) return;

    try {
      debugLog('export', 'PowerPoint PNG zip export started', {
        op: 'deck-png-zip',
        slideCount: engine.slideCount
      });
      const indices = Array.from({ length: engine.slideCount }, (_, index) => index);
      const elements = this.host.collectSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      this.notice('powerpoint:notice.exportingDeckPngs');
      const baseName = this.getExportBaseName();
      const bytes = await exportSlidesToPngZip(elements, this.host.ownerDocument, baseName);
      await this.saveExportArtifact(`${baseName}-slides`, 'zip', bytes);
      debugLog('export', 'PowerPoint PNG zip export completed', {
        op: 'deck-png-zip',
        slideCount: indices.length,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PNG zip export failed', { op: 'deck-png-zip', error });
      this.notice('powerpoint:notice.couldNotExportPngImages', { message: cleanError(error) });
    }
  }

  private getExportBaseName(): string {
    const file = this.host.sourceFile;
    return file?.basename || 'presentation';
  }

  private promptOverwriteChoice(): Promise<ArtifactConflictChoice> {
    return new Promise<ArtifactConflictChoice>((resolve) => {
      const menu = this.host.createNativeMenu();
      menu.addItem((item) => item.setTitle('Replace existing file').setIcon('refresh-cw').onClick(() => resolve('replace')));
      menu.addItem((item) => item.setTitle('Keep both (numbered copy)').setIcon('copy-plus').onClick(() => resolve('keep-both')));
      menu.addItem((item) => item.setTitle('Cancel export').setIcon('x').onClick(() => resolve('cancel')));
      menu.onHide(() => resolve('cancel'));
      const view = this.host.ownerDocument.defaultView;
      const x = view ? view.innerWidth / 2 : 200;
      const y = view ? view.innerHeight / 3 : 200;
      menu.showAtPosition({ x, y });
    });
  }

  private async saveExportArtifact(baseName: string, extension: string, data: ArrayBuffer): Promise<void> {
    const vault = this.host.app.vault;
    const exists = (path: string): boolean => Boolean(vault.getAbstractFileByPath(path));
    const folderPrefix = getVaultFolderPrefix(this.host.sourceFile?.parent?.path);
    const safeBaseName = sanitizeArtifactBaseName(baseName, 'presentation');
    const requestedPath = normalizePath(`${folderPrefix}${safeBaseName}.${extension}`);

    let target: ArtifactWriteTarget = { path: requestedPath, existingFile: null, replace: false };
    const existing = vault.getAbstractFileByPath(requestedPath);
    if (existing) {
      if (!(existing instanceof TFile)) {
        this.notice('powerpoint:notice.pathExistsNotFile', { path: requestedPath });
        return;
      }

      const resolved = resolveArtifactConflict(requestedPath, existing, await this.promptOverwriteChoice(), exists);
      if (!resolved) return;
      target = resolved;
    }

    const written = await writeVaultBinaryArtifact(vault, target, data);

    debugLog('export', 'PowerPoint export artifact written', {
      op: 'save-artifact',
      targetPath: written.path,
      extension,
      bytes: data.byteLength,
      replaced: target.replace
    });
    this.notice('powerpoint:notice.exportedTo', { path: written.path });
  }
}
