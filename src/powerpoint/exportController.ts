import { type App, type Menu, Notice, TFile, normalizePath } from 'obsidian';

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
  constructor(private readonly host: ExportHost) {}

  openMenu(anchor: HTMLElement): void {
    const engine = this.host.engine;
    if (!engine || engine.slideCount === 0) {
      new Notice('Open a presentation with at least one slide to export.');
      return;
    }

    const menu = this.host.createNativeMenu();
    menu.addItem((item) =>
      item
        .setTitle('Current slide as PNG')
        .setIcon('image')
        .onClick(() => void this.exportCurrentSlideAsPng())
    );
    menu.addItem((item) =>
      item
        .setTitle('Current slide as PDF')
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(true))
    );
    menu.addSeparator();
    menu.addItem((item) =>
      item
        .setTitle('Whole deck as PDF')
        .setIcon('file-output')
        .onClick(() => void this.exportDeckAsPdf(false))
    );
    menu.addItem((item) =>
      item
        .setTitle('Whole deck as PNGs (zip)')
        .setIcon('file-archive')
        .onClick(() => void this.exportDeckAsPngZip())
    );

    const rect = anchor.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom + 4 });
  }

  async exportCurrentSlideAsPng(): Promise<void> {
    if (!this.host.engine) return;

    try {
      debugLog('export', 'PowerPoint PNG export started', { slide: this.host.currentSlide });
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
        slide: this.host.currentSlide,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PNG export failed', { slide: this.host.currentSlide, error });
      new Notice(`Could not export slide as PNG: ${cleanError(error)}`);
    }
  }

  async exportDeckAsPdf(currentSlideOnly: boolean): Promise<void> {
    const engine = this.host.engine;
    if (!engine) return;

    try {
      const indices = currentSlideOnly
        ? [this.host.currentSlide]
        : Array.from({ length: engine.slideCount }, (_, index) => index);
      const elements = this.host.collectSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      new Notice(currentSlideOnly ? 'Exporting slide to PDF...' : 'Exporting deck to PDF...');
      debugLog('export', 'PowerPoint PDF export started', {
        currentSlideOnly,
        slideCount: indices.length
      });
      const bytes = await exportSlidesToPdf(elements, this.host.ownerDocument);
      const baseName = currentSlideOnly
        ? `${this.getExportBaseName()}-slide-${this.host.currentSlide + 1}`
        : this.getExportBaseName();
      await this.saveExportArtifact(baseName, 'pdf', bytes);
      debugLog('export', 'PowerPoint PDF export completed', {
        currentSlideOnly,
        slideCount: indices.length,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PDF export failed', { currentSlideOnly, error });
      new Notice(`Could not export PDF: ${cleanError(error)}`);
    }
  }

  async exportDeckAsPngZip(): Promise<void> {
    const engine = this.host.engine;
    if (!engine) return;

    try {
      const indices = Array.from({ length: engine.slideCount }, (_, index) => index);
      const elements = this.host.collectSvgElements(indices);
      if (elements.length === 0) {
        throw new Error('No slides could be rendered for export.');
      }

      new Notice('Exporting deck to PNG images...');
      const baseName = this.getExportBaseName();
      debugLog('export', 'PowerPoint PNG zip export started', { slideCount: indices.length });
      const bytes = await exportSlidesToPngZip(elements, this.host.ownerDocument, baseName);
      await this.saveExportArtifact(`${baseName}-slides`, 'zip', bytes);
      debugLog('export', 'PowerPoint PNG zip export completed', {
        slideCount: indices.length,
        bytes: bytes.byteLength
      });
    } catch (error) {
      errorLog('export', 'PowerPoint PNG zip export failed', { error });
      new Notice(`Could not export PNG images: ${cleanError(error)}`);
    }
  }

  private getExportBaseName(): string {
    const file = this.host.sourceFile;
    return file?.basename || 'presentation';
  }

  private getAvailableNumberedPath(path: string): string {
    const lastSlashIndex = path.lastIndexOf('/');
    const folderPrefix = lastSlashIndex >= 0 ? `${path.slice(0, lastSlashIndex)}/` : '';
    const fileName = lastSlashIndex >= 0 ? path.slice(lastSlashIndex + 1) : path;
    const extensionIndex = fileName.lastIndexOf('.');
    const baseName = extensionIndex > 0 ? fileName.slice(0, extensionIndex) : fileName;
    const extension = extensionIndex > 0 ? fileName.slice(extensionIndex) : '';

    for (let index = 2; index < 1000; index += 1) {
      const candidatePath = normalizePath(`${folderPrefix}${baseName} ${index}${extension}`);
      if (!this.host.app.vault.getAbstractFileByPath(candidatePath)) {
        return candidatePath;
      }
    }

    return normalizePath(`${folderPrefix}${baseName} ${Date.now()}${extension}`);
  }

  private async saveExportArtifact(baseName: string, extension: string, data: ArrayBuffer): Promise<void> {
    const source = this.host.sourceFile;
    const folderPath = source?.parent?.path;
    const folderPrefix = folderPath && folderPath !== '/' ? `${folderPath}/` : '';
    const safeBaseName = baseName.replace(/[\\/:*?"<>|]/g, '_') || 'presentation';
    let targetPath = normalizePath(`${folderPrefix}${safeBaseName}.${extension}`);

    const existing = this.host.app.vault.getAbstractFileByPath(targetPath);
    if (existing) {
      if (!(existing instanceof TFile)) {
        new Notice(`${targetPath} already exists and is not a file.`);
        return;
      }

      const choice = await new Promise<'replace' | 'keep-both' | 'cancel'>((resolve) => {
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

      if (choice === 'cancel') return;
      if (choice === 'keep-both') {
        targetPath = this.getAvailableNumberedPath(targetPath);
      }
    }

    const existingTarget = this.host.app.vault.getAbstractFileByPath(targetPath);
    if (existingTarget instanceof TFile) {
      await this.host.app.vault.modifyBinary(existingTarget, data);
    } else {
      await this.host.app.vault.createBinary(targetPath, data);
    }

    debugLog('export', 'PowerPoint export artifact written', {
      targetPath,
      extension,
      bytes: data.byteLength,
      replaced: existingTarget instanceof TFile
    });
    new Notice(`Exported to ${targetPath}`);
  }
}
