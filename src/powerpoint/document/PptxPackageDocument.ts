import { buildZip, type PptxRenderer } from 'pptx-svg';
import { createPresentationRenderer, type PptxRendererBackend } from '../backend/rendererBackend';
import { FontFidelity } from '../../FontFidelity';
import { debugLog, errorLog } from '../../logger';
import { mergeSlideGraphicFramesFromBuffer } from '../../SlideInsertions';
import { getSlidePath } from '../ooxmlXml';
import { preserveSlideExtensionLists } from '../slideExtensionPreserve';

interface SlideOoxmlReadable {
  getSlideOoxml(slideIdx: number): string;
}

export interface PptxPackageDocumentHooks {
  /**
   * Reconcile renderer output against the authoritative lossless package.
   * Renderer state is derived and must never replace this package directly.
   */
  reconcileExport(authoritativePackage: ArrayBuffer, renderedExport: ArrayBuffer): Promise<ArrayBuffer>;
  refreshDerivedState(buffer: ArrayBuffer): Promise<void>;
}

/**
 * Owns the lossless PPTX package and the renderer derived from it.
 *
 * Every mutation that requires a package export must enter through `export()`,
 * `foldLiveSlidesIntoPackage()`, or `syncPackageFromPendingSlides()`. This
 * preserves OOXML the renderer does not model before it becomes the next
 * authoritative package. Prefer `foldLiveSlidesIntoPackage()` when a full WASM
 * `exportPptx()` is unnecessary (history snapshots, slide delete/reorder).
 */
export class PptxPackageDocument {
  private _renderer: PptxRenderer;
  private _rendererBackend: PptxRendererBackend;
  private _fontFidelity: FontFidelity;
  private _slideCount: number;
  private currentBuffer: ArrayBuffer;
  private pendingSlideXml = new Map<number, string>();

  private constructor(
    private readonly hooks: PptxPackageDocumentHooks,
    renderer: PptxRenderer,
    rendererBackend: PptxRendererBackend,
    fontFidelity: FontFidelity,
    slideCount: number,
    buffer: ArrayBuffer,
  ) {
    this._renderer = renderer;
    this._rendererBackend = rendererBackend;
    this._fontFidelity = fontFidelity;
    this._slideCount = slideCount;
    this.currentBuffer = buffer.slice(0);
  }

  static async load(buffer: ArrayBuffer, hooks: PptxPackageDocumentHooks): Promise<PptxPackageDocument> {
    const state = await createPresentationRenderer(buffer);
    const document = new PptxPackageDocument(
      hooks,
      state.renderer,
      state.rendererBackend,
      state.fontFidelity,
      state.slideCount,
      buffer,
    );
    return document;
  }

  get renderer(): PptxRenderer {
    return this._renderer;
  }

  get rendererBackend(): PptxRendererBackend {
    return this._rendererBackend;
  }

  get fontFidelity(): FontFidelity {
    return this._fontFidelity;
  }

  get slideCount(): number {
    return this._slideCount;
  }

  get packageBuffer(): ArrayBuffer {
    return this.currentBuffer;
  }

  set packageBuffer(buffer: ArrayBuffer) {
    this.currentBuffer = buffer.slice(0);
  }

  recordPendingSlideXml(slideIndex: number, xml: string): void {
    this.pendingSlideXml.set(slideIndex, xml);
  }

  async restore(buffer: ArrayBuffer): Promise<void> {
    await this.reload(buffer, undefined);
  }

  async reload(buffer: ArrayBuffer, expectedSlideCount: number | undefined): Promise<void> {
    const state = await createPresentationRenderer(buffer);
    if (expectedSlideCount !== undefined && state.slideCount !== expectedSlideCount) {
      throw new Error(`Slide management export mismatch: expected ${expectedSlideCount}, got ${state.slideCount}.`);
    }
    this._renderer = state.renderer;
    this._rendererBackend = state.rendererBackend;
    this._fontFidelity = state.fontFidelity;
    this._slideCount = state.slideCount;
    this.currentBuffer = buffer.slice(0);
    this.pendingSlideXml.clear();
    await this.hooks.refreshDerivedState(buffer);
  }

  async export(): Promise<ArrayBuffer> {
    const startedAt = Date.now();
    // Slide-local text edits defer folding pending XML into `currentBuffer`.
    // Drain that queue first so reconcile sees the lossless slide parts (and so
    // we do not discard pending without applying it).
    await this.syncPackageFromPendingSlides();
    const pendingSlideCount = this.pendingSlideXml.size;
    debugLog('engine', 'Package export transaction started', {
      op: 'export-package',
      pendingSlideCount,
      authoritativePackage: true,
    });
    try {
      const rawExport = await this.patchSlidesFromRendererOoxml(await this._renderer.exportPptx());
      const reconciledExport = await this.hooks.reconcileExport(this.currentBuffer, rawExport);
      this.currentBuffer = reconciledExport.slice(0);
      this.pendingSlideXml.clear();
      await this.hooks.refreshDerivedState(reconciledExport);
      debugLog('engine', 'Package export transaction committed', {
        op: 'export-package',
        pendingSlideCount,
        outputBytes: reconciledExport.byteLength,
        authoritativePackage: true,
        ms: Date.now() - startedAt,
      });
      return reconciledExport;
    } catch (error) {
      errorLog('engine', 'Package export transaction failed', {
        op: 'export-package',
        pendingSlideCount,
        authoritativePackage: true,
        error,
      });
      throw error;
    }
  }

  async syncPackageFromPendingSlides(): Promise<void> {
    if (this.pendingSlideXml.size === 0) return;

    const startedAt = Date.now();
    const pending = this.pendingSlideXml;
    const slides = Array.from(pending.keys());
    debugLog('engine', 'Pending slide package transaction started', {
      op: 'sync-pending-slides',
      pendingSlideCount: slides.length,
      slides,
      authoritativePackage: true,
    });
    try {
      this.pendingSlideXml = new Map();
      const previousBuffer = this.currentBuffer;
      let buffer = await buildZip(
        previousBuffer,
        new Map(Array.from(pending, ([slideIndex, xml]) => [getSlidePath(slideIndex), xml]))
      );
      for (const slideIndex of slides) {
        buffer = await mergeSlideGraphicFramesFromBuffer(previousBuffer, buffer, slideIndex);
      }
      this.currentBuffer = await preserveSlideExtensionLists(previousBuffer, buffer);
      debugLog('engine', 'Pending slide package transaction committed', {
        op: 'sync-pending-slides',
        pendingSlideCount: slides.length,
        slides,
        outputBytes: this.currentBuffer.byteLength,
        authoritativePackage: true,
        ms: Date.now() - startedAt,
      });
    } catch (error) {
      this.pendingSlideXml = pending;
      errorLog('engine', 'Pending slide package transaction failed', {
        op: 'sync-pending-slides',
        pendingSlideCount: slides.length,
        slides,
        authoritativePackage: true,
        error,
      });
      throw error;
    }
  }

  /**
   * Fold live renderer slide OOXML into the authoritative package without
   * `exportPptx()`. Used for history snapshots and slide-management ops where a
   * full WASM package encode is the dominant cost on image-heavy decks.
   */
  async foldLiveSlidesIntoPackage(): Promise<ArrayBuffer> {
    const startedAt = Date.now();
    await this.syncPackageFromPendingSlides();
    const previousBuffer = this.currentBuffer;
    const reader = this._renderer as Partial<SlideOoxmlReadable>;
    const modifications = new Map<string, string>();
    if (typeof reader.getSlideOoxml === 'function') {
      for (let slideIndex = 0; slideIndex < this._slideCount; slideIndex++) {
        try {
          const slideXml = reader.getSlideOoxml(slideIndex);
          if (slideXml.includes('</p:sld>')) {
            modifications.set(getSlidePath(slideIndex), slideXml);
          }
        } catch {
          // Keep the prior package part when the runtime cannot serialize a slide.
        }
      }
    }

    debugLog('engine', 'Package fold transaction started', {
      op: 'fold-live-slides',
      slideCount: this._slideCount,
      patchedSlideCount: modifications.size,
      authoritativePackage: true,
    });
    try {
      const patched =
        modifications.size > 0 ? await buildZip(previousBuffer, modifications) : previousBuffer.slice(0);
      const reconciledExport = await this.hooks.reconcileExport(previousBuffer, patched);
      this.currentBuffer = reconciledExport.slice(0);
      this.pendingSlideXml.clear();
      await this.hooks.refreshDerivedState(reconciledExport);
      debugLog('engine', 'Package fold transaction committed', {
        op: 'fold-live-slides',
        slideCount: this._slideCount,
        patchedSlideCount: modifications.size,
        outputBytes: reconciledExport.byteLength,
        authoritativePackage: true,
        ms: Date.now() - startedAt,
      });
      return reconciledExport;
    } catch (error) {
      errorLog('engine', 'Package fold transaction failed', {
        op: 'fold-live-slides',
        slideCount: this._slideCount,
        patchedSlideCount: modifications.size,
        authoritativePackage: true,
        error,
      });
      throw error;
    }
  }

  private async patchSlidesFromRendererOoxml(exportedBuffer: ArrayBuffer): Promise<ArrayBuffer> {
    const reader = this._renderer as Partial<SlideOoxmlReadable>;
    if (typeof reader.getSlideOoxml !== 'function') return exportedBuffer;
    const modifications = new Map<string, string>();
    for (let slideIndex = 0; slideIndex < this._slideCount; slideIndex++) {
      try {
        const slideXml = reader.getSlideOoxml(slideIndex);
        if (slideXml.includes('</p:sld>')) modifications.set(getSlidePath(slideIndex), slideXml);
      } catch {
        // Keep the renderer export for slides the runtime cannot serialize.
      }
    }
    return modifications.size > 0 ? buildZip(exportedBuffer, modifications) : exportedBuffer;
  }
}
