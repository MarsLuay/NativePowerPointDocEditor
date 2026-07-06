import {
  PptxRenderer,
  buildZip,
  bytesToBase64,
  degreesToOoxml,
  emuToPx,
  extractZip,
  getAllShapes,
  getShapeTransform,
  getSlideScale,
  ooxmlToDegrees,
  pxToEmu
} from 'pptx-svg';
import type { ShapeTransform, ZipContents } from 'pptx-svg';
import wasmBytes from 'pptx-svg/wasm';
import {
  getChartDataDescriptor,
  updateChartTextLabel as patchChartTextLabel,
  updateChartData as patchChartData,
  type ChartDataDescriptor,
  type ChartDataGrid,
  type ChartDataUpdate
} from './ChartData';
import { FontFidelity, type FontSubstitution } from './FontFidelity';
import {
  createSlideObjectClipboard,
  pasteSlideObject,
  type SlideObjectClipboard
} from './ShapeClipboard';
import {
  applyParagraphListStyle,
  insertChartIntoPresentation,
  insertTableIntoPresentation,
  mergeMissingPackageParts,
  mergeSlideGraphicFramesFromBuffer,
  type ParagraphListStyle
} from './SlideInsertions';
import {
  DRAWINGML_NAMESPACE,
  IMAGE_RELATIONSHIP_TYPE,
  SHAPE_ELEMENT_NAMES,
  contentTypeForImageExtension,
  createRelationshipsDocument,
  cropPercentToPermille,
  ensureDefaultContentType,
  getBlipEmbedId,
  getDescendants,
  getElementChildren,
  getPartExtension,
  getSlidePath,
  getSlideRelationshipsPath,
  imageExtensionForMime,
  nextImageMediaPath,
  nextRelationshipId,
  normalizeLabelText,
  parseXml,
  resolvePartPath,
  serializeXml,
  setBlipEmbedId,
} from './powerpoint/ooxmlXml';
import {
  applyRunPropertyChange,
  applyRunStyleToParagraphRange,
  disableShrinkAutofit,
  getDrawingParagraphs,
  getDrawingRunText,
  getDrawingRuns,
  getParagraphProperties,
  getRunProperties,
  getShapeRunPositions,
  isParagraphRangeStyled,
  normalizeHexColor,
  replaceTextInParagraph,
  resolvePptxRunAlignment,
  setDrawingParagraphText,
  setDrawingText,
  setDrawingTextRun,
} from './powerpoint/drawingmlText';

import {
  findChartPartPath,
  formatChartAxisValue,
  getChartAxisFormats,
  getChartTextSources,
  getChartTextValues,
  getChartTickRuns,
  removeRedundantTickRuns,
  type ChartAxisFormat,
} from './powerpoint/chartAxisFormatting';
import {
  graftAuthoredRunPropsIntoSlideDoc,
  graftHighlightsIntoSlideDoc,
  needsHighlightRegraft,
  readSlideRunCacheFromDoc,
  remapSlideRunCacheAfterDeletedShape as remapSlideRunCacheEntry,
  restoreLostRunPropsIntoSlideDoc as restoreLostRunPropsFromCache,
  type RunHighlightInfo,
  type SlideRunCacheEntry,
} from './powerpoint/slideRunCache';

export type { RunHighlightInfo };
import {
  normalizeSlideManifest,
  preserveSlideExtensionLists,
} from './powerpoint/slideExtensionPreserve';
import {
  adjacentUnselectedShape,
  applyTransformToShape,
  getShapeBox,
  getShapeElement,
  getShapeElementByRendererIndex,
  getSpTreeShapes,
  nextShapeId,
  qualifyName,
} from './powerpoint/slideShapeOoxml';

const SLIDE_LAYOUT_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const SLIDE_MASTER_RELATIONSHIP_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';

export type InsertableShapeGeometry =
  | 'rect'
  | 'ellipse'
  | 'roundRect'
  | 'line'
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow';

export type SlideLayoutKind = 'blank' | 'title' | 'titleBody';

/** Renderer augmented with the build-time `initJsBackend` patch (see esbuild.config.mjs). */
interface JsBackendCapableRenderer {
  initJsBackend(engine: unknown): void;
}

/**
 * Renderer build-patched with the single-slide entry point (see
 * scripts/lib/patch-pptx-renderer.mjs). Replaces one slide's XML in the live
 * file map and re-parses, so edits skip the whole-deck export/reload.
 */
interface SlideXmlLoadable {
  loadSlideXml(slideIdx: number, xml: string): void;
}

/**
 * True when a failure from the Wasm renderer indicates the runtime lacks
 * WebAssembly GC (Obsidian installer < 1.5.8 / Chromium < 119). Mirrors the
 * detection in NativePowerPointView so both layers agree.
 */
function isWasmGcUnsupportedError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error || '');
  return /WebAssembly GC support|Wasm init failed/i.test(message);
}

import {
  configureForceJsBackendOverrideReader,
  resetForceJsBackendOverride,
  setForceJsBackendOverride,
  shouldForceJsBackend
} from './powerpoint/forceJsBackend';

export { formatChartAxisValue };

export {
  configureForceJsBackendOverrideReader,
  resetForceJsBackendOverride,
  setForceJsBackendOverride
};

async function initJsBackend(renderer: PptxRenderer): Promise<void> {
  const { createPptxJsEngine } = await import('./vendor/pptx-js-engine.mjs');
  (renderer as unknown as JsBackendCapableRenderer).initJsBackend(createPptxJsEngine());
}

/**
 * Initialize the renderer's backend. Prefers the fast Wasm (wasm-gc) engine and,
 * if the runtime cannot run it, lazily loads the pure-JS engine fallback so PPTX
 * files still open on older Obsidian installers. The fallback module is only
 * fetched/evaluated when actually needed. The fallback can also be forced for
 * testing (see {@link setForceJsBackendOverride}).
 */
async function initRendererBackend(renderer: PptxRenderer): Promise<void> {
  if (shouldForceJsBackend()) {
    await initJsBackend(renderer);
    return;
  }

  try {
    await renderer.init(wasmBytes);
  } catch (error) {
    if (!isWasmGcUnsupportedError(error)) throw error;
    await initJsBackend(renderer);
  }
}

function assertOk(result: string, fallback: string): void {
  if (result.startsWith('ERROR:')) {
    throw new Error(result.slice('ERROR:'.length).trim() || fallback);
  }
}


/** Inset crop, expressed as a percentage (0-100) of the source image edge. */
export interface ImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Raw image bytes resolved from an embedded picture, with its MIME type. */
export interface ShapeImageData {
  bytes: Uint8Array;
  mimeType: string;
}

export type GeneratedTextKind = 'chart' | 'table';

export interface GeneratedTextEdit {
  kind: GeneratedTextKind;
  labelIndex: number;
  occurrence: number;
  previousText: string;
}

export type ParagraphAlignment = 'l' | 'ctr' | 'r' | 'just';

/** A specific text run inside a shape, identified by paragraph and run index. */
export interface RunTarget {
  paragraphIndex: number;
  runIndex: number;
}

/** A character range inside one DrawingML paragraph. */
export interface ParagraphTextRange {
  paragraphIndex: number;
  start: number;
  end: number;
}

/**
 * Requested run-level style changes. Omitted fields are left unchanged.
 * `color`/`highlight` use uppercase `RRGGBB` hex (no `#`); `null` clears them.
 */
export interface RunStyleChange {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string | null;
  highlight?: string | null;
}

/** Resolved run-level style read back from a slide, for reflecting toolbar state. */
export interface RunStyleInfo {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontFamily: string | null;
  fontSizePt: number | null;
  color: string | null;
  highlight: string | null;
  alignment: ParagraphAlignment | null;
}

export type ShapeReorderMode = 'front' | 'back' | 'forward' | 'backward';

export interface RenderedSlide {
  svg: string;
  slideCount: number;
}

export interface SlideMoveResult {
  slideIndex: number;
  slideCount: number;
}

export class PresentationEngine {
  private renderer: PptxRenderer;
  private fontFidelity: FontFidelity;
  private currentBuffer: ArrayBuffer;
  private slideCountValue = 0;
  private chartTextValues = new Map<string, string[]>();
  private chartAxisFormats = new Map<string, ChartAxisFormat[]>();
  private chartDataDescriptors = new Map<string, ChartDataDescriptor>();
  private slideBackgroundImageHrefs = new Map<number, string>();
  // Authoritative per-slide run formatting. The renderer's SlideData model drops
  // authored run properties it doesn't model whenever a Wasm-primitive edit
  // re-serializes a slide -- empirically <a:highlight>, <a:hlinkClick>, <a:uFill>
  // and @normalizeH, and any future un-modeled rPr content. This cache is the
  // single source of truth for that formatting: the live view reads highlights
  // from it (getSlideRunHighlights) and the export funnel re-grafts every dropped
  // property into any slide the renderer serialized lossily
  // (reconcileRunPropsIntoBuffer). It is seeded lazily from the lossless model,
  // refreshed on every OOXML-path commit, remapped when a shape is deleted, and
  // reset when the renderer is reinitialized.
  private slideRunCache = new Map<number, SlideRunCacheEntry>();
  // Fast-path (`loadSlideXml`) commits update the live renderer model but not
  // `currentBuffer`, so `currentBuffer` goes stale for those slides' modeled
  // content. `currentBuffer` can't simply be overwritten from the renderer:
  // it is the authoritative source of content the renderer's serialization
  // drops (extension lists, un-modeled graphic frames) -- the very thing
  // `preserveSlideExtensionLists`/`mergeSlideGraphicFramesFromBuffer` graft back.
  // Instead we record each committed slide's lossless XML here and fold it into
  // `currentBuffer` lazily (see `syncCurrentBuffer`), restoring the dropped
  // content from the prior buffer so `currentBuffer` stays authoritative.
  private pendingSlideXml = new Map<number, string>();

  private constructor(renderer: PptxRenderer, fontFidelity: FontFidelity, slideCount: number, buffer: ArrayBuffer) {
    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
  }

  static async load(buffer: ArrayBuffer): Promise<PresentationEngine> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    const engine = new PresentationEngine(renderer, fontFidelity, slideCount, buffer);
    await engine.refreshChartTextValues(buffer);
    await engine.refreshSlideBackgroundImages(buffer);
    return engine;
  }

  private static async createRenderer(buffer: ArrayBuffer): Promise<{
    renderer: PptxRenderer;
    fontFidelity: FontFidelity;
    slideCount: number;
  }> {
    const fontFidelity = new FontFidelity();
    const renderer = new PptxRenderer({
      logLevel: 'error',
      fontFallbacks: fontFidelity.getRendererFallbacks(),
      measureText: (text, fontFace, fontSizePx) => fontFidelity.measureText(text, fontFace, fontSizePx)
    });

    await initRendererBackend(renderer);
    const { slideCount } = await renderer.loadPptx(buffer);
    return { renderer, fontFidelity, slideCount };
  }

  static async validateRoundTrip(buffer: ArrayBuffer, expectedSlideCount: number): Promise<void> {
    const engine = await PresentationEngine.load(buffer);
    if (engine.slideCount !== expectedSlideCount) {
      throw new Error(`Round-trip slide count mismatch: expected ${expectedSlideCount}, got ${engine.slideCount}.`);
    }

    if (engine.slideCount > 0) {
      engine.renderSlide(0);
    }
  }

  get slideCount(): number {
    return this.slideCountValue;
  }

  renderSlide(slideIndex: number): RenderedSlide {
    const svg = this.reconcileRenderedSlideBackground(
      this.renderer.renderSlideSvg(slideIndex),
      slideIndex
    );
    assertOk(svg, 'Could not render slide.');
    return { svg, slideCount: this.slideCountValue };
  }

  renderShape(slideIndex: number, shapeIndex: number): string {
    const svg = this.renderer.renderShapeSvg(slideIndex, shapeIndex);
    assertOk(svg, 'Could not render shape.');
    return svg;
  }

  private reconcileRenderedSlideBackground(svg: string, slideIndex: number): string {
    const backgroundHref = this.slideBackgroundImageHrefs.get(slideIndex);
    if (!backgroundHref || !svg.startsWith('<svg')) return svg;

    try {
      const doc = parseXml(svg, `rendered slide ${slideIndex + 1} SVG`);
      const root = doc.documentElement;
      const width = root.getAttribute('width') ?? '';
      const height = root.getAttribute('height') ?? '';
      if (!width || !height) return svg;

      const background = getElementChildren(root).find((child) =>
        child.localName === 'image'
          && child.getAttribute('x') === '0'
          && child.getAttribute('y') === '0'
          && child.getAttribute('width') === width
          && child.getAttribute('height') === height
          && child.getAttribute('preserveAspectRatio') === 'none'
      );

      if (background) {
        background.setAttribute('href', backgroundHref);
      } else {
        const image = doc.createElementNS(SVG_NAMESPACE, 'image');
        image.setAttribute('x', '0');
        image.setAttribute('y', '0');
        image.setAttribute('width', width);
        image.setAttribute('height', height);
        image.setAttribute('preserveAspectRatio', 'none');
        image.setAttribute('href', backgroundHref);

        const firstChild = root.firstChild;
        const insertBefore = firstChild?.nodeType === 1
          && (firstChild as Element).localName === 'rect'
          && (firstChild as Element).getAttribute('fill') === 'none'
          ? firstChild.nextSibling
          : firstChild;
        root.insertBefore(image, insertBefore);
      }

      return serializeXml(doc);
    } catch {
      return svg;
    }
  }

  getShapes(svg: SVGSVGElement): SVGGElement[] {
    return getAllShapes(svg);
  }

  applyFontFidelity(svg: SVGSVGElement | SVGGElement): FontSubstitution[] {
    return this.fontFidelity.applySvgSubstitutions(svg);
  }

  getChartDataGrid(slideIndex: number, shapeIndex: number): ChartDataGrid | null {
    return this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex))?.grid ?? null;
  }

  async updateChartData(slideIndex: number, shapeIndex: number, update: ChartDataUpdate): Promise<void> {
    const descriptor = this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex));
    if (!descriptor) {
      throw new Error('Could not find chart data for the selected object.');
    }

    const rawExport = await this.exportRendererState();
    const patchedExport = await patchChartData(rawExport, descriptor, update);
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  formatChartAxisLabels(svg: SVGSVGElement, slideIndex: number): void {
    for (const chartGroup of getDescendants(svg, 'g')) {
      if (chartGroup.getAttribute('data-ooxml-shape-type') !== 'chart') {
        continue;
      }

      const shapeIndex = Number(chartGroup.getAttribute('data-ooxml-shape-idx'));
      const formats = this.chartAxisFormats.get(this.getChartTextKey(slideIndex, shapeIndex));

      if (!Number.isInteger(shapeIndex) || formats === undefined) {
        continue;
      }

      for (const orientation of ['horizontal', 'vertical'] as const) {
        const axes = formats.filter((axis) => axis.orientation === orientation);
        const runs = getChartTickRuns(chartGroup).filter((run) => run.orientation === orientation);

        if (axes.length === 0 || runs.length === 0) {
          continue;
        }

        const defaultAxis = axes[0];
        if (!defaultAxis) {
          continue;
        }

        const visibleRuns = axes.length === 1 ? removeRedundantTickRuns(runs, defaultAxis) : runs;

        visibleRuns.forEach((run, index) => {
          const axis = axes[Math.min(index, axes.length - 1)] ?? defaultAxis;
          const step = run.elements.length > 1 ? (axis.max - axis.min) / (run.elements.length - 1) : 0;

          run.elements.forEach((element, tickIndex) => {
            element.textContent = formatChartAxisValue(
              axis.min + step * tickIndex,
              axis.formatCode,
              step,
              axis.date1904
            );
            element.setAttribute('data-native-powerpoint-axis-tick', 'true');
          });
        });
      }
    }
  }

  getShapeTransform(shape: SVGGElement): ShapeTransform {
    return getShapeTransform(shape);
  }

  getSlideScale(svg: SVGSVGElement): number {
    return getSlideScale(svg);
  }

  emuToPx(emu: number): number {
    return emuToPx(emu);
  }

  pxToEmu(px: number): number {
    return pxToEmu(px);
  }

  ooxmlToDegrees(value: number): number {
    return ooxmlToDegrees(value);
  }

  degreesToOoxml(value: number): number {
    return degreesToOoxml(value);
  }

  async updateShapeTransform(
    slideIndex: number,
    shapeIndex: number,
    transform: ShapeTransform
  ): Promise<void> {
    // Seed the highlight cache from the still-lossless model before the edit
    // re-serializes the slide (which drops every <a:highlight>); the cache then
    // feeds both the live overlays and the export-funnel reconciliation.
    this.ensureSlideRunCacheSeeded(slideIndex);
    const result = this.renderer.updateShapeTransform(
      slideIndex,
      shapeIndex,
      Math.round(transform.x),
      Math.round(transform.y),
      Math.max(1, Math.round(transform.cx)),
      Math.max(1, Math.round(transform.cy)),
      Math.round(transform.rot)
    );
    if (!result.startsWith('ERROR:')) {
      return;
    }

    const message = result.slice('ERROR:'.length).trim().toLowerCase();
    if (!message.includes('out of range')) {
      assertOk(result, 'Could not update shape transform.');
    }

    await this.updateShapeTransformInOoxml(slideIndex, shapeIndex, transform);
  }

  private async updateShapeTransformInOoxml(
    slideIndex: number,
    shapeIndex: number,
    transform: ShapeTransform
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElementByRendererIndex(slideDoc, shapeIndex);
    if (!applyTransformToShape(shape, transform)) {
      throw new Error('Could not update shape transform.');
    }

    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  async updateShapeText(slideIndex: number, shapeIndex: number, text: string): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    const textElements = getDescendants(shape, 't')
      .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
    if (textElements.length > 0) {
      setDrawingText(shape, text);
      const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
      await this.reloadFromBuffer(patchedExport, this.slideCountValue);
      return;
    }

    this.ensureSlideRunCacheSeeded(slideIndex);
    const addResult = this.renderer.addShapeText(slideIndex, shapeIndex, text, 1800, 0, 0, 0);
    assertOk(addResult, 'Could not update shape text.');
  }

  async updateParagraphText(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    text: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    setDrawingParagraphText(shape, paragraphIndex, text);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  async updateTextRun(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    runIndex: number,
    text: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    setDrawingTextRun(shape, paragraphIndex, runIndex, text);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  /**
   * Replace text across slides, returning how many occurrences were changed.
   * Replacement works per paragraph over the concatenation of its DrawingML
   * runs (a:t), so a match that spans multiple runs is replaced too (mirroring
   * the scope the find feature searches). The replacement is anchored in the
   * run where the match starts, preserving that run's formatting, and the
   * matched characters are stripped from any later runs the match covers. Pass
   * `slideIndex`/`shapeIndex` to limit the replacement to a single shape (used
   * for "Replace" on the current match); omit them to replace everywhere
   * ("Replace all").
   */
  async replaceText(
    query: string,
    replacement: string,
    options: { matchCase?: boolean; slideIndex?: number; shapeIndex?: number } = {}
  ): Promise<number> {
    if (!query) {
      return 0;
    }

    const matchCase = options.matchCase ?? false;
    const scoped = options.slideIndex !== undefined && options.shapeIndex !== undefined;
    const slideStart = scoped ? (options.slideIndex as number) : 0;
    const slideEnd = scoped ? (options.slideIndex as number) + 1 : this.slideCountValue;

    const rawExport = await this.exportRendererState();
    const zip = await extractZip(rawExport);
    const updatedFiles = new Map<string, string>();
    let total = 0;

    for (let slideIndex = slideStart; slideIndex < slideEnd; slideIndex++) {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) continue;

      const slideDoc = parseXml(slideXml, slidePath);
      let scope: Element | XMLDocument = slideDoc;
      if (scoped) {
        try {
          scope = getShapeElement(slideDoc, options.shapeIndex as number);
        } catch {
          continue;
        }
      }

      const paragraphs = getDescendants(scope, 'p')
        .filter((element) => element.namespaceURI === DRAWINGML_NAMESPACE);
      let slideChanged = false;
      for (const paragraph of paragraphs) {
        const count = replaceTextInParagraph(paragraph, query, replacement, matchCase);
        if (count > 0) {
          total += count;
          slideChanged = true;
        }
      }

      if (slideChanged) {
        updatedFiles.set(slidePath, serializeXml(slideDoc));
      }
    }

    if (total > 0) {
      const patchedExport = await buildZip(rawExport, updatedFiles);
      await this.reloadFromBuffer(patchedExport, this.slideCountValue);
    }

    return total;
  }

  /**
   * Read the resolved style of a single text run for reflecting toolbar state.
   * Only directly-authored run/paragraph properties are reported; values
   * inherited from a placeholder, layout, or master are not resolved here.
   */
  /**
   * Collect every run on the slide that carries an <a:highlight> color. The
   * SVG renderer discards highlights, so the view uses this to repaint them as
   * overlay rects. Shape/paragraph/run indices match {@link getRunStyle} (and
   * therefore the renderer's tspan tags).
   */
  getSlideRunHighlights(slideIndex: number): RunHighlightInfo[] {
    return this.getOrSeedSlideRunCache(slideIndex).highlights;
  }

  /**
   * The authoritative run cache for a slide, seeded lazily from the lossless
   * model on first access. (No in-place Wasm edit can run before the view first
   * renders/queries the slide, so the seed is always lossless. Subsequent Wasm
   * edits go lossy but the cache, once seeded, stays authoritative.)
   */
  private getOrSeedSlideRunCache(slideIndex: number): SlideRunCacheEntry {
    let cached = this.slideRunCache.get(slideIndex);
    if (cached === undefined) {
      cached = this.readSlideRunCacheFromModel(slideIndex);
      this.slideRunCache.set(slideIndex, cached);
    }
    return cached;
  }

  /** Parse the renderer's current OOXML and capture its run formatting. */
  private readSlideRunCacheFromModel(slideIndex: number): SlideRunCacheEntry {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      return readSlideRunCacheFromDoc(slideDoc);
    } catch {
      return { highlights: [], authoredRuns: [] };
    }
  }

  /** Seed the run cache for a slide from the (lossless) model if unseen. */
  private ensureSlideRunCacheSeeded(slideIndex: number): void {
    if (!this.slideRunCache.has(slideIndex)) {
      this.slideRunCache.set(slideIndex, this.readSlideRunCacheFromModel(slideIndex));
    }
  }

  /** Replace a slide's cached run formatting with the model's current (lossless) set. */
  private refreshSlideRunCache(slideIndex: number): void {
    this.slideRunCache.set(slideIndex, this.readSlideRunCacheFromModel(slideIndex));
  }

  /** Drop the whole cache after a renderer reinitialize (indices/contents change). */
  private resetSlideRunCache(): void {
    this.slideRunCache.clear();
  }

  /**
   * After an in-place `deleteShape`, the renderer renumbers the surviving shapes.
   * Remap the cached run formatting to match: drop the deleted shape's entries
   * and shift every higher shape index down by one.
   */
  private remapSlideRunCacheAfterDeletedShape(slideIndex: number, deletedShapeIndex: number): void {
    const cached = this.slideRunCache.get(slideIndex);
    if (!cached) return;
    this.slideRunCache.set(slideIndex, remapSlideRunCacheEntry(cached, deletedShapeIndex));
  }

  /**
   * If a Wasm-primitive edit ran since the cache was last authoritative, the
   * renderer may have stripped authored run properties from this slide. Re-graft
   * them into `slideDoc` before an OOXML-path edit reads/mutates it. When the
   * model is still lossless the document already has them, so this is a no-op and
   * the common path stays byte-identical to before.
   */
  private restoreLostRunPropsIntoSlideDoc(slideIndex: number, slideDoc: XMLDocument): void {
    const cached = this.slideRunCache.get(slideIndex);
    if (!cached) return;
    restoreLostRunPropsFromCache(slideDoc, cached);
  }

  /**
   * Re-graft every run property the renderer dropped into any slide it
   * serialized lossily. This is the single reconciliation point in the export
   * funnel: every save and every reload-based structural edit routes its buffer
   * through here, mirroring how `preserveSlideExtensionLists` restores the
   * per-slide extension lists the renderer strips.
   */
  private async reconcileRunPropsIntoBuffer(buffer: ArrayBuffer): Promise<ArrayBuffer> {
    if (this.slideRunCache.size === 0) return buffer;
    const zip = await extractZip(buffer);
    const modifications = new Map<string, string>();

    for (const [slideIndex, cached] of this.slideRunCache) {
      if (cached.highlights.length === 0 && cached.authoredRuns.length === 0) continue;
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) continue;

      const slideDoc = parseXml(slideXml, slidePath);
      let changed = false;
      // The renderer drops every highlight on a slide at once, so only re-graft
      // (which splits runs) when the export is actually missing some.
      if (needsHighlightRegraft(slideDoc, cached.highlights)) {
        changed = graftHighlightsIntoSlideDoc(slideDoc, cached.highlights) || changed;
      }
      // Generic re-graft is idempotent: it only restores rPr content the export
      // is missing, so it is safe to run unconditionally.
      changed = graftAuthoredRunPropsIntoSlideDoc(slideDoc, cached.authoredRuns) || changed;
      if (changed) {
        modifications.set(slidePath, serializeXml(slideDoc));
      }
    }

    return modifications.size > 0 ? buildZip(buffer, modifications) : buffer;
  }

  /**
   * The authoritative run-only text of a paragraph: the concatenation of every
   * run's `<a:t>` in document order. This is the offset space the range-based
   * styling APIs operate in. The view maps its SVG-derived editor offsets onto
   * this string before calling those APIs, because the rendered SVG drops the
   * whitespace swallowed at soft-wrap boundaries.
   */
  getParagraphRunText(slideIndex: number, shapeIndex: number, paragraphIndex: number): string | null {
    let shape: Element;
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      shape = getShapeElement(slideDoc, shapeIndex);
    } catch {
      return null;
    }

    const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
    if (!paragraph) return null;

    return getDrawingRuns(paragraph).map((run) => getDrawingRunText(run)).join('');
  }

  getRunStyle(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    runIndex: number
  ): RunStyleInfo | null {
    let shape: Element;
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      shape = getShapeElement(slideDoc, shapeIndex);
    } catch {
      return null;
    }

    const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
    if (!paragraph) {
      return null;
    }

    const run = getDrawingRuns(paragraph)[runIndex];
    if (!run) {
      return null;
    }

    const runProperties = getElementChildren(run)
      .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE) ?? null;
    const paragraphProperties = getElementChildren(paragraph)
      .find((element) => element.localName === 'pPr' && element.namespaceURI === DRAWINGML_NAMESPACE) ?? null;

    const bold = runProperties?.getAttribute('b');
    const italic = runProperties?.getAttribute('i');
    const underline = runProperties?.getAttribute('u');
    const fontSize = runProperties?.getAttribute('sz');
    const latin = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'latin')
      : undefined;
    const solidFill = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'solidFill')
      : undefined;
    const fillColor = solidFill
      ? getElementChildren(solidFill).find((element) => element.localName === 'srgbClr')
      : undefined;
    const highlight = runProperties
      ? getElementChildren(runProperties).find((element) => element.localName === 'highlight')
      : undefined;
    const highlightColor = highlight
      ? getElementChildren(highlight).find((element) => element.localName === 'srgbClr')
      : undefined;
    const parsedFontSize = fontSize ? Number(fontSize) : Number.NaN;

    return {
      bold: bold === '1' || bold === 'true',
      italic: italic === '1' || italic === 'true',
      underline: Boolean(underline) && underline !== 'none',
      fontFamily: latin?.getAttribute('typeface') ?? null,
      fontSizePt: Number.isFinite(parsedFontSize) ? parsedFontSize / 100 : null,
      color: this.readColorValue(fillColor),
      highlight: this.readColorValue(highlightColor),
      alignment: resolvePptxRunAlignment(paragraphProperties?.getAttribute('algn') ?? null)
    };
  }

  private readColorValue(colorElement: Element | undefined): string | null {
    const value = colorElement?.getAttribute('val');
    return value ? normalizeHexColor(value) : null;
  }

  /**
   * Apply run-level formatting to a single run, or — when `target` is null — to
   * every run in the shape. All properties (bold/italic/underline/size/color/
   * font/highlight) are applied via direct OOXML editing rather than the WASM
   * renderer's run-style setters: the renderer drops <a:highlight> whenever it
   * re-serializes a slide, so routing every edit through OOXML keeps the
   * highlight intact across subsequent formatting actions.
   */
  async setRunStyle(
    slideIndex: number,
    shapeIndex: number,
    target: RunTarget | null,
    change: RunStyleChange
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const positions = getShapeRunPositions(shape);
      const targets = target
        ? positions.filter(
            (position) => position.paragraphIndex === target.paragraphIndex && position.runIndex === target.runIndex
          )
        : positions;

      let changed = false;
      for (const { run } of targets) {
        applyRunPropertyChange(getRunProperties(run, slideDoc), slideDoc, change);
        changed = true;
      }
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /**
   * Apply run-level formatting to the character range `[startOffset, endOffset)`
   * inside a single paragraph. When the range is collapsed, the run containing
   * the caret is styled. Runs are split at the range boundaries as needed.
   */
  async setRunStyleForRange(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    startOffset: number,
    endOffset: number,
    change: RunStyleChange
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
      if (!paragraph) {
        throw new Error('Could not find the selected text paragraph.');
      }

      const changed = applyRunStyleToParagraphRange(paragraph, slideDoc, startOffset, endOffset, change);
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /**
   * Apply run-level formatting across one or more paragraph ranges in the same
   * shape. This keeps a multi-paragraph toolbar action in one OOXML mutation.
   */
  async setRunStyleForRanges(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    change: RunStyleChange
  ): Promise<void> {
    const normalizedRanges = ranges.filter((range) => (
      Number.isFinite(range.paragraphIndex)
      && Number.isFinite(range.start)
      && Number.isFinite(range.end)
    ));
    if (normalizedRanges.length === 0) return;

    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      let changed = false;
      for (const range of normalizedRanges) {
        const paragraph = paragraphs[range.paragraphIndex];
        if (!paragraph) {
          throw new Error('Could not find the selected text paragraph.');
        }
        changed = applyRunStyleToParagraphRange(paragraph, slideDoc, range.start, range.end, change) || changed;
      }
      if (change.fontSizePt !== undefined && changed) {
        disableShrinkAutofit(shape, slideDoc);
      }
      return changed;
    });
  }

  /** Whether every non-empty run in `[startOffset, endOffset)` has `flag` set. */
  isRangeStyled(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    startOffset: number,
    endOffset: number,
    flag: 'bold' | 'italic' | 'underline'
  ): boolean {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
      if (!paragraph) return false;
      return isParagraphRangeStyled(paragraph, startOffset, endOffset, flag);
    } catch {
      return false;
    }
  }

  /** Whether every non-empty run in every selected range has `flag` set. */
  areRangesStyled(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    flag: 'bold' | 'italic' | 'underline'
  ): boolean {
    const normalizedRanges = ranges.filter((range) => (
      Number.isFinite(range.paragraphIndex)
      && Number.isFinite(range.start)
      && Number.isFinite(range.end)
      && range.start !== range.end
    ));
    if (normalizedRanges.length === 0) return false;

    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      const paragraphs = getDrawingParagraphs(shape);
      return normalizedRanges.every((range) => {
        const paragraph = paragraphs[range.paragraphIndex];
        if (!paragraph) return false;
        return isParagraphRangeStyled(paragraph, range.start, range.end, flag);
      });
    } catch {
      return false;
    }
  }

  /**
   * Set paragraph alignment on a single paragraph, or — when `paragraphIndex`
   * is null — on every paragraph in the shape. Applied via OOXML for the same
   * highlight-preservation reason as {@link setRunStyle}.
   */
  async setParagraphAlignment(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number | null,
    align: ParagraphAlignment
  ): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      const targets = paragraphIndex !== null
        ? [paragraphs[paragraphIndex]]
        : paragraphs;

      let changed = false;
      for (const paragraph of targets) {
        if (!paragraph) continue;
        getParagraphProperties(paragraph, slideDoc).setAttribute('algn', align);
        changed = true;
      }
      return changed;
    });
  }

  /** Set paragraph alignment on the paragraphs touched by the selected ranges. */
  async setParagraphAlignmentForRanges(
    slideIndex: number,
    shapeIndex: number,
    ranges: ParagraphTextRange[],
    align: ParagraphAlignment
  ): Promise<void> {
    const paragraphIndices = new Set(
      ranges
        .map((range) => range.paragraphIndex)
        .filter((paragraphIndex) => Number.isFinite(paragraphIndex))
    );
    if (paragraphIndices.size === 0) return;

    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      const paragraphs = getDrawingParagraphs(shape);
      let changed = false;
      for (const paragraphIndex of paragraphIndices) {
        const paragraph = paragraphs[paragraphIndex];
        if (!paragraph) {
          throw new Error('Could not find the selected text paragraph.');
        }
        getParagraphProperties(paragraph, slideDoc).setAttribute('algn', align);
        changed = true;
      }
      return changed;
    });
  }

  private async editSlideShape(
    slideIndex: number,
    shapeIndex: number,
    mutate: (shape: Element, slideDoc: XMLDocument) => boolean
  ): Promise<void> {
    const slidePath = getSlidePath(slideIndex);
    const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), slidePath);
    // If a prior in-place Wasm edit stripped this slide's highlights, restore the
    // cached set before mutating so the edit (and the committed XML) is lossless.
    this.restoreLostRunPropsIntoSlideDoc(slideIndex, slideDoc);
    const shape = getShapeElementByRendererIndex(slideDoc, shapeIndex);
    if (!mutate(shape, slideDoc)) {
      return;
    }

    await this.commitSlideDoc(slideIndex, slideDoc);
    // The commit reinitialized the model, so its run formatting is authoritative
    // again; capture it (this also records highlight applies/clears and any other
    // run-property edit just made).
    this.refreshSlideRunCache(slideIndex);
  }

  /**
   * Persist an in-memory slide document back into the renderer.
   *
   * Fast path: when the patched renderer exposes `loadSlideXml`, replace just
   * this slide's XML in the live file map and re-parse — no full-deck export,
   * no zip round-trip, no renderer teardown, no chart re-scan. The slide is then
   * re-rendered by the caller via `renderSlide`. Falls back to the whole-deck
   * export/patch/reload when the entry point is unavailable (unpatched renderer).
   */
  private async commitSlideDoc(slideIndex: number, slideDoc: XMLDocument): Promise<void> {
    const serialized = serializeXml(slideDoc);
    const loader = this.renderer as Partial<SlideXmlLoadable>;
    if (typeof loader.loadSlideXml === 'function') {
      loader.loadSlideXml(slideIndex, serialized);
      // The renderer now holds this slide; `currentBuffer` is behind for it.
      // Record the lossless XML so a later `currentBuffer` reader can fold it in
      // (preserving renderer-dropped content) instead of reading stale slide XML.
      this.pendingSlideXml.set(slideIndex, serialized);
      return;
    }

    const rawExport = await this.exportRendererState();
    const patchedExport = await buildZip(rawExport, new Map([[getSlidePath(slideIndex), serialized]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  canUpdateGeneratedText(slideIndex: number, shapeIndex: number, edit: GeneratedTextEdit): boolean {
    if (edit.kind === 'table') return true;
    const chartValues = this.chartTextValues.get(this.getChartTextKey(slideIndex, shapeIndex)) ?? [];
    return chartValues.includes(normalizeLabelText(edit.previousText));
  }

  async updateGeneratedText(slideIndex: number, shapeIndex: number, edit: GeneratedTextEdit, text: string): Promise<void> {
    const rawExport = await this.exportRendererState();
    const zip = await extractZip(rawExport);
    const modifications = new Map<string, string>();

    if (edit.kind === 'table') {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) {
        throw new Error(`Missing slide XML part: ${slidePath}`);
      }

      const slideDoc = parseXml(slideXml, slidePath);
      const shape = getShapeElement(slideDoc, shapeIndex);
      const table = getDescendants(shape, 'tbl')[0];
      const cell = table ? getDescendants(table, 'tc')[edit.labelIndex] : null;
      if (!cell) {
        throw new Error('Could not find the selected table cell.');
      }

      setDrawingText(cell, text);
      modifications.set(slidePath, serializeXml(slideDoc));
    } else {
      const descriptor = this.chartDataDescriptors.get(this.getChartTextKey(slideIndex, shapeIndex));
      if (descriptor) {
        const patchedExport = await patchChartTextLabel(rawExport, descriptor, edit.previousText, edit.occurrence, text);
        await this.reloadFromBuffer(patchedExport, this.slideCountValue);
        return;
      }

      const chartPath = findChartPartPath(zip.textFiles, slideIndex, shapeIndex);
      const chartXml = zip.textFiles.get(chartPath);
      if (!chartXml) {
        throw new Error(`Missing chart XML part: ${chartPath}`);
      }

      const chartDoc = parseXml(chartXml, chartPath);
      const previousText = normalizeLabelText(edit.previousText);
      const matches = getChartTextSources(chartDoc)
        .filter((element) => normalizeLabelText(element.textContent || '') === previousText);
      const source = matches[edit.occurrence] ?? matches[0];
      if (!source) {
        throw new Error('This chart label is generated from chart scale or numeric data and cannot be renamed directly.');
      }

      source.textContent = text;
      modifications.set(chartPath, serializeXml(chartDoc));
    }

    const patchedExport = await buildZip(rawExport, modifications);
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  private parseInsertedShapeIndex(result: string, fallback: string): number {
    assertOk(result, fallback);
    const shapeIndex = Number(result.split(':')[1]);
    if (!Number.isFinite(shapeIndex)) {
      throw new Error('The renderer did not return a valid shape index.');
    }
    return shapeIndex;
  }

  addImage(
    slideIndex: number,
    imageData: Uint8Array,
    mimeType: string,
    widthPx = 320,
    heightPx = 240
  ): number {
    this.ensureSlideRunCacheSeeded(slideIndex);
    const x = pxToEmu(140);
    const y = pxToEmu(120);
    const cx = pxToEmu(widthPx);
    const cy = pxToEmu(heightPx);
    const result = this.renderer.addImage(slideIndex, imageData, mimeType, x, y, cx, cy);
    return this.parseInsertedShapeIndex(result, 'Could not insert image.');
  }

  addShapeGeometry(slideIndex: number, geometry: InsertableShapeGeometry): number {
    this.ensureSlideRunCacheSeeded(slideIndex);
    const x = pxToEmu(160);
    const y = pxToEmu(140);
    const cx = pxToEmu(geometry === 'line' ? 220 : 240);
    const cy = pxToEmu(geometry === 'line' ? 0 : 160);
    const result = this.renderer.addShape(slideIndex, geometry, x, y, cx, cy, 66, 133, 244);
    return this.parseInsertedShapeIndex(result, 'Could not insert shape.');
  }

  addTextBox(slideIndex: number): number {
    this.ensureSlideRunCacheSeeded(slideIndex);
    const x = pxToEmu(180);
    const y = pxToEmu(120);
    const cx = pxToEmu(300);
    const cy = pxToEmu(80);
    const result = this.renderer.addShape(slideIndex, 'rect', x, y, cx, cy, -1, -1, -1);
    const shapeIndex = this.parseInsertedShapeIndex(result, 'Could not add text box.');

    const textResult = this.renderer.addShapeText(slideIndex, shapeIndex, 'New text', 1800, -1, -1, -1);
    assertOk(textResult, 'Could not add text to the new text box.');
    return shapeIndex;
  }

  async addTable(slideIndex: number, rows: number, cols: number): Promise<number> {
    const historyBuffer = await this.exportRendererState();
    const inserted = await insertTableIntoPresentation(historyBuffer, slideIndex, rows, cols);
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    return inserted.shapeIndex;
  }

  async addChart(slideIndex: number): Promise<number> {
    const historyBuffer = await this.exportRendererState();
    const inserted = await insertChartIntoPresentation(historyBuffer, slideIndex);
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    await this.refreshChartTextValues(inserted.buffer);
    return inserted.shapeIndex;
  }

  async applyListStyle(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    style: ParagraphListStyle
  ): Promise<void> {
    // This funnel reconciles the renderer export against `currentBuffer`, so it
    // must reflect any fast-path commits first or those edits would be lost.
    await this.syncCurrentBuffer();
    const rawExport = await this.renderer.exportPptx();
    const mergedSlide = await mergeSlideGraphicFramesFromBuffer(this.currentBuffer, rawExport, slideIndex);
    const mergedPackage = await mergeMissingPackageParts(this.currentBuffer, mergedSlide);
    const patched = await applyParagraphListStyle(mergedPackage, slideIndex, shapeIndex, paragraphIndex, style);
    const preserved = await preserveSlideExtensionLists(this.currentBuffer, patched);
    // Re-graft any highlights the renderer stripped before reloading.
    const reconciled = await this.reconcileRunPropsIntoBuffer(preserved);
    await this.reloadFromBuffer(reconciled, this.slideCountValue);
  }

  deleteShape(slideIndex: number, shapeIndex: number): void {
    this.ensureSlideRunCacheSeeded(slideIndex);
    const result = this.renderer.deleteShape(slideIndex, shapeIndex);
    assertOk(result, 'Could not delete shape.');
    // The renderer renumbers the surviving shapes, so realign the cached
    // highlights (drop this shape, shift higher indices down) to keep the
    // overlays and export reconciliation pointing at the right runs.
    this.remapSlideRunCacheAfterDeletedShape(slideIndex, shapeIndex);
  }

  async copyShape(slideIndex: number, shapeIndex: number): Promise<SlideObjectClipboard> {
    return createSlideObjectClipboard(await this.exportRendererState(), slideIndex, shapeIndex);
  }

  async pasteShape(
    clipboard: SlideObjectClipboard,
    destinationSlideIndex: number
  ): Promise<number> {
    const rawExport = await this.exportRendererState();
    const result = await pasteSlideObject(rawExport, clipboard, destinationSlideIndex);
    await this.reloadFromBuffer(result.buffer, this.slideCountValue);
    return result.shapeIndex;
  }

  async duplicateShape(slideIndex: number, shapeIndex: number): Promise<number> {
    return this.pasteShape(await this.copyShape(slideIndex, shapeIndex), slideIndex);
  }

  /**
   * Apply a structural slide-XML mutation to a slide's shape tree, then reload
   * the renderer from the patched buffer. The mutation runs against the live
   * DOM and its return value is forwarded to the caller. Reordering, grouping,
   * and ungrouping all edit OOXML directly (the renderer has no equivalent API)
   * so the existing shape identities are preserved across the round-trip.
   */
  private async mutateSlideTree<T>(
    slideIndex: number,
    mutate: (slideDoc: XMLDocument, shapeTree: Element) => T
  ): Promise<T> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shapeTree = getDescendants(slideDoc, 'spTree')[0];
    if (!shapeTree) {
      throw new Error('Could not find the slide shape tree.');
    }

    const result = mutate(slideDoc, shapeTree);
    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
    return result;
  }

  /**
   * Change the stacking order of one or more top-level shapes on a slide.
   * Selected shapes keep their relative order. Returns their new shape indices.
   */
  async reorderShapes(
    slideIndex: number,
    shapeIndexes: number[],
    mode: ShapeReorderMode
  ): Promise<number[]> {
    return this.mutateSlideTree(slideIndex, (_slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const selected = new Set(
        shapeIndexes
          .map((index) => shapes[index])
          .filter((element): element is Element => Boolean(element))
      );
      if (selected.size === 0) {
        throw new Error('Select an object to reorder.');
      }

      const ordered = shapes.filter((element) => selected.has(element));
      if (mode === 'front') {
        for (const element of ordered) shapeTree.appendChild(element);
      } else if (mode === 'back') {
        const anchor = shapes.find((element) => !selected.has(element)) ?? null;
        if (anchor) {
          for (const element of ordered) shapeTree.insertBefore(element, anchor);
        }
      } else if (mode === 'forward') {
        for (let index = ordered.length - 1; index >= 0; index--) {
          const element = ordered[index];
          if (!element) continue;
          const next = adjacentUnselectedShape(element, selected, 1);
          if (next) shapeTree.insertBefore(element, next.nextSibling);
        }
      } else {
        for (const element of ordered) {
          const previous = adjacentUnselectedShape(element, selected, -1);
          if (previous) shapeTree.insertBefore(element, previous);
        }
      }

      const finalShapes = getSpTreeShapes(shapeTree);
      return ordered.map((element) => finalShapes.indexOf(element));
    });
  }

  /**
   * Wrap the selected top-level shapes into a new group. The group's bounding
   * box is the union of the children, and chOff/chExt mirror off/ext so each
   * child keeps its slide coordinates. Returns the new group's shape index.
   */
  async groupShapes(slideIndex: number, shapeIndexes: number[]): Promise<number> {
    return this.mutateSlideTree(slideIndex, (slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const selected = new Set(
        shapeIndexes
          .map((index) => shapes[index])
          .filter((element): element is Element => Boolean(element))
      );
      if (selected.size < 2) {
        throw new Error('Select at least two objects to group.');
      }

      const ordered = shapes.filter((element) => selected.has(element));
      const anchor = ordered[0];
      if (!anchor) {
        throw new Error('Could not resolve the objects to group.');
      }

      let minX = Infinity;
      let minY = Infinity;
      let maxX = -Infinity;
      let maxY = -Infinity;
      for (const element of ordered) {
        const box = getShapeBox(element);
        if (!box) continue;
        minX = Math.min(minX, box.x);
        minY = Math.min(minY, box.y);
        maxX = Math.max(maxX, box.x + box.cx);
        maxY = Math.max(maxY, box.y + box.cy);
      }
      if (!Number.isFinite(minX)) {
        minX = 0;
        minY = 0;
        maxX = 0;
        maxY = 0;
      }

      const offsetX = minX;
      const offsetY = minY;
      const extentCx = Math.max(1, maxX - minX);
      const extentCy = Math.max(1, maxY - minY);
      const presentationNs = shapeTree.namespaceURI;
      const newId = nextShapeId(slideDoc);

      const groupShape = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'grpSp'));
      const nonVisual = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'nvGrpSpPr'));
      const cNvPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'cNvPr'));
      cNvPr.setAttribute('id', String(newId));
      cNvPr.setAttribute('name', `Group ${newId}`);
      const cNvGrpSpPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'cNvGrpSpPr'));
      const nvPr = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'nvPr'));
      nonVisual.appendChild(cNvPr);
      nonVisual.appendChild(cNvGrpSpPr);
      nonVisual.appendChild(nvPr);

      const groupProps = slideDoc.createElementNS(presentationNs, qualifyName(shapeTree, 'grpSpPr'));
      const xfrm = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:xfrm');
      const off = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:off');
      off.setAttribute('x', String(Math.round(offsetX)));
      off.setAttribute('y', String(Math.round(offsetY)));
      const ext = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:ext');
      ext.setAttribute('cx', String(Math.round(extentCx)));
      ext.setAttribute('cy', String(Math.round(extentCy)));
      const chOff = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:chOff');
      chOff.setAttribute('x', String(Math.round(offsetX)));
      chOff.setAttribute('y', String(Math.round(offsetY)));
      const chExt = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:chExt');
      chExt.setAttribute('cx', String(Math.round(extentCx)));
      chExt.setAttribute('cy', String(Math.round(extentCy)));
      xfrm.appendChild(off);
      xfrm.appendChild(ext);
      xfrm.appendChild(chOff);
      xfrm.appendChild(chExt);
      groupProps.appendChild(xfrm);

      groupShape.appendChild(nonVisual);
      groupShape.appendChild(groupProps);
      shapeTree.insertBefore(groupShape, anchor);
      for (const element of ordered) groupShape.appendChild(element);

      return getSpTreeShapes(shapeTree).indexOf(groupShape);
    });
  }

  /**
   * Unwrap a group: move its child shapes back into the slide shape tree at the
   * group's position and remove the now-empty group. Child coordinates stay
   * valid because grouped shapes are stored with chOff/chExt equal to off/ext.
   * Returns the resulting shape indices of the freed children.
   */
  async ungroupShapes(slideIndex: number, shapeIndex: number): Promise<number[]> {
    return this.mutateSlideTree(slideIndex, (_slideDoc, shapeTree) => {
      const shapes = getSpTreeShapes(shapeTree);
      const group = shapes[shapeIndex];
      if (!group || group.localName !== 'grpSp') {
        throw new Error('Select a group to ungroup.');
      }

      const children = getElementChildren(group).filter((element) =>
        SHAPE_ELEMENT_NAMES.has(element.localName)
      );
      for (const child of children) shapeTree.insertBefore(child, group);
      shapeTree.removeChild(group);

      const finalShapes = getSpTreeShapes(shapeTree);
      return children.map((element) => finalShapes.indexOf(element));
    });
  }

  async addSlide(afterIndex: number): Promise<SlideMoveResult> {
    const sourceIndex = Math.max(0, Math.min(afterIndex, this.slideCountValue - 1));
    const { slideCount, insertedIdx } = await this.renderer.addSlide(afterIndex, sourceIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: insertedIdx, slideCount };
  }

  async deleteSlide(slideIndex: number): Promise<SlideMoveResult> {
    if (this.slideCountValue <= 1) {
      throw new Error('A presentation must keep at least one slide.');
    }

    const { slideCount } = await this.renderer.deleteSlide(slideIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: Math.min(slideIndex, slideCount - 1), slideCount };
  }

  async moveSlide(slideIndex: number, direction: -1 | 1): Promise<SlideMoveResult> {
    const targetIndex = slideIndex + direction;
    if (targetIndex < 0 || targetIndex >= this.slideCountValue) {
      return { slideIndex, slideCount: this.slideCountValue };
    }

    const order = Array.from({ length: this.slideCountValue }, (_, index) => index);
    const [moved] = order.splice(slideIndex, 1);
    if (moved === undefined) {
      return { slideIndex, slideCount: this.slideCountValue };
    }
    order.splice(targetIndex, 0, moved);
    const { slideCount } = await this.renderer.reorderSlides(order);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: targetIndex, slideCount };
  }

  async duplicateSlide(slideIndex: number): Promise<SlideMoveResult> {
    const { slideCount, insertedIdx } = await this.renderer.addSlide(slideIndex, slideIndex);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: insertedIdx, slideCount };
  }

  async reorderSlides(newOrder: number[]): Promise<SlideMoveResult> {
    const { slideCount } = await this.renderer.reorderSlides(newOrder);
    await this.reloadAfterSlideManagement(slideCount);
    return { slideIndex: 0, slideCount };
  }

  async addSlideWithLayout(afterIndex: number, layout: SlideLayoutKind): Promise<SlideMoveResult> {
    const result = await this.addSlide(afterIndex);
    if (layout === 'blank') {
      return result;
    }

    const slideIndex = result.slideIndex;
    const { cx, cy } = await this.getSlideSizeEmu();
    const margin = pxToEmu(48);
    const contentWidth = Math.max(pxToEmu(120), cx - margin * 2);

    if (layout === 'title') {
      const titleHeight = pxToEmu(140);
      this.addLayoutPlaceholder(
        slideIndex,
        'Click to add title',
        margin,
        Math.max(margin, Math.round((cy - titleHeight) / 2)),
        contentWidth,
        titleHeight,
        4000,
        true
      );
    } else {
      const titleHeight = pxToEmu(120);
      this.addLayoutPlaceholder(slideIndex, 'Click to add title', margin, margin, contentWidth, titleHeight, 3600, true);
      const bodyTop = margin + titleHeight + pxToEmu(24);
      const bodyHeight = Math.max(pxToEmu(120), cy - bodyTop - margin);
      this.addLayoutPlaceholder(slideIndex, 'Click to add text', margin, bodyTop, contentWidth, bodyHeight, 1800, false);
    }

    return result;
  }

  private addLayoutPlaceholder(
    slideIndex: number,
    text: string,
    x: number,
    y: number,
    cx: number,
    cy: number,
    fontSize: number,
    center: boolean
  ): void {
    const shapeResult = this.renderer.addShape(
      slideIndex,
      'rect',
      Math.round(x),
      Math.round(y),
      Math.max(1, Math.round(cx)),
      Math.max(1, Math.round(cy)),
      -1,
      -1,
      -1
    );
    const shapeIndex = this.parseInsertedShapeIndex(shapeResult, 'Could not add layout placeholder.');
    assertOk(
      this.renderer.addShapeText(slideIndex, shapeIndex, text, fontSize, -1, -1, -1),
      'Could not add layout placeholder text.'
    );
    if (center) {
      this.renderer.updateParagraphAlign(slideIndex, shapeIndex, 0, 'ctr');
    }
  }

  async getSlideSizeEmu(): Promise<{ cx: number; cy: number }> {
    const fallback = { cx: 9144000, cy: 6858000 };
    try {
      const presentationPath = 'ppt/presentation.xml';
      const zip = await extractZip(this.currentBuffer);
      const presentationXml = zip.textFiles.get(presentationPath);
      if (!presentationXml) {
        return fallback;
      }

      const slideSize = getDescendants(parseXml(presentationXml, presentationPath), 'sldSz')[0];
      const cx = Number(slideSize?.getAttribute('cx'));
      const cy = Number(slideSize?.getAttribute('cy'));
      return {
        cx: Number.isFinite(cx) && cx > 0 ? cx : fallback.cx,
        cy: Number.isFinite(cy) && cy > 0 ? cy : fallback.cy
      };
    } catch {
      return fallback;
    }
  }

  async setSlideBackgroundColor(slideIndex: number, hex: string): Promise<void> {
    const normalizedHex = hex.replace(/^#/, '').trim().toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(normalizedHex)) {
      throw new Error('Background color must be a 6-digit RRGGBB hex value.');
    }

    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const commonSlide = getDescendants(slideDoc, 'cSld')[0];
    if (!commonSlide) {
      throw new Error('Slide is missing its common slide data.');
    }

    const presentationNamespace = slideDoc.documentElement.namespaceURI;
    for (const existingBackground of getElementChildren(commonSlide).filter((element) => element.localName === 'bg')) {
      commonSlide.removeChild(existingBackground);
    }

    const background = slideDoc.createElementNS(presentationNamespace, 'p:bg');
    const backgroundProperties = slideDoc.createElementNS(presentationNamespace, 'p:bgPr');
    const solidFill = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
    const color = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
    color.setAttribute('val', normalizedHex);
    solidFill.appendChild(color);
    backgroundProperties.appendChild(solidFill);
    backgroundProperties.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:effectLst'));
    background.appendChild(backgroundProperties);
    commonSlide.insertBefore(background, commonSlide.firstChild);

    const patchedExport = await buildZip(rawExport, new Map([[slidePath, serializeXml(slideDoc)]]));
    await this.reloadFromBuffer(patchedExport, this.slideCountValue);
  }

  getSlideBackgroundColor(slideIndex: number): string | null {
    try {
      const slideXml = this.renderer.getSlideOoxml(slideIndex);
      if (!slideXml || slideXml.startsWith('ERROR:')) {
        return null;
      }

      const commonSlide = getDescendants(parseXml(slideXml, getSlidePath(slideIndex)), 'cSld')[0];
      const background = commonSlide
        ? getElementChildren(commonSlide).find((element) => element.localName === 'bg')
        : undefined;
      if (!background) {
        return null;
      }

      const value = getDescendants(background, 'srgbClr')[0]?.getAttribute('val');
      return value && /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : null;
    } catch {
      return null;
    }
  }

  /**
   * Report whether a top-level shape is an embedded picture (`p:pic` with an
   * `<a:blip>`). Read directly from the slide OOXML so image-only menu items
   * are gated on the authoritative model rather than the rendered SVG.
   */
  isImageShape(slideIndex: number, shapeIndex: number): boolean {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      if (shape.localName !== 'pic') return false;
      return getDescendants(shape, 'blip').some((blip) => Boolean(getBlipEmbedId(blip)));
    } catch {
      return false;
    }
  }

  /**
   * Read the current inset crop (`<a:srcRect>`) of a picture as percentages.
   * Returns zeros when the picture is uncropped, or null when the shape is not
   * a picture.
   */
  getImageCrop(slideIndex: number, shapeIndex: number): ImageCrop | null {
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      if (shape.localName !== 'pic') return null;

      const srcRect = getDescendants(shape, 'srcRect')[0];
      const read = (attribute: string): number => {
        const value = Number(srcRect?.getAttribute(attribute));
        return Number.isFinite(value) ? value / 1000 : 0;
      };
      return { left: read('l'), top: read('t'), right: read('r'), bottom: read('b') };
    } catch {
      return null;
    }
  }

  /**
   * Apply an inset crop to a picture via `<a:srcRect>` inside its `<a:blipFill>`.
   * Percentages are stored in OOXML 1000ths-of-a-percent units. Position and
   * size are untouched.
   */
  async setImageCrop(slideIndex: number, shapeIndex: number, crop: ImageCrop): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      if (shape.localName !== 'pic') {
        throw new Error('The selected object is not an image.');
      }

      const blipFill = getDescendants(shape, 'blipFill')[0];
      if (!blipFill) {
        throw new Error('The selected image has no picture fill to crop.');
      }

      let srcRect = getElementChildren(blipFill)
        .find((element) => element.localName === 'srcRect' && element.namespaceURI === DRAWINGML_NAMESPACE);
      if (!srcRect) {
        srcRect = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:srcRect');
        // CT_BlipFillProperties order: blip, srcRect, (tile | stretch).
        const fillMode = getElementChildren(blipFill)
          .find((element) => element.localName === 'stretch' || element.localName === 'tile');
        if (fillMode) {
          blipFill.insertBefore(srcRect, fillMode);
        } else {
          blipFill.appendChild(srcRect);
        }
      }

      srcRect.setAttribute('l', String(cropPercentToPermille(crop.left)));
      srcRect.setAttribute('t', String(cropPercentToPermille(crop.top)));
      srcRect.setAttribute('r', String(cropPercentToPermille(crop.right)));
      srcRect.setAttribute('b', String(cropPercentToPermille(crop.bottom)));
      return true;
    });
  }

  /**
   * Reset a picture to its original appearance: removes any inset crop
   * (`<a:srcRect>`) and common recolor effects (duotone, biLevel, grayscl, lum,
   * clrChange) from the `<a:blip>`. Position, size, and the embedded image are
   * preserved.
   */
  async resetImage(slideIndex: number, shapeIndex: number): Promise<void> {
    const recolorEffects = new Set(['duotone', 'biLevel', 'grayscl', 'lum', 'clrChange']);
    await this.editSlideShape(slideIndex, shapeIndex, (shape) => {
      if (shape.localName !== 'pic') {
        throw new Error('The selected object is not an image.');
      }

      const blipFill = getDescendants(shape, 'blipFill')[0];
      if (!blipFill) return false;

      let changed = false;
      for (const srcRect of getElementChildren(blipFill).filter((element) => element.localName === 'srcRect')) {
        blipFill.removeChild(srcRect);
        changed = true;
      }

      const blip = getElementChildren(blipFill).find((element) => element.localName === 'blip');
      if (blip) {
        for (const effect of getElementChildren(blip).filter((element) => recolorEffects.has(element.localName))) {
          blip.removeChild(effect);
          changed = true;
        }
      }
      return changed;
    });
  }

  /**
   * Toggle a horizontal or vertical flip on a shape by editing the `flipH` /
   * `flipV` attributes of its `<a:xfrm>`. The renderer's transform API does not
   * expose flip, so this is applied directly in OOXML.
   */
  async flipShape(slideIndex: number, shapeIndex: number, axis: 'horizontal' | 'vertical'): Promise<void> {
    await this.editSlideShape(slideIndex, shapeIndex, (shape) => {
      const xfrm = getDescendants(shape, 'xfrm')[0];
      if (!xfrm) {
        throw new Error('The selected object cannot be flipped.');
      }

      const attribute = axis === 'horizontal' ? 'flipH' : 'flipV';
      const current = xfrm.getAttribute(attribute);
      if (current === '1' || current === 'true') {
        xfrm.removeAttribute(attribute);
      } else {
        xfrm.setAttribute(attribute, '1');
      }
      return true;
    });
  }

  /**
   * Swap the picture's embedded image for new bytes while preserving its
   * position, size, and crop. A fresh media part and slide relationship are
   * added, the `<a:blip r:embed>` is repointed, and the content-type default is
   * registered for the new extension. The previous media part is left in place.
   */
  async replaceImage(
    slideIndex: number,
    shapeIndex: number,
    bytes: Uint8Array,
    mimeType: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const shape = getShapeElement(slideDoc, shapeIndex);
    if (shape.localName !== 'pic') {
      throw new Error('The selected object is not an image.');
    }

    const blip = getDescendants(shape, 'blip')[0];
    if (!blip) {
      throw new Error('The selected image has no embedded picture data.');
    }

    const extension = imageExtensionForMime(mimeType);
    const mediaPath = nextImageMediaPath(zip.textFiles, zip.binaryFiles, extension);
    const relationship = this.buildSlideImageRelationship(zip, slideIndex, mediaPath);
    setBlipEmbedId(blip, relationship.relationshipId);

    const contentTypesDoc = parseXml(
      zip.textFiles.get('[Content_Types].xml') ?? '<Types/>',
      '[Content_Types].xml'
    );
    ensureDefaultContentType(contentTypesDoc, extension, contentTypeForImageExtension(extension));

    const textModifications = new Map<string, string>([
      [slidePath, serializeXml(slideDoc)],
      [relationship.relationshipsPath, relationship.relationshipsXml],
      ['[Content_Types].xml', serializeXml(contentTypesDoc)]
    ]);
    const binaryModifications = new Map<string, Uint8Array>([[mediaPath, bytes]]);

    const patched = await buildZip(rawExport, textModifications, undefined, binaryModifications);
    await this.reloadFromBuffer(patched, this.slideCountValue);
  }

  /**
   * Read the bytes of a picture's embedded image by resolving its
   * `<a:blip r:embed>` relationship to a media part. Returns null when the
   * shape is not a picture or the media part cannot be located.
   */
  async getShapeImageData(slideIndex: number, shapeIndex: number): Promise<ShapeImageData | null> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) return null;

    const shape = getShapeElement(parseXml(slideXml, slidePath), shapeIndex);
    if (shape.localName !== 'pic') return null;

    const blip = getDescendants(shape, 'blip')[0];
    const relationshipId = blip ? getBlipEmbedId(blip) : null;
    if (!relationshipId) return null;

    const relationshipsPath = getSlideRelationshipsPath(slideIndex);
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    if (!relationshipsXml) return null;

    const relationship = getDescendants(parseXml(relationshipsXml, relationshipsPath), 'Relationship')
      .find((element) => element.getAttribute('Id') === relationshipId);
    const target = relationship?.getAttribute('Target');
    if (!target || relationship?.getAttribute('TargetMode') === 'External') return null;

    const mediaPath = resolvePartPath(slidePath, target);
    const bytes = zip.binaryFiles.get(mediaPath);
    if (!bytes) return null;

    return {
      bytes: bytes.slice(),
      mimeType: contentTypeForImageExtension(getPartExtension(mediaPath))
    };
  }

  /**
   * Set the slide background to a stretched image, mirroring how
   * {@link setSlideBackgroundColor} rebuilds `<p:bg><p:bgPr>`. A new media part
   * and slide relationship are created and any existing background is replaced.
   */
  async setSlideBackgroundImage(
    slideIndex: number,
    bytes: Uint8Array,
    mimeType: string
  ): Promise<void> {
    const rawExport = await this.exportRendererState();
    const slidePath = getSlidePath(slideIndex);
    const zip = await extractZip(rawExport);
    const slideXml = zip.textFiles.get(slidePath);
    if (!slideXml) {
      throw new Error(`Missing slide XML part: ${slidePath}`);
    }

    const slideDoc = parseXml(slideXml, slidePath);
    const commonSlide = getDescendants(slideDoc, 'cSld')[0];
    if (!commonSlide) {
      throw new Error('Slide is missing its common slide data.');
    }

    const extension = imageExtensionForMime(mimeType);
    const mediaPath = nextImageMediaPath(zip.textFiles, zip.binaryFiles, extension);
    const relationship = this.buildSlideImageRelationship(zip, slideIndex, mediaPath);

    const presentationNamespace = slideDoc.documentElement.namespaceURI;
    for (const existingBackground of getElementChildren(commonSlide).filter((element) => element.localName === 'bg')) {
      commonSlide.removeChild(existingBackground);
    }

    const background = slideDoc.createElementNS(presentationNamespace, 'p:bg');
    const backgroundProperties = slideDoc.createElementNS(presentationNamespace, 'p:bgPr');
    const blipFill = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:blipFill');
    const blip = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:blip');
    setBlipEmbedId(blip, relationship.relationshipId);
    blipFill.appendChild(blip);
    const stretch = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:stretch');
    stretch.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:fillRect'));
    blipFill.appendChild(stretch);
    backgroundProperties.appendChild(blipFill);
    backgroundProperties.appendChild(slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:effectLst'));
    background.appendChild(backgroundProperties);
    commonSlide.insertBefore(background, commonSlide.firstChild);

    const contentTypesDoc = parseXml(
      zip.textFiles.get('[Content_Types].xml') ?? '<Types/>',
      '[Content_Types].xml'
    );
    ensureDefaultContentType(contentTypesDoc, extension, contentTypeForImageExtension(extension));

    const textModifications = new Map<string, string>([
      [slidePath, serializeXml(slideDoc)],
      [relationship.relationshipsPath, relationship.relationshipsXml],
      ['[Content_Types].xml', serializeXml(contentTypesDoc)]
    ]);
    const binaryModifications = new Map<string, Uint8Array>([[mediaPath, bytes]]);

    const patched = await buildZip(rawExport, textModifications, undefined, binaryModifications);
    await this.reloadFromBuffer(patched, this.slideCountValue);
  }

  /**
   * Compose a new image relationship for a slide's `.rels` part. Returns the
   * new relationship id alongside the serialized relationships XML and its part
   * path, so the caller can include them in a single buildZip pass (the
   * extracted zip is not mutated here).
   */
  private buildSlideImageRelationship(
    zip: { textFiles: Map<string, string>; binaryFiles: Map<string, Uint8Array> },
    slideIndex: number,
    mediaPath: string
  ): { relationshipId: string; relationshipsPath: string; relationshipsXml: string } {
    const relationshipsPath = getSlideRelationshipsPath(slideIndex);
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    const relationships = relationshipsXml
      ? parseXml(relationshipsXml, relationshipsPath)
      : createRelationshipsDocument();

    const relationshipId = nextRelationshipId(relationships);
    const relationship = relationships.createElementNS(
      relationships.documentElement.namespaceURI,
      'Relationship'
    );
    relationship.setAttribute('Id', relationshipId);
    relationship.setAttribute('Type', IMAGE_RELATIONSHIP_TYPE);
    relationship.setAttribute('Target', `../media/${mediaPath.split('/').pop()}`);
    relationships.documentElement.appendChild(relationship);

    return {
      relationshipId,
      relationshipsPath,
      relationshipsXml: serializeXml(relationships)
    };
  }

  async export(): Promise<ArrayBuffer> {
    return this.exportRendererState();
  }

  async restoreSnapshot(buffer: ArrayBuffer): Promise<void> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
    this.resetSlideRunCache();
    this.pendingSlideXml.clear();
    await this.refreshChartTextValues(buffer);
    await this.refreshSlideBackgroundImages(buffer);
  }

  private async reloadAfterSlideManagement(expectedSlideCount: number): Promise<void> {
    const rawExport = await this.exportRendererState();
    const normalizedExport = await normalizeSlideManifest(rawExport, expectedSlideCount);
    await this.reloadFromBuffer(normalizedExport, expectedSlideCount);
  }

  private async reloadFromBuffer(buffer: ArrayBuffer, expectedSlideCount: number): Promise<void> {
    const { renderer, fontFidelity, slideCount } = await PresentationEngine.createRenderer(buffer);
    if (slideCount !== expectedSlideCount) {
      throw new Error(`Slide management export mismatch: expected ${expectedSlideCount}, got ${slideCount}.`);
    }

    this.renderer = renderer;
    this.fontFidelity = fontFidelity;
    this.currentBuffer = buffer.slice(0);
    this.slideCountValue = slideCount;
    // The reinitialized model serves lossless highlights again; drop the stale
    // cache (shape indices may have shifted) and let it re-seed on demand.
    this.resetSlideRunCache();
    // `currentBuffer` is now this freshly-reloaded buffer, so any recorded
    // fast-path slide XML is obsolete.
    this.pendingSlideXml.clear();
    await this.refreshChartTextValues(buffer);
    await this.refreshSlideBackgroundImages(buffer);
  }

  private async exportRendererState(): Promise<ArrayBuffer> {
    const rawExport = await this.renderer.exportPptx();
    const preservedExport = await preserveSlideExtensionLists(this.currentBuffer, rawExport);
    const reconciledExport = await this.reconcileRunPropsIntoBuffer(preservedExport);
    this.currentBuffer = reconciledExport.slice(0);
    await this.refreshSlideBackgroundImages(reconciledExport);
    // The full export already reflects every fast-path commit (they live in the
    // renderer model), so any recorded slide XML is now folded in by definition.
    this.pendingSlideXml.clear();
    return reconciledExport;
  }

  /**
   * Bring `currentBuffer` in sync with fast-path (`loadSlideXml`) commits that
   * updated only the renderer model. Folds each recorded slide's lossless XML
   * into `currentBuffer`, then restores the content the renderer's serialization
   * drops (un-modeled graphic frames, extension lists) from the prior buffer --
   * the same preservation the `applyListStyle` funnel relies on. Cheap no-op when
   * nothing is pending, so callers can guard any `currentBuffer` read with it.
   */
  private async syncCurrentBuffer(): Promise<void> {
    if (this.pendingSlideXml.size === 0) return;

    const pending = this.pendingSlideXml;
    this.pendingSlideXml = new Map();

    const previousBuffer = this.currentBuffer;
    let buffer = await buildZip(
      previousBuffer,
      new Map(Array.from(pending, ([slideIndex, xml]) => [getSlidePath(slideIndex), xml]))
    );
    for (const slideIndex of pending.keys()) {
      buffer = await mergeSlideGraphicFramesFromBuffer(previousBuffer, buffer, slideIndex);
    }
    buffer = await preserveSlideExtensionLists(previousBuffer, buffer);
    this.currentBuffer = buffer;
  }

  private async refreshSlideBackgroundImages(buffer: ArrayBuffer): Promise<void> {
    const zip = await extractZip(buffer);
    const backgroundHrefs = new Map<number, string>();

    for (let slideIndex = 0; slideIndex < this.slideCountValue; slideIndex++) {
      const href = this.resolveSlideBackgroundImageHref(zip, slideIndex);
      if (href) backgroundHrefs.set(slideIndex, href);
    }

    this.slideBackgroundImageHrefs = backgroundHrefs;
  }

  private resolveSlideBackgroundImageHref(zip: ZipContents, slideIndex: number): string | null {
    const slidePath = getSlidePath(slideIndex);
    const slideBackground = this.resolvePartBackgroundImageHref(zip, slidePath);
    if (slideBackground.explicit) return slideBackground.href;

    const layoutPath = this.resolveRelationshipTargetByType(zip, slidePath, SLIDE_LAYOUT_RELATIONSHIP_TYPE);
    if (!layoutPath) return null;

    const layoutBackground = this.resolvePartBackgroundImageHref(zip, layoutPath);
    if (layoutBackground.explicit) return layoutBackground.href;

    const masterPath = this.resolveRelationshipTargetByType(zip, layoutPath, SLIDE_MASTER_RELATIONSHIP_TYPE);
    if (!masterPath) return null;

    const masterBackground = this.resolvePartBackgroundImageHref(zip, masterPath);
    return masterBackground.href;
  }

  private resolvePartBackgroundImageHref(zip: ZipContents, partPath: string): { explicit: boolean; href: string | null } {
    const xml = zip.textFiles.get(partPath);
    if (!xml) return { explicit: false, href: null };

    const doc = parseXml(xml, partPath);
    const background = getDescendants(doc, 'bg')[0];
    if (!background) return { explicit: false, href: null };

    const blip = getDescendants(background, 'blip')[0];
    if (!blip) return { explicit: true, href: null };

    const relationshipId = getBlipEmbedId(blip);
    if (!relationshipId) return { explicit: true, href: null };

    const imagePath = this.resolveRelationshipTargetById(zip, partPath, relationshipId, IMAGE_RELATIONSHIP_TYPE);
    if (!imagePath) return { explicit: true, href: null };

    const imageBytes = zip.binaryFiles.get(imagePath);
    if (!imageBytes) return { explicit: true, href: null };

    const mimeType = contentTypeForImageExtension(getPartExtension(imagePath));
    return {
      explicit: true,
      href: `data:${mimeType};base64,${bytesToBase64(imageBytes)}`
    };
  }

  private resolveRelationshipTargetByType(zip: ZipContents, sourcePath: string, relationshipType: string): string | null {
    const relationship = this.findRelationship(zip, sourcePath, (candidate) =>
      candidate.getAttribute('Type') === relationshipType
    );
    return this.resolveRelationshipTarget(sourcePath, relationship);
  }

  private resolveRelationshipTargetById(
    zip: ZipContents,
    sourcePath: string,
    relationshipId: string,
    relationshipType: string
  ): string | null {
    const relationship = this.findRelationship(zip, sourcePath, (candidate) =>
      candidate.getAttribute('Id') === relationshipId
        && candidate.getAttribute('Type') === relationshipType
    );
    return this.resolveRelationshipTarget(sourcePath, relationship);
  }

  private findRelationship(
    zip: ZipContents,
    sourcePath: string,
    predicate: (relationship: Element) => boolean
  ): Element | null {
    const relationshipsPath = this.getRelationshipsPathForPart(sourcePath);
    const relationshipsXml = zip.textFiles.get(relationshipsPath);
    if (!relationshipsXml) return null;

    const relationshipsDoc = parseXml(relationshipsXml, relationshipsPath);
    return getDescendants(relationshipsDoc, 'Relationship')
      .find((relationship) =>
        relationship.getAttribute('TargetMode') !== 'External' && predicate(relationship)
      ) ?? null;
  }

  private resolveRelationshipTarget(sourcePath: string, relationship: Element | null): string | null {
    const target = relationship?.getAttribute('Target');
    if (!target) return null;

    const normalized = target.replace(/\\/g, '/');
    if (normalized.startsWith('/')) return normalized.slice(1);
    return resolvePartPath(sourcePath, normalized);
  }

  private getRelationshipsPathForPart(partPath: string): string {
    const slashIndex = partPath.lastIndexOf('/');
    const directory = slashIndex >= 0 ? partPath.slice(0, slashIndex + 1) : '';
    const fileName = slashIndex >= 0 ? partPath.slice(slashIndex + 1) : partPath;
    return `${directory}_rels/${fileName}.rels`;
  }

  private async refreshChartTextValues(buffer: ArrayBuffer): Promise<void> {
    const zip = await extractZip(buffer);
    const chartTextValues = new Map<string, string[]>();
    const chartAxisFormats = new Map<string, ChartAxisFormat[]>();
    const chartDataDescriptors = new Map<string, ChartDataDescriptor>();

    for (let slideIndex = 0; slideIndex < this.slideCountValue; slideIndex++) {
      const slidePath = getSlidePath(slideIndex);
      const slideXml = zip.textFiles.get(slidePath);
      if (!slideXml) continue;

      const shapes = getElementChildren(getDescendants(parseXml(slideXml, slidePath), 'spTree')[0])
        .filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));
      shapes.forEach((shape, shapeIndex) => {
        if (!getDescendants(shape, 'chart')[0]) return;
        try {
          const chartPath = findChartPartPath(zip.textFiles, slideIndex, shapeIndex);
          chartTextValues.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartTextValues(zip.textFiles, slideIndex, shapeIndex)
          );
          chartAxisFormats.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartAxisFormats(zip.textFiles, slideIndex, shapeIndex)
          );
          chartDataDescriptors.set(
            this.getChartTextKey(slideIndex, shapeIndex),
            getChartDataDescriptor(zip, chartPath)
          );
        } catch {
          // Unsupported chart variants remain viewable and read-only.
        }
      });
    }

    this.chartTextValues = chartTextValues;
    this.chartAxisFormats = chartAxisFormats;
    this.chartDataDescriptors = chartDataDescriptors;
  }

  private getChartTextKey(slideIndex: number, shapeIndex: number): string {
    return `${slideIndex}:${shapeIndex}`;
  }
}
