import {
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
import {
  getChartDataDescriptor,
  updateChartTextLabel as patchChartTextLabel,
  updateChartData as patchChartData,
  type ChartDataDescriptor,
  type ChartDataGrid,
  type ChartDataUpdate
} from './ChartData';
import { type FontSubstitution } from './FontFidelity';
import { debugLog, errorLog } from './logger';
import {
  type PptxRendererBackend,
} from './powerpoint/backend/rendererBackend';
import { PptxPackageDocument } from './powerpoint/document/PptxPackageDocument';
import {
  createSlideObjectClipboard,
  pasteSlideObject,
  type SlideObjectClipboard
} from './ShapeClipboard';
import {
  applyParagraphListStyle,
  insertChartIntoPresentation,
  insertShapeIntoPresentation,
  insertTableIntoPresentation,
  insertTextBoxIntoPresentation,
  mergeMissingPackageParts,
  mergeSlideGraphicFramesFromBuffer,
  permuteSlidesInBuffer,
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
export type { PptxRendererBackend } from './powerpoint/backend/rendererBackend';
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
const OOXML_PCT_MAX = 100000;

interface SlideBackgroundImage {
  href: string;
  crop: ImageCrop | null;
}

function readSrcRectCrop(srcRect: Element | null): ImageCrop | null {
  if (!srcRect) return null;

  const read = (attribute: string): number => {
    const value = Number(srcRect.getAttribute(attribute));
    return Number.isFinite(value) ? value / 1000 : 0;
  };

  const crop: ImageCrop = {
    left: read('l'),
    top: read('t'),
    right: read('r'),
    bottom: read('b')
  };
  return crop.left > 0 || crop.top > 0 || crop.right > 0 || crop.bottom > 0 ? crop : null;
}

function hasImageCrop(crop: ImageCrop | null): crop is ImageCrop {
  return crop !== null;
}

function computeCroppedImageGeometry(
  frameX: number,
  frameY: number,
  frameWidth: number,
  frameHeight: number,
  crop: ImageCrop
): { x: number; y: number; width: number; height: number } {
  const srcL = cropPercentToPermille(crop.left);
  const srcT = cropPercentToPermille(crop.top);
  const srcR = cropPercentToPermille(crop.right);
  const srcB = cropPercentToPermille(crop.bottom);
  const visibleWidth = OOXML_PCT_MAX - srcL - srcR;
  const visibleHeight = OOXML_PCT_MAX - srcT - srcB;
  const width = visibleWidth > 0 ? (frameWidth * OOXML_PCT_MAX) / visibleWidth : frameWidth;
  const height = visibleHeight > 0 ? (frameHeight * OOXML_PCT_MAX) / visibleHeight : frameHeight;
  return {
    x: frameX - (width * srcL) / OOXML_PCT_MAX,
    y: frameY - (height * srcT) / OOXML_PCT_MAX,
    width,
    height
  };
}

function findFullBleedSlideBackgroundImage(root: Element, width: string, height: string): Element | null {
  return getElementChildren(root).find((child) =>
    child.localName === 'image'
      && child.getAttribute('x') === '0'
      && child.getAttribute('y') === '0'
      && child.getAttribute('width') === width
      && child.getAttribute('height') === height
      && child.getAttribute('preserveAspectRatio') === 'none'
      && !child.getAttribute('clip-path')
  ) ?? null;
}

function findCroppedSlideBackgroundImage(root: Element, slideIndex: number): Element | null {
  const clipId = `bgclip-s${slideIndex + 1}`;
  return getElementChildren(root).find((child) =>
    child.localName === 'image'
      && child.getAttribute('clip-path') === `url(#${clipId})`
  ) ?? null;
}

function getSlideBackgroundInsertBefore(root: Element): ChildNode | null {
  const firstChild = root.firstChild;
  return firstChild?.nodeType === 1
    && (firstChild as Element).localName === 'rect'
    && (firstChild as Element).getAttribute('fill') === 'none'
    ? firstChild.nextSibling
    : firstChild;
}

function insertRenderedSlideBackgroundImage(
  doc: XMLDocument,
  root: Element,
  slideIndex: number,
  width: string,
  height: string,
  background: SlideBackgroundImage
): void {
  const frameWidth = Number(width);
  const frameHeight = Number(height);
  if (!Number.isFinite(frameWidth) || !Number.isFinite(frameHeight) || frameWidth <= 0 || frameHeight <= 0) {
    return;
  }

  const insertBefore = getSlideBackgroundInsertBefore(root);

  if (hasImageCrop(background.crop)) {
    const clipId = `bgclip-s${slideIndex + 1}`;
    const geometry = computeCroppedImageGeometry(0, 0, frameWidth, frameHeight, background.crop);

    const defs = doc.createElementNS(SVG_NAMESPACE, 'defs');
    const clipPath = doc.createElementNS(SVG_NAMESPACE, 'clipPath');
    clipPath.setAttribute('id', clipId);
    const clipRect = doc.createElementNS(SVG_NAMESPACE, 'rect');
    clipRect.setAttribute('x', '0');
    clipRect.setAttribute('y', '0');
    clipRect.setAttribute('width', width);
    clipRect.setAttribute('height', height);
    clipPath.appendChild(clipRect);
    defs.appendChild(clipPath);
    root.insertBefore(defs, insertBefore);

    const image = doc.createElementNS(SVG_NAMESPACE, 'image');
    image.setAttribute('x', String(geometry.x));
    image.setAttribute('y', String(geometry.y));
    image.setAttribute('width', String(geometry.width));
    image.setAttribute('height', String(geometry.height));
    image.setAttribute('preserveAspectRatio', 'none');
    image.setAttribute('href', background.href);
    image.setAttribute('clip-path', `url(#${clipId})`);
    image.setAttribute('pointer-events', 'none');
    root.insertBefore(image, insertBefore);
    return;
  }

  const image = doc.createElementNS(SVG_NAMESPACE, 'image');
  image.setAttribute('x', '0');
  image.setAttribute('y', '0');
  image.setAttribute('width', width);
  image.setAttribute('height', height);
  image.setAttribute('preserveAspectRatio', 'none');
  image.setAttribute('href', background.href);
  image.setAttribute('pointer-events', 'none');
  root.insertBefore(image, insertBefore);
}

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

/**
 * Renderer build-patched with the single-slide entry point (see
 * scripts/lib/patch-pptx-renderer.mjs). Replaces one slide's XML in the live
 * file map and re-parses, so edits skip the whole-deck export/reload.
 */
interface SlideXmlLoadable {
  loadSlideXml(slideIdx: number, xml: string): void;
}

import {
  configureForceJsBackendOverrideReader,
  resetForceJsBackendOverride,
  setForceJsBackendOverride
} from './powerpoint/forceJsBackend';

export { formatChartAxisValue };

export {
  configureForceJsBackendOverrideReader,
  resetForceJsBackendOverride,
  setForceJsBackendOverride
};

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

/** Shape-level fill and outline read from slide OOXML for agent describe snapshots. */
export interface ShapeVisualStyle {
  fill: string | null;
  stroke: string | null;
  strokeWidthPt: number | null;
}

/** Resolved slide background for agent describe snapshots. */
export interface SlideBackgroundDescribe {
  colorHex: string | null;
  imageHref: string | null;
  crop: ImageCrop | null;
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
  private document!: PptxPackageDocument;
  private chartTextValues = new Map<string, string[]>();
  private chartAxisFormats = new Map<string, ChartAxisFormat[]>();
  private chartDataDescriptors = new Map<string, ChartDataDescriptor>();
  private slideBackgroundImages = new Map<number, SlideBackgroundImage>();
  // Invariant: the lossless package buffer is authoritative; renderer state is derived.
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
  private constructor() {
  }

  static async load(buffer: ArrayBuffer): Promise<PresentationEngine> {
    const engine = new PresentationEngine();
    engine.document = await PptxPackageDocument.load(buffer, {
      reconcileExport: (authoritativePackage, renderedExport) =>
        engine.reconcileRendererExport(authoritativePackage, renderedExport),
      refreshDerivedState: (packageBuffer) => engine.refreshDerivedState(packageBuffer),
    });
    await engine.refreshDerivedState(buffer);
    return engine;
  }

  private get renderer() {
    return this.document.renderer;
  }

  private get fontFidelity() {
    return this.document.fontFidelity;
  }

  private get currentBuffer(): ArrayBuffer {
    return this.document.packageBuffer;
  }

  private set currentBuffer(buffer: ArrayBuffer) {
    this.document.packageBuffer = buffer;
  }

  private get slideCountValue(): number {
    return this.document.slideCount;
  }

  getRendererBackend(): PptxRendererBackend {
    return this.document.rendererBackend;
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

  getSlideXml(slideIndex: number): string {
    const slidePath = getSlidePath(slideIndex);
    const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), slidePath);
    this.restoreLostRunPropsIntoSlideDoc(slideIndex, slideDoc);
    return serializeXml(slideDoc);
  }

  async restoreSlideXml(slideIndex: number, xml: string): Promise<void> {
    const slideDoc = parseXml(xml, getSlidePath(slideIndex));
    await this.commitSlideDoc(slideIndex, slideDoc);
    this.refreshSlideRunCache(slideIndex);
  }

  private reconcileRenderedSlideBackground(svg: string, slideIndex: number): string {
    const background = this.slideBackgroundImages.get(slideIndex);
    if (!background?.href || !svg.startsWith('<svg')) return svg;

    try {
      const doc = parseXml(svg, `rendered slide ${slideIndex + 1} SVG`);
      const root = doc.documentElement;
      const width = root.getAttribute('width') ?? '';
      const height = root.getAttribute('height') ?? '';
      if (!width || !height) return svg;

      const croppedBackground = findCroppedSlideBackgroundImage(root, slideIndex);
      const fullBleedBackground = findFullBleedSlideBackgroundImage(root, width, height);

      if (croppedBackground) {
        croppedBackground.setAttribute('href', background.href);
        croppedBackground.setAttribute('pointer-events', 'none');
        if (fullBleedBackground && fullBleedBackground !== croppedBackground) {
          root.removeChild(fullBleedBackground);
        }
      } else if (fullBleedBackground) {
        fullBleedBackground.setAttribute('href', background.href);
        fullBleedBackground.setAttribute('pointer-events', 'none');
      } else {
        insertRenderedSlideBackgroundImage(doc, root, slideIndex, width, height, background);
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

  /**
   * Apply a transform to a shape that was just inserted through the OOXML-side
   * insertion funnel (`insertShapeGeometry`, `insertTextBox`, etc.). Uses the
   * OOXML export path instead of the renderer shortcut so unrelated slide markup
   * (for example `a14:hiddenFill` on pictures) is not corrupted when renderer
   * shape indices diverge from the serialized slide tree.
   */
  async applyInsertedShapeTransform(
    slideIndex: number,
    shapeIndex: number,
    transform: ShapeTransform
  ): Promise<void> {
    await this.updateShapeTransformInOoxml(slideIndex, shapeIndex, transform);
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
    await this.editSlideShape(slideIndex, shapeIndex, (shape) => {
      setDrawingParagraphText(shape, paragraphIndex, text);
      return true;
    });
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
    this.realignSlideRunCacheShapeIndices(slideIndex);
  }

  /**
   * After a renderer-side delete, cached shape indices can drift from the live
   * model when the slide tree has gaps (graphic frames, groups). Re-resolve each
   * cached highlight against the paragraph text that still exists in the model.
   */
  private realignSlideRunCacheShapeIndices(slideIndex: number): void {
    const cached = this.slideRunCache.get(slideIndex);
    if (!cached || cached.highlights.length === 0) return;

    const highlights = cached.highlights.map((highlight) => {
      if (highlight.end <= highlight.start) return highlight;
      let bestIndex = highlight.shapeIndex;
      let bestDistance = Number.POSITIVE_INFINITY;
      for (let candidate = 0; candidate < 64; candidate++) {
        const text = this.getParagraphRunText(slideIndex, candidate, highlight.paragraphIndex);
        if (!text || highlight.end > text.length) continue;
        if (text.slice(highlight.start, highlight.end).length !== highlight.end - highlight.start) continue;
        const distance = Math.abs(candidate - highlight.shapeIndex);
        if (distance < bestDistance) {
          bestDistance = distance;
          bestIndex = candidate;
        }
      }
      return { ...highlight, shapeIndex: bestIndex };
    });

    this.slideRunCache.set(slideIndex, { ...cached, highlights });
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

  getTextRunText(
    slideIndex: number,
    shapeIndex: number,
    paragraphIndex: number,
    runIndex: number,
  ): string | null {
    let shape: Element;
    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      shape = getShapeElement(slideDoc, shapeIndex);
    } catch {
      return null;
    }

    const paragraph = getDrawingParagraphs(shape)[paragraphIndex];
    if (!paragraph) return null;

    const run = getDrawingRuns(paragraph)[runIndex];
    if (!run) return null;

    return getDrawingRunText(run);
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

  private readFillColorFromNode(fillNode: Element | undefined): string | null {
    if (!fillNode) return null;
    const srgb = getElementChildren(fillNode).find((element) => element.localName === 'srgbClr');
    return this.readColorValue(srgb);
  }

  private readShapeVisualStyle(properties: Element): ShapeVisualStyle {
    const solidFill = getElementChildren(properties).find((element) => element.localName === 'solidFill');
    const line = getElementChildren(properties).find((element) => element.localName === 'ln');
    const lineFill = line
      ? getElementChildren(line).find((element) => element.localName === 'solidFill')
      : undefined;
    const strokeWidth = line?.getAttribute('w');
    const parsedStrokeWidth = strokeWidth ? Number(strokeWidth) : Number.NaN;

    return {
      fill: this.readFillColorFromNode(solidFill),
      stroke: this.readFillColorFromNode(lineFill),
      strokeWidthPt: Number.isFinite(parsedStrokeWidth) ? parsedStrokeWidth / 12700 : null,
    };
  }

  getShapeVisualStyle(slideIndex: number, shapeIndex: number): ShapeVisualStyle | null {
    if (!Number.isInteger(shapeIndex) || shapeIndex < 0) {
      return null;
    }

    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElement(slideDoc, shapeIndex);
      const properties = getDescendants(shape, 'spPr')[0]
        ?? getDescendants(shape, 'grpSpPr')[0]
        ?? getDescendants(shape, 'picSpPr')[0];
      if (!properties) {
        return { fill: null, stroke: null, strokeWidthPt: null };
      }
      return this.readShapeVisualStyle(properties);
    } catch {
      return null;
    }
  }

  /** Whether the selected slide object owns ordinary shape properties with a fill. */
  canSetShapeFillColor(slideIndex: number, shapeIndex: number): boolean {
    if (!Number.isInteger(shapeIndex) || shapeIndex < 0) return false;

    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElementByRendererIndex(slideDoc, shapeIndex);
      return shape.localName === 'sp'
        && getElementChildren(shape).some((element) => element.localName === 'spPr');
    } catch {
      return false;
    }
  }

  /** Detect a freeform text box without conflating it with an auto shape that merely has a label. */
  isTextBoxShape(slideIndex: number, shapeIndex: number): boolean {
    if (!Number.isInteger(shapeIndex) || shapeIndex < 0) return false;

    try {
      const slideDoc = parseXml(this.renderer.getSlideOoxml(slideIndex), getSlidePath(slideIndex));
      const shape = getShapeElementByRendererIndex(slideDoc, shapeIndex);
      if (shape.localName !== 'sp') return false;

      const nonVisualProperties = getDescendants(shape, 'cNvSpPr')[0];
      const textBoxFlag = nonVisualProperties?.getAttribute('txBox');
      if (textBoxFlag === '1' || textBoxFlag === 'true') return true;

      // Text boxes inserted by older plugin builds predate the txBox marker,
      // but consistently use the standard non-visual TextBox name.
      const name = getDescendants(shape, 'cNvPr')[0]?.getAttribute('name') ?? '';
      return /^TextBox(?:\s|$)/i.test(name);
    } catch {
      return false;
    }
  }

  /** Set an explicit solid fill on an ordinary auto shape or text box. */
  async setShapeFillColor(slideIndex: number, shapeIndex: number, hex: string): Promise<void> {
    const normalizedHex = hex.replace(/^#/, '').trim().toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(normalizedHex)) {
      throw new Error('Shape fill color must be a 6-digit RRGGBB hex value.');
    }

    await this.editSlideShape(slideIndex, shapeIndex, (shape, slideDoc) => {
      if (shape.localName !== 'sp') {
        throw new Error('The selected object does not support a fill color.');
      }

      const properties = getElementChildren(shape).find((element) => element.localName === 'spPr');
      if (!properties) {
        throw new Error('The selected object does not support a fill color.');
      }

      const fillNames = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
      const fills = getElementChildren(properties).filter((element) => fillNames.has(element.localName));
      if (
        fills.length === 1
        && fills[0]?.localName === 'solidFill'
        && this.readFillColorFromNode(fills[0]) === normalizedHex
      ) {
        return false;
      }
      for (const fill of fills) properties.removeChild(fill);

      const solidFill = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:solidFill');
      const color = slideDoc.createElementNS(DRAWINGML_NAMESPACE, 'a:srgbClr');
      color.setAttribute('val', normalizedHex);
      solidFill.appendChild(color);

      const insertionPoint = getElementChildren(properties).find((element) =>
        ['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].includes(element.localName)
      );
      properties.insertBefore(solidFill, insertionPoint ?? null);
      return true;
    });
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
    const slidePath = getSlidePath(slideIndex);
    const startedAt = Date.now();
    debugLog('mutate', 'Slide document commit started', {
      op: 'commit-slide-doc',
      slide: slideIndex,
      path: slidePath,
      authoritativePackage: true,
    });
    try {
      const loader = this.renderer as Partial<SlideXmlLoadable>;
      if (typeof loader.loadSlideXml === 'function') {
        loader.loadSlideXml(slideIndex, serialized);
        // The renderer now holds this slide; `currentBuffer` is behind for it.
        // Record the lossless XML so a later `currentBuffer` reader can fold it in
        // (preserving renderer-dropped content) instead of reading stale slide XML.
        this.document.recordPendingSlideXml(slideIndex, serialized);
        debugLog('mutate', 'Slide document commit completed', {
          op: 'commit-slide-doc',
          slide: slideIndex,
          path: slidePath,
          pathType: 'pending-lossless-package',
          ms: Date.now() - startedAt,
        });
        return;
      }

      const rawExport = await this.exportRendererState();
      const patchedExport = await buildZip(rawExport, new Map([[slidePath, serialized]]));
      await this.reloadFromBuffer(patchedExport, this.slideCountValue);
      debugLog('mutate', 'Slide document commit completed', {
        op: 'commit-slide-doc',
        slide: slideIndex,
        path: slidePath,
        pathType: 'package-reload',
        ms: Date.now() - startedAt,
      });
    } catch (error) {
      errorLog('mutate', 'Slide document commit failed', {
        op: 'commit-slide-doc',
        slide: slideIndex,
        path: slidePath,
        error,
      });
      throw error;
    }
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
  ): Promise<number> {
    return this.insertImage(slideIndex, imageData, mimeType, widthPx, heightPx);
  }

  async insertImage(
    slideIndex: number,
    imageData: Uint8Array,
    mimeType: string,
    widthPx = 320,
    heightPx = 240
  ): Promise<number> {
    await this.syncCurrentBuffer();
    this.ensureSlideRunCacheSeeded(slideIndex);
    const { x, y, cx, cy } = await this.getDefaultInsertExtents('image', widthPx, heightPx);
    const result = this.renderer.addImage(slideIndex, imageData, mimeType, x, y, cx, cy);
    const shapeIndex = this.parseInsertedShapeIndex(result, 'Could not insert image.');
    await this.exportRendererState();
    return shapeIndex;
  }

  addShapeGeometry(slideIndex: number, geometry: InsertableShapeGeometry): Promise<number> {
    return this.insertShapeGeometry(slideIndex, geometry);
  }

  async insertShapeGeometry(slideIndex: number, geometry: InsertableShapeGeometry): Promise<number> {
    await this.syncCurrentBuffer();
    this.ensureSlideRunCacheSeeded(slideIndex);
    const { x, y, cx, cy } = await this.getDefaultInsertExtents(geometry);
    const inserted = await insertShapeIntoPresentation(
      this.currentBuffer,
      slideIndex,
      geometry,
      x,
      y,
      cx,
      cy,
      { red: 66, green: 133, blue: 244 },
    );
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    return inserted.shapeIndex;
  }

  addTextBox(slideIndex: number): Promise<number> {
    return this.insertTextBox(slideIndex);
  }

  async insertTextBox(slideIndex: number): Promise<number> {
    await this.syncCurrentBuffer();
    this.ensureSlideRunCacheSeeded(slideIndex);
    const { x, y, cx, cy } = await this.getDefaultInsertExtents('textBox');
    const inserted = await insertTextBoxIntoPresentation(
      this.currentBuffer,
      slideIndex,
      'New text',
      x,
      y,
      cx,
      cy,
    );
    await this.reloadFromBuffer(inserted.buffer, this.slideCountValue);
    return inserted.shapeIndex;
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
    // highlights (drop this shape, shift higher indices down, then re-resolve
    // against the live model) to keep overlays and export reconciliation aligned.
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
    const slidePath = getSlidePath(slideIndex);
    const startedAt = Date.now();
    debugLog('mutate', 'Slide tree mutation started', {
      op: 'mutate-slide-tree',
      slide: slideIndex,
      path: slidePath,
      authoritativePackage: true,
    });
    try {
      const rawExport = await this.exportRendererState();
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
      debugLog('mutate', 'Slide tree mutation committed', {
        op: 'mutate-slide-tree',
        slide: slideIndex,
        path: slidePath,
        ms: Date.now() - startedAt,
      });
      return result;
    } catch (error) {
      errorLog('mutate', 'Slide tree mutation failed', {
        op: 'mutate-slide-tree',
        slide: slideIndex,
        path: slidePath,
        error,
      });
      throw error;
    }
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
    this.currentBuffer = await permuteSlidesInBuffer(this.currentBuffer, order);
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
    this.currentBuffer = await permuteSlidesInBuffer(this.currentBuffer, newOrder);
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

  private static readonly EMU_PER_INCH = 914400;

  private async getDefaultInsertExtents(
    kind: InsertableShapeGeometry | 'textBox' | 'image',
    widthPx = 320,
    heightPx = 240
  ): Promise<{ x: number; y: number; cx: number; cy: number }> {
    const slideSize = await this.getSlideSizeEmu();
    let cx: number;
    let cy: number;

    if (kind === 'image') {
      cx = pxToEmu(widthPx);
      cy = pxToEmu(heightPx);
    } else if (kind === 'textBox') {
      cx = Math.round(3 * PresentationEngine.EMU_PER_INCH);
      cy = Math.round(0.75 * PresentationEngine.EMU_PER_INCH);
    } else if (kind === 'line') {
      cx = Math.round(2 * PresentationEngine.EMU_PER_INCH);
      cy = 0;
    } else if (kind === 'rightArrow' || kind === 'leftArrow' || kind === 'upArrow' || kind === 'downArrow') {
      cx = Math.round(2 * PresentationEngine.EMU_PER_INCH);
      cy = Math.round(1 * PresentationEngine.EMU_PER_INCH);
    } else {
      cx = Math.round(2 * PresentationEngine.EMU_PER_INCH);
      cy = Math.round(1.5 * PresentationEngine.EMU_PER_INCH);
    }

    const x = Math.round((slideSize.cx - cx) / 2);
    const y = Math.round((slideSize.cy - cy) / 2);
    return { x, y, cx, cy };
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

  getSlideBackgroundDescribe(slideIndex: number): SlideBackgroundDescribe {
    const image = this.slideBackgroundImages.get(slideIndex);
    return {
      colorHex: this.getSlideBackgroundColor(slideIndex),
      imageHref: image?.href ?? null,
      crop: image?.crop ?? null,
    };
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

  /**
   * Finalize a command transaction. The package document reconciles the
   * renderer's derived state with the lossless package and drains pending
   * slide-XML commits before the session is allowed to become dirty.
   */
  async commitMutation(): Promise<void> {
    await this.exportRendererState();
  }

  async restoreSnapshot(buffer: ArrayBuffer): Promise<void> {
    await this.document.restore(buffer);
    this.resetSlideRunCache();
  }

  private async reloadAfterSlideManagement(expectedSlideCount: number): Promise<void> {
    const rawExport = await this.exportRendererState();
    const normalizedExport = await normalizeSlideManifest(rawExport, expectedSlideCount);
    await this.reloadFromBuffer(normalizedExport, expectedSlideCount);
  }

  private async reloadFromBuffer(buffer: ArrayBuffer, expectedSlideCount: number): Promise<void> {
    // The reinitialized model serves lossless highlights again; drop the stale
    // cache (shape indices may have shifted) and let it re-seed on demand.
    this.resetSlideRunCache();
    await this.document.reload(buffer, expectedSlideCount);
  }

  private async exportRendererState(): Promise<ArrayBuffer> {
    return this.document.export();
  }

  /**
   * The package document owns export/reload state and calls this reconciliation
   * seam before accepting renderer output as authoritative OOXML.
   */
  private async reconcileRendererExport(
    authoritativePackage: ArrayBuffer,
    renderedExport: ArrayBuffer
  ): Promise<ArrayBuffer> {
    let mergedExport = renderedExport;
    for (let slideIndex = 0; slideIndex < this.slideCountValue; slideIndex++) {
      mergedExport = await mergeSlideGraphicFramesFromBuffer(authoritativePackage, mergedExport, slideIndex);
    }
    const preservedExport = await preserveSlideExtensionLists(authoritativePackage, mergedExport);
    return this.reconcileRunPropsIntoBuffer(preservedExport);
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
    await this.document.syncPackageFromPendingSlides();
  }

  private async refreshDerivedState(buffer: ArrayBuffer): Promise<void> {
    await this.refreshChartTextValues(buffer);
    await this.refreshSlideBackgroundImages(buffer);
  }

  private async refreshSlideBackgroundImages(buffer: ArrayBuffer): Promise<void> {
    const zip = await extractZip(buffer);
    const backgrounds = new Map<number, SlideBackgroundImage>();

    for (let slideIndex = 0; slideIndex < this.slideCountValue; slideIndex++) {
      const background = this.resolveSlideBackgroundImage(zip, slideIndex);
      if (background) backgrounds.set(slideIndex, background);
    }

    this.slideBackgroundImages = backgrounds;
  }

  private resolveSlideBackgroundImage(zip: ZipContents, slideIndex: number): SlideBackgroundImage | null {
    const slidePath = getSlidePath(slideIndex);
    const slideBackground = this.resolvePartBackgroundImage(zip, slidePath);
    if (slideBackground.image) return slideBackground.image;

    const layoutPath = this.resolveRelationshipTargetByType(zip, slidePath, SLIDE_LAYOUT_RELATIONSHIP_TYPE);
    if (!layoutPath) return null;

    const layoutBackground = this.resolvePartBackgroundImage(zip, layoutPath);
    if (layoutBackground.image) return layoutBackground.image;

    const masterPath = this.resolveRelationshipTargetByType(zip, layoutPath, SLIDE_MASTER_RELATIONSHIP_TYPE);
    if (!masterPath) return null;

    const masterBackground = this.resolvePartBackgroundImage(zip, masterPath);
    return masterBackground.image;
  }

  private resolvePartBackgroundImage(
    zip: ZipContents,
    partPath: string
  ): { explicit: boolean; image: SlideBackgroundImage | null } {
    const xml = zip.textFiles.get(partPath);
    if (!xml) return { explicit: false, image: null };

    const doc = parseXml(xml, partPath);
    const background = getDescendants(doc, 'bg')[0];
    if (!background) return { explicit: false, image: null };

    const blipFill = getDescendants(background, 'blipFill')[0];
    const blip = getDescendants(background, 'blip')[0];
    if (!blip) return { explicit: true, image: null };

    const relationshipId = getBlipEmbedId(blip);
    if (!relationshipId) return { explicit: true, image: null };

    const imagePath = this.resolveRelationshipTargetById(zip, partPath, relationshipId, IMAGE_RELATIONSHIP_TYPE);
    if (!imagePath) return { explicit: true, image: null };

    const imageBytes = zip.binaryFiles.get(imagePath);
    if (!imageBytes) return { explicit: true, image: null };

    const mimeType = contentTypeForImageExtension(getPartExtension(imagePath));
    const srcRect = blipFill
      ? getElementChildren(blipFill).find((element) => element.localName === 'srcRect') ?? null
      : null;
    return {
      explicit: true,
      image: {
        href: `data:${mimeType};base64,${bytesToBase64(imageBytes)}`,
        crop: readSrcRectCrop(srcRect)
      }
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
