import {
  DRAWINGML_NAMESPACE,
  SHAPE_ELEMENT_NAMES,
  getDescendants,
  getElementChildren,
} from './ooxmlXml';
import {
  RUN_PROPERTY_CHILD_ORDER,
  applyRunStyleToParagraphRange,
  getDrawingParagraphs,
  getDrawingRunText,
  getDrawingRuns,
  getRunProperties,
  normalizeHexColor,
} from './drawingmlText';

/**
 * A single run that carries an <a:highlight> color, located by the same
 * shape/paragraph/run indices the SVG renderer tags onto its tspans. The SVG
 * renderer drops <a:highlight>, so the view repaints these as overlay rects.
 */
export interface RunHighlightInfo {
  shapeIndex: number;
  paragraphIndex: number;
  runIndex: number;
  color: string;
  /** Authoritative run-only, paragraph-relative offsets straight from the OOXML. */
  start: number;
  end: number;
}

/**
 * A run's authored `<a:rPr>`, captured while the renderer model is lossless so
 * properties it doesn't model can be re-grafted after a lossy re-serialize.
 * `rPr` is a detached clone; `text` is used to verify the run still lines up.
 */
export interface AuthoredRunRpr {
  shapeIndex: number;
  paragraphIndex: number;
  runIndex: number;
  text: string;
  rPr: Element;
}

export interface SlideRunCacheEntry {
  highlights: RunHighlightInfo[];
  authoredRuns: AuthoredRunRpr[];
}

// rPr attributes the renderer is known to preserve. A run is only cached when it
// carries an element child or an attribute outside this set, so the cache stays
// small (plain runs are skipped) while still catching un-modeled attributes.
const PRESERVED_RPR_ATTRS = new Set([
  'kumimoji', 'lang', 'altLang', 'sz', 'b', 'i', 'u', 'strike', 'kern', 'cap',
  'spc', 'normalizeH', 'baseline', 'noProof', 'dirty', 'err', 'smtClean', 'smtId', 'bmk'
]);

function rprChildOrderIndex(localName: string): number {
  const index = RUN_PROPERTY_CHILD_ORDER.indexOf(localName);
  return index === -1 ? RUN_PROPERTY_CHILD_ORDER.length : index;
}

/** Insert a child into an `<a:rPr>` at its schema-ordered position. */
function insertRprChildInOrder(rPr: Element, child: Element): void {
  const order = rprChildOrderIndex(child.localName);
  const successor = getElementChildren(rPr).find((existing) => rprChildOrderIndex(existing.localName) > order);
  rPr.insertBefore(child, successor ?? null);
}

/**
 * Collect every run carrying an `<a:highlight>` color from a parsed slide
 * document. Shape/paragraph/run indices and run-relative offsets match
 * `PresentationEngine.getRunStyle` (and the renderer's tspan tags).
 */
export function collectSlideHighlights(slideDoc: XMLDocument): RunHighlightInfo[] {
  const highlights: RunHighlightInfo[] = [];
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) return highlights;
  const shapes = getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));

  shapes.forEach((shape, shapeIndex) => {
    getDrawingParagraphs(shape).forEach((paragraph, paragraphIndex) => {
      let offset = 0;
      getDrawingRuns(paragraph).forEach((run, runIndex) => {
        const text = getDrawingRunText(run);
        const start = offset;
        const end = offset + text.length;
        offset = end;

        const rPr = getElementChildren(run)
          .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
        if (!rPr) return;
        const highlight = getElementChildren(rPr)
          .find((element) => element.localName === 'highlight' && element.namespaceURI === DRAWINGML_NAMESPACE);
        if (!highlight) return;
        const srgb = getElementChildren(highlight).find((element) => element.localName === 'srgbClr');
        const color = srgb?.getAttribute('val');
        if (!color) return;
        highlights.push({ shapeIndex, paragraphIndex, runIndex, color: normalizeHexColor(color), start, end });
      });
    });
  });
  return highlights;
}

/**
 * Capture the authored `<a:rPr>` of every "decorated" run (one with element
 * children or a non-core attribute) so the renderer's lossy re-serialize can be
 * reconciled later. Indices match {@link collectSlideHighlights}.
 */
export function collectAuthoredRunRprs(slideDoc: XMLDocument): AuthoredRunRpr[] {
  const authored: AuthoredRunRpr[] = [];
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) return authored;
  const shapes = getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));

  shapes.forEach((shape, shapeIndex) => {
    getDrawingParagraphs(shape).forEach((paragraph, paragraphIndex) => {
      getDrawingRuns(paragraph).forEach((run, runIndex) => {
        const rPr = getElementChildren(run)
          .find((element) => element.localName === 'rPr' && element.namespaceURI === DRAWINGML_NAMESPACE);
        if (!rPr) return;
        const hasChild = getElementChildren(rPr).length > 0;
        const hasUnmodeledAttr = Array.from(rPr.attributes).some((attr) => !PRESERVED_RPR_ATTRS.has(attr.name));
        if (!hasChild && !hasUnmodeledAttr) return;
        authored.push({
          shapeIndex,
          paragraphIndex,
          runIndex,
          text: getDrawingRunText(run),
          rPr: rPr.cloneNode(true) as Element
        });
      });
    });
  });
  return authored;
}

export function readSlideRunCacheFromDoc(slideDoc: XMLDocument): SlideRunCacheEntry {
  return { highlights: collectSlideHighlights(slideDoc), authoredRuns: collectAuthoredRunRprs(slideDoc) };
}

/**
 * Graft the given highlights into a slide document via the lossless OOXML path.
 * Shape/paragraph indices are resolved against the document's shape tree, so
 * the caller must pass highlights whose indices match `slideDoc`'s shape order.
 * Returns whether anything changed.
 */
export function graftHighlightsIntoSlideDoc(slideDoc: XMLDocument, highlights: RunHighlightInfo[]): boolean {
  if (highlights.length === 0) return false;
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) return false;
  const shapes = getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));

  let changed = false;
  for (const highlight of highlights) {
    const shape = shapes[highlight.shapeIndex];
    if (!shape) continue;
    const paragraph = getDrawingParagraphs(shape)[highlight.paragraphIndex];
    if (!paragraph) continue;
    if (applyRunStyleToParagraphRange(paragraph, slideDoc, highlight.start, highlight.end, { highlight: highlight.color })) {
      changed = true;
    }
  }
  return changed;
}

/**
 * Re-graft any authored run property the renderer dropped (everything except
 * `<a:highlight>`, which the offset-based highlight path owns) back into a
 * slide document. Runs are matched 1:1 by index and verified by text, then
 * each missing rPr child/attribute is re-inserted in schema order. This is a
 * no-op when nothing is missing, so the lossless path is untouched.
 */
export function graftAuthoredRunPropsIntoSlideDoc(slideDoc: XMLDocument, authoredRuns: AuthoredRunRpr[]): boolean {
  if (authoredRuns.length === 0) return false;
  const shapeTree = getDescendants(slideDoc, 'spTree')[0];
  if (!shapeTree) return false;
  const shapes = getElementChildren(shapeTree).filter((element) => SHAPE_ELEMENT_NAMES.has(element.localName));

  let changed = false;
  for (const authored of authoredRuns) {
    const shape = shapes[authored.shapeIndex];
    if (!shape) continue;
    const paragraph = getDrawingParagraphs(shape)[authored.paragraphIndex];
    if (!paragraph) continue;
    const run = getDrawingRuns(paragraph)[authored.runIndex];
    if (!run || getDrawingRunText(run) !== authored.text) continue;

    const rPr = getRunProperties(run, slideDoc);
    for (const attr of Array.from(authored.rPr.attributes)) {
      if (!rPr.hasAttribute(attr.name)) {
        rPr.setAttribute(attr.name, attr.value);
        changed = true;
      }
    }
    for (const child of getElementChildren(authored.rPr)) {
      if (child.localName === 'highlight') continue;
      const present = getElementChildren(rPr)
        .some((existing) => existing.localName === child.localName && existing.namespaceURI === child.namespaceURI);
      if (present) continue;
      insertRprChildInOrder(rPr, slideDoc.importNode(child, true));
      changed = true;
    }
  }
  return changed;
}

/** True when the slide document is missing cached highlights the renderer dropped. */
export function needsHighlightRegraft(slideDoc: XMLDocument, cachedHighlights: RunHighlightInfo[]): boolean {
  return cachedHighlights.length > 0 && collectSlideHighlights(slideDoc).length < cachedHighlights.length;
}

/**
 * After an in-place `deleteShape`, the renderer renumbers the surviving shapes.
 * Remap the cached run formatting to match: drop the deleted shape's entries
 * and shift every higher shape index down by one.
 */
export function remapSlideRunCacheAfterDeletedShape(
  cached: SlideRunCacheEntry,
  deletedShapeIndex: number
): SlideRunCacheEntry {
  const remapShape = <T extends { shapeIndex: number }>(entry: T): T =>
    entry.shapeIndex > deletedShapeIndex ? { ...entry, shapeIndex: entry.shapeIndex - 1 } : entry;
  return {
    highlights: cached.highlights.filter((h) => h.shapeIndex !== deletedShapeIndex).map(remapShape),
    authoredRuns: cached.authoredRuns.filter((r) => r.shapeIndex !== deletedShapeIndex).map(remapShape)
  };
}

/**
 * Re-graft run properties the renderer dropped into `slideDoc` from a cached
 * entry. When the model is still lossless the document already has them, so
 * this is a no-op and the common path stays byte-identical to before.
 */
export function restoreLostRunPropsIntoSlideDoc(slideDoc: XMLDocument, cached: SlideRunCacheEntry): void {
  if (needsHighlightRegraft(slideDoc, cached.highlights)) {
    graftHighlightsIntoSlideDoc(slideDoc, cached.highlights);
  }
  graftAuthoredRunPropsIntoSlideDoc(slideDoc, cached.authoredRuns);
}
