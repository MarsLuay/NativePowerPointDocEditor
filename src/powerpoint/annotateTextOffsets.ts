// Stamp `data-ooxml-char-start/end` onto rendered run tspans after each slide
// (or single-shape) render. Extracted from NativePowerPointView so the stamping
// pipeline can be exercised headlessly against real engine SVG in Electron.

import {
  alignRunTspansToOoxml,
  type RunTspanAlignment,
  type RunTspanOffset,
} from './textUtils';

export interface StampableRunSpan {
  textContent: string | null;
  setAttribute(name: string, value: string): void;
  getAttribute(name: string): string | null;
}

export interface ParagraphStampResult {
  paragraphIndex: number;
  spanCount: number;
  editorLength: number;
  ooxmlLength: number;
  reconciled: boolean;
}

export interface ShapeStampResult {
  shapeIndex: number;
  paragraphs: ParagraphStampResult[];
}

/** True when `el` is a run or paragraph tspan in browser or xmldom SVG trees. */
export function isRunTspanElement(el: Element): boolean {
  const tag = (el.tagName || el.nodeName || '').toLowerCase();
  return tag === 'tspan' || tag.endsWith(':tspan');
}

/**
 * Group leaf run tspans by paragraph index inside a shape `<g>`. Mirrors the
 * DOM walk in {@link annotateShapeGroupTextOffsets}.
 */
export function collectRunSpansByParagraph(shapeGroup: Element): Map<number, Element[]> {
  const runsByParagraph = new Map<number, Element[]>();
  const spans = shapeGroup.querySelectorAll('tspan[data-ooxml-run-idx]');
  for (let i = 0, len = spans.length; i < len; i++) {
    const span = spans[i] as Element;
    if (!isRunTspanElement(span)) continue;
    const paraContainer = span.closest('tspan[data-ooxml-para-idx]');
    const paragraphIndex = Number(paraContainer?.getAttribute('data-ooxml-para-idx'));
    if (!Number.isFinite(paragraphIndex)) continue;
    const list = runsByParagraph.get(paragraphIndex) ?? [];
    list.push(span);
    runsByParagraph.set(paragraphIndex, list);
  }
  return runsByParagraph;
}

/**
 * Align run tspans to OOXML text and stamp `data-ooxml-char-start/end` onto each.
 * Returns the alignment tiles for downstream range mapping / verification.
 */
export function stampParagraphRunOffsets(
  runSpans: readonly StampableRunSpan[],
  ooxmlText: string,
): RunTspanAlignment {
  const alignment = alignRunTspansToOoxml(runSpans.map((span) => span.textContent || ''), ooxmlText);
  runSpans.forEach((span, index) => {
    const entry = alignment.spans[index];
    if (!entry) return;
    span.setAttribute('data-ooxml-char-start', String(entry.charStart));
    span.setAttribute('data-ooxml-char-end', String(entry.charEnd));
  });
  return alignment;
}

/**
 * Stamp every text paragraph inside one shape group. `getParagraphRunText` should
 * return the authoritative OOXML run text (what PresentationEngine exposes).
 */
export function annotateShapeGroupTextOffsets(
  shapeGroup: Element,
  getParagraphRunText: (paragraphIndex: number) => string | null,
): ShapeStampResult | null {
  const shapeIndex = Number(shapeGroup.getAttribute('data-ooxml-shape-idx'));
  if (!Number.isFinite(shapeIndex)) return null;

  const runsByParagraph = collectRunSpansByParagraph(shapeGroup);
  const paragraphs: ParagraphStampResult[] = [];

  for (const [paragraphIndex, runSpans] of runsByParagraph) {
    const ooxmlText = getParagraphRunText(paragraphIndex);
    if (ooxmlText === null) continue;
    const alignment = stampParagraphRunOffsets(runSpans, ooxmlText);
    paragraphs.push({
      paragraphIndex,
      spanCount: runSpans.length,
      editorLength: alignment.editorLength,
      ooxmlLength: alignment.ooxmlLength,
      reconciled: alignment.reconciled,
    });
  }

  return { shapeIndex, paragraphs };
}

/**
 * Stamp every shape on a rendered slide SVG. This is the DOM half of
 * `NativePowerPointView.annotateSlideTextOffsets`.
 */
export function annotateSlideTextOffsets(
  svg: Element,
  getParagraphRunText: (shapeIndex: number, paragraphIndex: number) => string | null,
): ShapeStampResult[] {
  const results: ShapeStampResult[] = [];
  const shapeGroups = svg.querySelectorAll('g[data-ooxml-shape-idx]');
  for (let i = 0, len = shapeGroups.length; i < len; i++) {
    const shapeGroup = shapeGroups[i] as Element;
    const shapeIndex = Number(shapeGroup.getAttribute('data-ooxml-shape-idx'));
    if (!Number.isFinite(shapeIndex)) continue;
    const stamped = annotateShapeGroupTextOffsets(shapeGroup, (paragraphIndex) =>
      getParagraphRunText(shapeIndex, paragraphIndex),
    );
    if (stamped) results.push(stamped);
  }
  return results;
}

/** Read stamped tiles back from a paragraph's run tspans (post-stamp verification). */
export function readStampedTiles(runSpans: readonly StampableRunSpan[]): RunTspanOffset[] | null {
  const tiles: RunTspanOffset[] = [];
  let cursor = 0;
  for (const span of runSpans) {
    const rawStart = span.getAttribute('data-ooxml-char-start');
    const rawEnd = span.getAttribute('data-ooxml-char-end');
    if (rawStart === null || rawEnd === null) return null;
    const length = (span.textContent || '').length;
    tiles.push({
      editorStart: cursor,
      editorEnd: cursor + length,
      charStart: Number(rawStart),
      charEnd: Number(rawEnd),
    });
    cursor += length;
  }
  return tiles;
}
