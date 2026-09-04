/**
 * Paragraph border grouping and flow insets.
 *
 * OOXML `w:space` is the gap between text and the stroke (ECMA-376 §17.3.1.5).
 * Word includes that gap plus the stroke width in the paragraph's layout
 * height so the next paragraph starts after the rule. The paginator must
 * reserve the same inset; the painter must draw the overlay inside it.
 */

import type { BorderStyle, FlowBlock, ParagraphBlock, ParagraphBorders } from './types';

export function bordersEqual(a?: BorderStyle, b?: BorderStyle): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.style === b.style && a.width === b.width && a.color === b.color;
}

/**
 * Adjacent paragraphs with identical pBdr form a group (ECMA-376 §17.3.1.24).
 */
export function bordersFormGroup(a?: ParagraphBorders, b?: ParagraphBorders): boolean {
  if (!a && !b) return false;
  if (!a || !b) return false;
  return (
    bordersEqual(a.top, b.top) &&
    bordersEqual(a.bottom, b.bottom) &&
    bordersEqual(a.left, b.left) &&
    bordersEqual(a.right, b.right) &&
    bordersEqual(a.between, b.between)
  );
}

export function isPaintedBorder(side?: BorderStyle): side is BorderStyle {
  return Boolean(side && side.style && side.style !== 'none' && side.style !== 'nil');
}

/** Vertical space Word keeps for one painted side: gap plus stroke. */
export function borderFlowSize(side?: BorderStyle): number {
  if (!isPaintedBorder(side)) return 0;
  return (side.space ?? 0) + (side.width ?? 0);
}

export function resolveRenderedParagraphBorders(
  borders: ParagraphBorders | undefined,
  prevBorders?: ParagraphBorders,
  nextBorders?: ParagraphBorders
): { top?: BorderStyle; bottom?: BorderStyle } {
  if (!borders) return {};
  const groupedWithPrev = bordersFormGroup(prevBorders, borders);
  const groupedWithNext = bordersFormGroup(borders, nextBorders);
  return {
    top: groupedWithPrev ? borders.between : borders.top,
    bottom: !groupedWithNext ? borders.bottom : undefined,
  };
}

export function paragraphBorderFlowInsets(
  borders: ParagraphBorders | undefined,
  prevBorders?: ParagraphBorders,
  nextBorders?: ParagraphBorders
): { top: number; bottom: number } {
  const rendered = resolveRenderedParagraphBorders(borders, prevBorders, nextBorders);
  return {
    top: borderFlowSize(rendered.top),
    bottom: borderFlowSize(rendered.bottom),
  };
}

export function adjacentParagraphBorders(
  blocks: FlowBlock[],
  index: number
): { prev?: ParagraphBorders; next?: ParagraphBorders } {
  const prevBlock = index > 0 ? blocks[index - 1] : undefined;
  const nextBlock = index + 1 < blocks.length ? blocks[index + 1] : undefined;
  return {
    prev: prevBlock?.kind === 'paragraph' ? (prevBlock as ParagraphBlock).attrs?.borders : undefined,
    next: nextBlock?.kind === 'paragraph' ? (nextBlock as ParagraphBlock).attrs?.borders : undefined,
  };
}
