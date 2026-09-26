// Text, keyboard, and tooltip helpers for the PowerPoint view. Extracted from
// NativePowerPointView.ts; these are pure aside from reading Obsidian's Platform
// flags and the DOM element passed in.

import { Platform } from 'obsidian';

/** Keep in sync with `EMPTY_PARAGRAPH_RENDER_ANCHOR` in drawingmlText.ts. */
const EMPTY_PARAGRAPH_RENDER_ANCHOR = '\u200B';

export function isPrimaryFindShortcut(evt: KeyboardEvent): boolean {
  const key = evt.key.toLowerCase();
  const isMacFind = evt.metaKey && !evt.ctrlKey;
  const isNonMacFind = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
  const hasPrimaryModifier = isMacFind || isNonMacFind;
  return key === 'f' && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

/**
 * Prefer the physical key for delete direction. Electron/macOS can report the
 * Mac "delete" key (backspace) as `key === 'Delete'` while `code` stays
 * `'Backspace'`. Using `key` alone then forward-deletes the char after the
 * caret (e.g. removes `n` from `i|n` when the user meant to remove `i`).
 */
export function isBackwardDeleteKey(event: Pick<KeyboardEvent, 'key' | 'code'>): boolean {
  if (event.code === 'Backspace') return true;
  if (event.code === 'Delete') return false;
  return event.key === 'Backspace';
}

export function normalizeSearchText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

/**
 * Resolve the pre-edit text used by applyTextValue's unchanged check.
 * Session baseline (captured before live SVG preview mutation) always wins over
 * target.text and over live SVG — otherwise whole-shape preview recovery makes
 * the commit think nothing changed and skip the OOXML write.
 */
export function previousTextForInlineApply(args: {
  sessionBaseline?: string | null;
  targetText?: string | null;
  liveSvgText?: string | null;
}): { previousText: string; source: 'session-baseline' | 'target' | 'live-svg' | 'empty' } {
  if (args.sessionBaseline !== undefined && args.sessionBaseline !== null) {
    return { previousText: args.sessionBaseline, source: 'session-baseline' };
  }
  if (args.targetText !== undefined && args.targetText !== null) {
    return { previousText: args.targetText, source: 'target' };
  }
  if (args.liveSvgText !== undefined && args.liveSvgText !== null) {
    return { previousText: args.liveSvgText, source: 'live-svg' };
  }
  return { previousText: '', source: 'empty' };
}

/** Resolve the word a text caret sits in, using browser-native Unicode segmentation. */
export function getInlineWordRange(text: string, caretOffset: number): { start: number; end: number } {
  const offset = Math.max(0, Math.min(caretOffset, text.length));
  if (!text) return { start: 0, end: 0 };

  const wordSegments: { start: number; end: number }[] = [];
  for (const segment of new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)) {
    if (!segment.isWordLike) continue;
    wordSegments.push({ start: segment.index, end: segment.index + segment.segment.length });
  }

  const containing = wordSegments.find((range) => offset >= range.start && offset < range.end);
  if (containing) return containing;

  for (let index = wordSegments.length - 1; index >= 0; index--) {
    const range = wordSegments[index];
    if (range?.end === offset) return range;
  }
  return { start: offset, end: offset };
}

/** Join visual line fragments for one OOXML paragraph (soft wraps are not newlines). */
export function joinParagraphVisualLines(lineTexts: string[]): string {
  return lineTexts.join('');
}

export type PreviewFrameBox = {
  left: number;
  top: number;
  width: number;
  height: number;
};

/**
 * Convert a candidate frame + first-line left edge into the usable wrap width.
 * Returns null when the frame is too narrow to host horizontal preview text.
 *
 * `firstLineLeft` is only a valid body inset for start-aligned glyphs. Middle /
 * end `text-anchor` (common on empty .potx placeholders) places the glyph box
 * near the center or right edge; treating that as inset collapses wrap width to
 * ~the current glyph width and stacks each typed character on its own line.
 * Pass `firstLineWidth` so that collapse can fall back to the full frame.
 */
export function previewWrapMaxWidth(
  frame: Pick<PreviewFrameBox, 'left' | 'width'>,
  firstLineLeft: number,
  firstLineWidth?: number,
): number | null {
  if (!Number.isFinite(frame.width) || frame.width < 4) return null;

  const inset = Math.max(0, firstLineLeft - frame.left);
  const insetDerived = frame.width - inset * 2;
  if (!Number.isFinite(insetDerived) || insetDerived < 4) {
    // Empty / ZWSP / center-end glyph geometry drove inset past the frame.
    return frame.width;
  }

  const lineWidth = typeof firstLineWidth === 'number' && Number.isFinite(firstLineWidth)
    ? Math.max(0, firstLineWidth)
    : null;
  if (lineWidth !== null) {
    // Live repro (center-aligned empty .potx box): frameBoxWidth≈148, inset≈71,
    // firstLineWidth≈4 → insetDerived≈6. Strict `<= width+1` missed by 1–2px and
    // still wrapped one glyph per line. Treat large inset + short glyphs, or a
    // wrap width that only barely exceeds the glyph box, as anchor collapse.
    const looksCenteredOrEnd =
      inset > lineWidth + 2
      && inset > frame.width * 0.2;
    const collapsedToGlyph =
      insetDerived <= Math.max(lineWidth + 8, lineWidth * 1.5 + 2);
    if (looksCenteredOrEnd || collapsedToGlyph) {
      return frame.width;
    }
  }

  return insetDerived;
}

/**
 * Pick the first usable preview frame in priority order.
 * Callers must pass OOXML/transform before decorative `:scope > rect` candidates —
 * a thin accent rect (~10px) otherwise wraps each glyph onto its own line.
 */
export function pickInlinePreviewFrameBox(
  candidates: ReadonlyArray<{ source: string; box: PreviewFrameBox | null | undefined }>,
): { source: string; box: PreviewFrameBox } | null {
  for (const candidate of candidates) {
    const box = candidate.box;
    if (!box || !Number.isFinite(box.width) || box.width < 4) continue;
    return { source: candidate.source, box };
  }
  return null;
}

/**
 * Split text into visual lines without changing its characters. The caller
 * provides a screen-pixel measurement function so this stays pure and can be
 * shared by the SVG inline-preview path and its tests.
 */
export function wrapTextForPreview(
  text: string,
  maxWidth: number,
  measure: (value: string, startOffset?: number) => number
): string[] {
  if (!text || !Number.isFinite(maxWidth) || maxWidth <= 0) return [text];

  const lines: string[] = [];
  let lineStart = 0;
  let index = 0;
  let lastBreak = -1;

  while (index < text.length) {
    const character = text.charAt(index);
    if (character === '\n') return [text];

    const candidate = text.slice(lineStart, index + 1);
    if (measure(candidate, lineStart) <= maxWidth || index === lineStart) {
      if (/\s/.test(character)) lastBreak = index + 1;
      index += 1;
      continue;
    }

    // Prefer an existing word boundary. Keeping its whitespace in the prior
    // line preserves the editor string exactly while the SVG renders it as a
    // normal soft wrap.
    if (lastBreak > lineStart) {
      lines.push(text.slice(lineStart, lastBreak));
      lineStart = lastBreak;
      index = lineStart;
      lastBreak = -1;
      continue;
    }

    // A space that just crosses the limit belongs to the previous line; do
    // not make the next line begin with a visible leading gap.
    if (/\s/.test(character)) {
      lines.push(candidate);
      lineStart = index + 1;
      index = lineStart;
      lastBreak = -1;
      continue;
    }

    // Long unbroken words still need to stay inside the text box.
    lines.push(text.slice(lineStart, index));
    lineStart = index;
    lastBreak = -1;
  }

  lines.push(text.slice(lineStart));
  return lines;
}

export interface PreviewTextMeasurementSegment {
  start: number;
  end: number;
  measure: (value: string, start: number, end: number) => number;
}

/**
 * Measure a preview substring with the font metrics of each run it crosses.
 * `value` begins at `startOffset` in the full paragraph, so automatic wraps
 * retain the font size at the caret instead of borrowing the first run's size.
 */
export function measureSegmentedPreviewText(
  value: string,
  startOffset: number,
  segments: readonly PreviewTextMeasurementSegment[],
): number {
  const valueStart = Math.max(0, startOffset);
  const valueEnd = valueStart + value.length;
  let width = 0;
  for (const segment of segments) {
    const start = Math.max(valueStart, segment.start);
    const end = Math.min(valueEnd, segment.end);
    if (end <= start) continue;
    width += segment.measure(value, start - valueStart, end - valueStart);
  }
  return width;
}

/**
 * Keep an inline edit's run styling while it is redistributed across visual
 * line tspans.
 *
 * PowerPoint's renderer emits a separate run tspan for every soft-wrapped
 * fragment. This maps the changed range from the previous flat text to the new
 * flat text, allowing the view to retain run styling while it live-reflows the
 * preview. A final engine render on commit remains authoritative.
 */
export function redistributeTextAcrossVisualRuns(previousRunTexts: string[], nextText: string): string[] {
  if (previousRunTexts.length === 0) return [];

  const previousText = previousRunTexts.join('');
  if (previousText === nextText) return [...previousRunTexts];

  let commonPrefixLength = 0;
  const sharedLength = Math.min(previousText.length, nextText.length);
  while (
    commonPrefixLength < sharedLength
    && previousText.charAt(commonPrefixLength) === nextText.charAt(commonPrefixLength)
  ) {
    commonPrefixLength++;
  }

  let commonSuffixLength = 0;
  while (
    commonSuffixLength < sharedLength - commonPrefixLength
    && previousText.charAt(previousText.length - 1 - commonSuffixLength)
      === nextText.charAt(nextText.length - 1 - commonSuffixLength)
  ) {
    commonSuffixLength++;
  }

  const previousChangeStart = commonPrefixLength;
  const previousChangeEnd = previousText.length - commonSuffixLength;
  const nextChangeEnd = nextText.length - commonSuffixLength;
  const delta = nextText.length - previousText.length;
  const boundaries = [0];
  for (const runText of previousRunTexts) {
    boundaries.push((boundaries[boundaries.length - 1] ?? 0) + runText.length);
  }

  // On an insertion that lands on a run boundary, assign the inserted text to
  // the following run (or the final run at paragraph end). This keeps every
  // unchanged fragment on its existing SVG line.
  const ownerIndex = previousRunTexts.findIndex(
    (_, index) => (boundaries[index + 1] ?? 0) > previousChangeStart
  );
  const replacementOwnerIndex = ownerIndex === -1 ? previousRunTexts.length - 1 : ownerIndex;
  const isInsertion = previousChangeStart === previousChangeEnd;

  const mapBoundary = (boundary: number, boundaryIndex: number): number => {
    if (isInsertion) {
      if (boundary < previousChangeStart) return boundary;
      if (boundary > previousChangeStart) return boundary + delta;
      return boundaryIndex > replacementOwnerIndex ? nextChangeEnd : boundary;
    }
    if (boundary <= previousChangeStart) return boundary;
    if (boundary >= previousChangeEnd) return boundary + delta;
    return boundaryIndex <= replacementOwnerIndex ? previousChangeStart : nextChangeEnd;
  };

  const nextBoundaries = boundaries.map(mapBoundary);
  return previousRunTexts.map((_, index) =>
    nextText.slice(nextBoundaries[index] ?? 0, nextBoundaries[index + 1] ?? nextText.length)
  );
}

/**
 * Map a flat (paragraph-relative, run-only) character offset onto the wrapped
 * visual line that contains it.
 *
 * `runLineCharCounts` holds the character count of each visual line that
 * actually contains OOXML runs — bullet/number marker lines are intentionally
 * excluded by the caller, because the inline editor's text and the OOXML run
 * offsets never count the bullet glyph. Including the bullet would shift every
 * offset and underline/highlight the wrong characters.
 */
export function mapFlatOffsetToRunLine(
  runLineCharCounts: number[],
  flatOffset: number
): { lineIndex: number; localOffset: number } {
  if (runLineCharCounts.length === 0) {
    return { lineIndex: 0, localOffset: Math.max(0, flatOffset) };
  }

  const clamped = Math.max(0, flatOffset);
  let offset = 0;
  for (let lineIndex = 0; lineIndex < runLineCharCounts.length; lineIndex++) {
    const count = Math.max(0, runLineCharCounts[lineIndex] ?? 0);
    if (clamped <= offset + count) {
      return { lineIndex, localOffset: clamped - offset };
    }
    offset += count;
  }

  const lastIndex = runLineCharCounts.length - 1;
  return { lineIndex: lastIndex, localOffset: Math.max(0, runLineCharCounts[lastIndex] ?? 0) };
}

/**
 * Split a flat (run-only) character range into the per-visual-line segments that
 * should be highlighted. Returns only the lines that overlap the range.
 */
export interface RunLeafSpan {
  isRun: boolean;
  count: number;
}

/** Map a geometry (all-leaf) index to a run-only offset. Pure helper for tests. */
export function geometryIndexToRunOffset(leafSpans: RunLeafSpan[], geometryIndex: number): number {
  const clamped = Math.max(0, geometryIndex);
  let runOffset = 0;
  let leafOffset = 0;
  for (const span of leafSpans) {
    const count = Math.max(0, span.count);
    if (count <= 0) continue;
    if (clamped <= leafOffset + count) {
      if (span.isRun) {
        runOffset += Math.max(0, clamped - leafOffset);
      }
      return runOffset;
    }
    if (span.isRun) {
      runOffset += count;
    }
    leafOffset += count;
  }
  return runOffset;
}

/** Map a run-only offset to a geometry (all-leaf) index. Pure helper for tests. */
export function runOffsetToGeometryIndex(leafSpans: RunLeafSpan[], runOffset: number): number {
  const clamped = Math.max(0, runOffset);
  let runSeen = 0;
  let leafOffset = 0;
  for (const span of leafSpans) {
    const count = Math.max(0, span.count);
    if (count <= 0) continue;
    if (span.isRun) {
      if (clamped <= runSeen + count) {
        return leafOffset + Math.max(0, clamped - runSeen);
      }
      runSeen += count;
    }
    leafOffset += count;
  }
  return leafOffset;
}

export function mapFlatRangeToRunLineSegments(
  runLineCharCounts: number[],
  flatStart: number,
  flatEnd: number
): { lineIndex: number; localStart: number; localEnd: number }[] {
  const start = Math.max(0, Math.min(flatStart, flatEnd));
  const end = Math.max(flatStart, flatEnd);
  const segments: { lineIndex: number; localStart: number; localEnd: number }[] = [];

  let offset = 0;
  for (let lineIndex = 0; lineIndex < runLineCharCounts.length; lineIndex++) {
    const count = Math.max(0, runLineCharCounts[lineIndex] ?? 0);
    const localStart = Math.max(0, start - offset);
    const localEnd = Math.min(count, end - offset);
    if (localEnd > localStart) {
      segments.push({ lineIndex, localStart, localEnd });
    }
    offset += count;
  }

  return segments;
}

/**
 * The inline editor's text is rebuilt from the rendered SVG runs, which omit the
 * whitespace PowerPoint swallows at soft-wrap boundaries. The OOXML run text
 * keeps those characters, so an editor (SVG-space) offset is always <= the
 * matching OOXML offset. The editor text is therefore a subsequence of the
 * OOXML text; this walks both and returns the OOXML offset aligned with
 * `editorOffset`.
 *
 * `consumeTrailingGap` is used for a range END: any OOXML characters that were
 * dropped from the SVG and sit immediately after the mapped position are also
 * consumed, so a selection that visually reaches a wrap boundary spans the
 * swallowed whitespace instead of leaving it formatted/highlighted.
 */
export function mapEditorOffsetToOoxmlOffset(
  editorText: string,
  ooxmlText: string,
  editorOffset: number,
  consumeTrailingGap = false
): number {
  const target = Math.max(0, Math.min(editorOffset, editorText.length));
  if (editorText === ooxmlText) return target;

  let e = 0;
  let o = 0;
  while (e < target && o < ooxmlText.length) {
    if (editorText[e] === ooxmlText[o]) {
      e++;
      o++;
    } else if (editorText[e] === '\n') {
      // A soft break (`<a:br/>`) is an editor-only glyph: the OOXML run text has
      // no counterpart, so consume the newline from the editor side without
      // advancing OOXML. Otherwise the walk would treat a real OOXML character
      // as "dropped" and overshoot, landing a paragraph split on the wrong
      // glyph. When the comparison string itself carries the break (a pending
      // split mapped against editor-space text) the equality branch above
      // already matched it, so this only fires against run-only OOXML text.
      e++;
    } else {
      // OOXML char dropped from the SVG (whitespace at a wrap boundary).
      o++;
    }
  }

  if (consumeTrailingGap) {
    // Only whitespace (and empty-paragraph ZWSP anchors) is swallowed at a
    // soft-wrap / empty-render boundary, so a range END may absorb trailing
    // OOXML gap chars the editor dropped -- but it must never run past a real
    // character, which would over-clear onto text outside the selection.
    while (
      o < ooxmlText.length
      && (/\s/.test(ooxmlText.charAt(o)) || ooxmlText.charAt(o) === EMPTY_PARAGRAPH_RENDER_ANCHOR)
      && (e >= editorText.length || editorText[e] !== ooxmlText[o])
    ) {
      o++;
    }
  }

  return o;
}

export interface RunTspanOffset {
  /** Offset of this tspan's first char in the run-only editor string. */
  editorStart: number;
  /** Offset just past this tspan's last char in the run-only editor string. */
  editorEnd: number;
  /** OOXML run offset of this tspan's first char (leading dropped whitespace excluded). */
  charStart: number;
  /** OOXML run offset just past this tspan's last char (trailing dropped wrap whitespace included). */
  charEnd: number;
}

export interface RunTspanAlignment {
  spans: RunTspanOffset[];
  editorLength: number;
  ooxmlLength: number;
  /**
   * Whether the alignment fully reconciles: the editor text is a subsequence no
   * longer than the OOXML text, and the run tspans cover the OOXML text end to
   * end with monotonic, non-overlapping (except shared dropped-whitespace)
   * boundaries. A `false` here is the single signal that the SVG↔OOXML bridge
   * has drifted and should be logged loudly by the caller.
   */
  reconciled: boolean;
}

/**
 * Align the rendered run tspans (in document order) to the paragraph's OOXML run
 * text, computing each tspan's OOXML char range exactly once. The renderer emits
 * one tspan per (run × visual line) and drops the whitespace PowerPoint swallows
 * at soft-wrap boundaries, so the concatenated tspan text is a subsequence of the
 * OOXML run text. `charStart` uses START semantics (excludes leading dropped
 * whitespace) and `charEnd` uses END semantics (absorbs trailing dropped wrap
 * whitespace), so the ranges tile the whole paragraph with the swallowed spaces
 * attributed to the preceding line — which is exactly what selection/clear want.
 *
 * This is the one place the SVG↔OOXML offset inference happens; every consumer
 * reads the stamped result instead of re-deriving it.
 */
export function alignRunTspansToOoxml(runTexts: string[], ooxmlText: string): RunTspanAlignment {
  const editor = runTexts.join('');
  const spans: RunTspanOffset[] = [];
  let editorPos = 0;
  for (const text of runTexts) {
    const editorStart = editorPos;
    const editorEnd = editorPos + text.length;
    // Both boundaries consume dropped wrap whitespace forward, so a swallowed
    // space is attributed to the END of the line before it and the next line
    // starts on its first real glyph. This tiles the paragraph with no overlap.
    const charStart = mapEditorOffsetToOoxmlOffset(editor, ooxmlText, editorStart, true);
    const charEnd = mapEditorOffsetToOoxmlOffset(editor, ooxmlText, editorEnd, true);
    spans.push({ editorStart, editorEnd, charStart, charEnd });
    editorPos = editorEnd;
  }

  let monotonic = true;
  let previousStart = 0;
  let previousEnd = 0;
  for (const span of spans) {
    if (span.charEnd < span.charStart || span.charStart < previousStart || span.charEnd < previousEnd) {
      monotonic = false;
      break;
    }
    previousStart = span.charStart;
    previousEnd = span.charEnd;
  }

  const lastEnd = spans.length === 0 ? 0 : spans[spans.length - 1]!.charEnd;
  const coversEnd = spans.length === 0 ? editor.length === 0 : lastEnd === ooxmlText.length;
  const reconciled = monotonic && coversEnd && editor.length <= ooxmlText.length;

  return { spans, editorLength: editor.length, ooxmlLength: ooxmlText.length, reconciled };
}

/**
 * Map an inline-editor (SVG run-offset) range to OOXML offsets using pre-aligned
 * run-tspan tiles (the stamped `data-ooxml-char-*` ranges). Within a tile the two
 * texts agree, so the mapping is linear; START lands on the first real glyph and
 * END absorbs trailing dropped whitespace — matching the engine clear/format
 * semantics of {@link mapEditorOffsetToOoxmlOffset} with false/true, but as a
 * pure O(n) lookup over the stamps rather than a per-call subsequence walk.
 */
export function mapEditorRangeToOoxml(
  tiles: RunTspanOffset[],
  editorStart: number,
  editorEnd: number
): { start: number; end: number } | null {
  if (tiles.length === 0) return null;
  const last = tiles[tiles.length - 1]!;
  const editorTotal = last.editorEnd;

    // START lands on the first selected glyph: at a wrap boundary it skips the
    // dropped space (owned by the preceding line's END), so each swallowed space
    // is attributed to exactly one line and selections never overlap on it.
  const mapStart = (offset: number): number => {
    const clamped = Math.max(0, Math.min(editorTotal, offset));
    for (const tile of tiles) {
      if (clamped < tile.editorEnd || (clamped === tile.editorEnd && tile === last)) {
        return tile.charStart + Math.max(0, clamped - tile.editorStart);
      }
    }
    return last.charEnd;
  };
  const mapEnd = (offset: number): number => {
    const clamped = Math.max(0, Math.min(editorTotal, offset));
    for (const tile of tiles) {
      if (clamped > tile.editorStart && clamped <= tile.editorEnd) {
        return clamped === tile.editorEnd ? tile.charEnd : tile.charStart + (clamped - tile.editorStart);
      }
    }
    return tiles[0]!.charStart;
  };

  return { start: mapStart(editorStart), end: mapEnd(editorEnd) };
}

/** First face from a CSS `font-family` value, e.g. `"Calibri", sans-serif` → Calibri. */
export function parsePrimaryFontFamily(fontFamily: string): string | null {
  const trimmed = fontFamily.trim();
  if (!trimmed) return null;

  if (trimmed.startsWith('"')) {
    const end = trimmed.indexOf('"', 1);
    if (end > 1) return trimmed.slice(1, end);
  } else if (trimmed.startsWith("'")) {
    const end = trimmed.indexOf("'", 1);
    if (end > 1) return trimmed.slice(1, end);
  }

  const comma = trimmed.indexOf(',');
  const primary = (comma >= 0 ? trimmed.slice(0, comma) : trimmed).trim();
  if (!primary || primary === 'inherit' || primary === 'initial' || primary === 'unset') return null;

  const generic = new Set([
    'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif',
    'ui-monospace', 'ui-rounded', 'emoji', 'math', 'fangsong'
  ]);
  return generic.has(primary.toLowerCase()) ? null : primary;
}
