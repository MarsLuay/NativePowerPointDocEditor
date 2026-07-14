// Text, keyboard, and tooltip helpers for the PowerPoint view. Extracted from
// NativePowerPointView.ts; these are pure aside from reading Obsidian's Platform
// flags and the DOM element passed in.

import { Platform } from 'obsidian';

export function isPrimaryFindShortcut(evt: KeyboardEvent): boolean {
  const key = evt.key.toLowerCase();
  const isMacFind = evt.metaKey && !evt.ctrlKey;
  const isNonMacFind = evt.ctrlKey && !evt.metaKey && !Platform.isMacOS;
  const hasPrimaryModifier = isMacFind || isNonMacFind;
  return key === 'f' && hasPrimaryModifier && !evt.altKey && !evt.shiftKey;
}

export function normalizeSearchText(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
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
    } else {
      // OOXML char dropped from the SVG (whitespace at a wrap boundary).
      o++;
    }
  }

  if (consumeTrailingGap) {
    // Only whitespace is ever swallowed at a soft-wrap boundary, so a range END
    // may absorb trailing OOXML whitespace that the editor dropped -- but it must
    // never run past a real (non-whitespace) character, which would over-clear
    // onto text outside the selection if the two strings ever diverge.
    while (
      o < ooxmlText.length
      && /\s/.test(ooxmlText.charAt(o))
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
