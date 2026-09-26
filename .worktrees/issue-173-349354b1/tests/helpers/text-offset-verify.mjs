// Shared verification for text-offset stamping and range round-trips.

import assert from "node:assert/strict";

/**
 * Brute-force editor→OOXML range mapping (the semantics mapEditorRangeToOoxml
 * must match). Uses the full editor/OOXML strings, not tiles.
 */
export function referenceEditorRangeToOoxml(editorText, ooxmlText, editorStart, editorEnd, mapOffset) {
  return {
    start: mapOffset(editorText, ooxmlText, editorStart, false),
    end: mapOffset(editorText, ooxmlText, editorEnd, true),
  };
}

/** Assert stamped DOM tiles match the alignment tiles from alignRunTspansToOoxml. */
export function assertStampsMatchAlignment(runSpans, alignment) {
  assert.equal(runSpans.length, alignment.spans.length, "stamp count must match alignment span count");
  runSpans.forEach((span, index) => {
    const tile = alignment.spans[index];
    assert.equal(span.getAttribute("data-ooxml-char-start"), String(tile.charStart), `span ${index} charStart`);
    assert.equal(span.getAttribute("data-ooxml-char-end"), String(tile.charEnd), `span ${index} charEnd`);
  });
}

/**
 * Property check: mapEditorRangeToOoxml(tiles) agrees with the reference
 * subsequence walk for random editor ranges.
 */
export function fuzzEditorRangeRoundTrip({
  editorText,
  ooxmlText,
  tiles,
  mapEditorRangeToOoxml,
  mapEditorOffsetToOoxmlOffset,
  iterations = 48,
  seed = 1,
}) {
  let rng = seed >>> 0;
  const next = () => {
    rng = (rng * 1664525 + 1013904223) >>> 0;
    return rng / 0x1_0000_0000;
  };

  const editorLen = editorText.length;
  if (editorLen === 0) {
    const mapped = mapEditorRangeToOoxml(tiles, 0, 0);
    const reference = referenceEditorRangeToOoxml(
      editorText,
      ooxmlText,
      0,
      0,
      mapEditorOffsetToOoxmlOffset,
    );
    assert.deepEqual(mapped, reference, "empty editor range");
    return { iterations: 1, mismatches: [] };
  }

  const mismatches = [];
  const tries = Math.max(iterations, 1);
  for (let i = 0; i < tries; i++) {
    const a = Math.floor(next() * (editorLen + 1));
    const b = Math.floor(next() * (editorLen + 1));
    const start = Math.min(a, b);
    const end = Math.max(a, b);

    const mapped = mapEditorRangeToOoxml(tiles, start, end);
    const reference = referenceEditorRangeToOoxml(
      editorText,
      ooxmlText,
      start,
      end,
      mapEditorOffsetToOoxmlOffset,
    );
    if (
      !mapped
      || mapped.start !== reference.start
      || mapped.end !== reference.end
    ) {
      mismatches.push({ start, end, mapped, reference, editorText: editorText.slice(0, 80), ooxmlText: ooxmlText.slice(0, 80) });
    }
  }

  return { iterations: tries, mismatches };
}
