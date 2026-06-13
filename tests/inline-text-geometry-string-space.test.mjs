import assert from "node:assert/strict";
import test from "node:test";
import { loadInlineTextGeometryModule } from "./helpers/load-plugin-modules.mjs";

// A leaf <tspan> stand-in. `getNumberOfChars` reports the rendered GLYPH count,
// which can be smaller than the textContent string length when font shaping
// collapses characters into ligatures — the historical offset-drift source.
function fakeSpan(text, glyphCount = text.length) {
  return {
    textContent: text,
    getNumberOfChars() {
      return glyphCount;
    },
  };
}

test("leafCharCount counts the editor string, never the glyph count", async () => {
  const { InlineTextGeometry } = await loadInlineTextGeometryModule();
  const geometry = new InlineTextGeometry(() => null);

  // "office" shapes "ffi" into a single ligature glyph: 6 chars, 4 glyphs.
  const span = fakeSpan("office", 4);
  assert.equal(geometry.leafCharCount(span), 6, "count is the string length the stamps/editor index");
  assert.equal(geometry.getGlyphCount(span), 4, "glyph count tracks the rendered glyphs");
});

test("getGlyphCount falls back to the string length when probing throws", async () => {
  const { InlineTextGeometry } = await loadInlineTextGeometryModule();
  const geometry = new InlineTextGeometry(() => null);

  const detached = {
    textContent: "hello",
    getNumberOfChars() {
      throw new Error("not rendered");
    },
  };
  assert.equal(geometry.getGlyphCount(detached), 5);
});

test("clampGlyphIndex keeps string offsets inside the rendered glyph range", async () => {
  const { InlineTextGeometry } = await loadInlineTextGeometryModule();
  const geometry = new InlineTextGeometry(() => null);

  const span = fakeSpan("office", 4); // 6 chars, 4 glyphs (indices 0..3)
  assert.equal(geometry.clampGlyphIndex(span, 0), 0);
  assert.equal(geometry.clampGlyphIndex(span, 3), 3);
  // String indices 4 and 5 have no glyph of their own; clamp to the last glyph
  // instead of throwing out of range inside getStartPositionOfChar/getExtentOfChar.
  assert.equal(geometry.clampGlyphIndex(span, 4), 3);
  assert.equal(geometry.clampGlyphIndex(span, 5), 3);
  assert.equal(geometry.clampGlyphIndex(span, -2), 0);

  const empty = fakeSpan("", 0);
  assert.equal(geometry.clampGlyphIndex(empty, 0), 0);
});
