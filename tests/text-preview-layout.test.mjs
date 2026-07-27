import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { loadTextUtilsModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = join(dirname(fileURLToPath(import.meta.url)), "..");

test("inline paragraph preview preserves its wrapped SVG line containers after a backspace", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previousLines = [
    "They make setup easier for newer users and reduce the need to ",
    "watch the whole process.",
  ];
  const nextText = "They make setup easier for newer users and reduce the nee to watch the whole process.";

  const previewLines = redistributeTextAcrossVisualRuns(previousLines, nextText);

  assert.equal(previewLines.join(""), nextText);
  assert.equal(previewLines.length, previousLines.length);
  assert.equal(previewLines[0], previousLines[0].replace("need", "nee"));
  assert.equal(previewLines[1], previousLines[1]);
});

test("inline paragraph preview does not duplicate replacement text across wrapped lines", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previewLines = redistributeTextAcrossVisualRuns(["first ", "second ", "third"], "replacement");

  assert.deepEqual(previewLines, ["replacement", "", ""]);
});

test("inline paragraph preview keeps an end insertion inside the final visual line", async () => {
  const { redistributeTextAcrossVisualRuns } = await loadTextUtilsModule();
  const previewLines = redistributeTextAcrossVisualRuns(["first ", "second"], "first second!");

  assert.deepEqual(previewLines, ["first ", "second!"]);
});

test("inline paragraph preview creates new word-wrapped lines without losing text", async () => {
  const { wrapTextForPreview } = await loadTextUtilsModule();
  const lines = wrapTextForPreview("alpha beta gamma", 10, (value) => value.length);

  assert.deepEqual(lines, ["alpha ", "beta gamma"]);
  assert.equal(lines.join(""), "alpha beta gamma");
});

test("inline paragraph preview breaks an overlong unspaced word", async () => {
  const { wrapTextForPreview } = await loadTextUtilsModule();
  const lines = wrapTextForPreview("abcdefgh", 3, (value) => value.length);

  assert.deepEqual(lines, ["abc", "def", "gh"]);
  assert.equal(lines.join(""), "abcdefgh");
});

test("preview wrap width subtracts symmetric text insets from the frame", async () => {
  const { previewWrapMaxWidth } = await loadTextUtilsModule();
  assert.equal(previewWrapMaxWidth({ left: 100, width: 220 }, 112), 196);
  assert.equal(previewWrapMaxWidth({ left: 100, width: 12 }, 100), 12);
  // Collapsed inset (center/end empty glyph) falls back to the full frame.
  assert.equal(previewWrapMaxWidth({ left: 100, width: 10 }, 108), 10);
  assert.equal(previewWrapMaxWidth({ left: 100, width: 3 }, 100), null);
});

test("preview wrap width ignores center-anchored glyph left so typing stays horizontal", async () => {
  const { previewWrapMaxWidth, wrapTextForPreview } = await loadTextUtilsModule();
  // Empty / short centered line: glyph box sits mid-frame with ~current text width.
  // Old inset math made maxWidth ≈ glyph width → one character per visual line.
  const frame = { left: 100, width: 400 };
  const centeredGlyphLeft = 290;
  const centeredGlyphWidth = 20;
  const maxWidth = previewWrapMaxWidth(frame, centeredGlyphLeft, centeredGlyphWidth);
  assert.equal(maxWidth, 400);

  const lines = wrapTextForPreview("abc", maxWidth, (value) => value.length * 10);
  assert.deepEqual(lines, ["abc"]);
  assert.equal(lines.length, 1);
});

test("preview wrap width falls back when inset-derived width is only a few px above the glyph", async () => {
  const { previewWrapMaxWidth, wrapTextForPreview } = await loadTextUtilsModule();
  // Exact live-log shape from shapeIndex 37: frame 148, inset 71, glyph 4 →
  // insetDerived 6, which failed the old `<= width+1` check.
  const maxWidth = previewWrapMaxWidth({ left: 0, width: 148 }, 71, 4);
  assert.equal(maxWidth, 148);
  assert.deepEqual(
    wrapTextForPreview("dfof", maxWidth, (value) => value.length * 10),
    ["dfof"],
  );
});

test("preview wrap width still uses start-aligned insets when the line is shorter than the frame", async () => {
  const { previewWrapMaxWidth } = await loadTextUtilsModule();
  assert.equal(
    previewWrapMaxWidth({ left: 100, width: 220 }, 112, 40),
    196,
  );
});

test("preview frame picker prefers OOXML transform over a thin decorative rect", async () => {
  const { pickInlinePreviewFrameBox } = await loadTextUtilsModule();
  const picked = pickInlinePreviewFrameBox([
    { source: "ooxml-transform", box: { left: 40, top: 10, width: 320, height: 48 } },
    { source: "shape-bounds", box: { left: 40, top: 10, width: 318, height: 50 } },
    { source: "direct-rect", box: { left: 40, top: 10, width: 10, height: 48 } },
  ]);

  assert.equal(picked?.source, "ooxml-transform");
  assert.equal(picked?.box.width, 320);
});

test("preview frame picker skips unusable frames then accepts the next candidate", async () => {
  const { pickInlinePreviewFrameBox } = await loadTextUtilsModule();
  const picked = pickInlinePreviewFrameBox([
    { source: "ooxml-transform", box: null },
    { source: "shape-bounds", box: { left: 0, top: 0, width: 3, height: 10 } },
    { source: "direct-rect", box: { left: 0, top: 0, width: 180, height: 40 } },
  ]);

  assert.equal(picked?.source, "direct-rect");
  assert.equal(picked?.box.width, 180);
});

test("live inline reflow moves following paragraphs by its wrapped-line delta", () => {
  const view = readFileSync(
    join(projectRoot, "src/powerpoint/ui/NativePowerPointView.ts"),
    "utf8",
  );
  const reflow = view.slice(
    view.indexOf("private reflowShapeParagraphPreview"),
    view.indexOf("private createInlinePreviewTextMeasurer"),
  );

  assert.match(reflow, /const lineDelta = nextLines\.length - lineContainers\.length;/);
  assert.match(
    reflow,
    /this\.shiftLocalPreviewParagraphs\(\s*textElement,\s*target\.paragraphIndex \+ 1,\s*0,\s*lineStep \* lineDelta,\s*\);/,
  );
});

test("paragraph split preview rejects wrapped or downstream text before mutating live SVG", () => {
  const view = readFileSync(
    join(projectRoot, "src/powerpoint/ui/NativePowerPointView.ts"),
    "utf8",
  );
  const splitPreview = view.slice(
    view.indexOf("private previewInlineParagraphSplit"),
    view.indexOf("/** Durable geometry breadcrumb for paragraph-split regressions."),
  );

  assert.match(splitPreview, /sourceLinesBeforePreview\.length !== 1/);
  assert.match(splitPreview, /hasDownstreamParagraph/);
  assert.match(splitPreview, /trailingLineCount !== 1/);
  assert.match(splitPreview, /Skipped live PowerPoint paragraph split preview/);
});
