import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertShapeIndex,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";

function renderedShapeCount(engine, slideIndex = 0) {
  return (engine.renderSlide(slideIndex).svg.match(/data-ooxml-shape-idx=/g) ?? []).length;
}

test("clipboard matrix: engine copy, paste, and duplicate preserve a usable deck", async () => {
  const engine = await loadEngine("simple-edit.pptx");
  const baselineCount = renderedShapeCount(engine);

  const clipboard = await engine.copyShape(0, 0);
  assert.ok(clipboard, "copy must produce a clipboard payload");

  const pastedIndex = await engine.pasteShape(clipboard, 0);
  assertShapeIndex("paste", pastedIndex);
  assert.equal(renderedShapeCount(engine), baselineCount + 1, "paste must add one shape");

  const duplicateIndex = await engine.duplicateShape(0, 0);
  assertShapeIndex("duplicate", duplicateIndex);
  assert.equal(renderedShapeCount(engine), baselineCount + 2, "duplicate must add one shape");

  await assertExportRoundTrips("clipboard matrix", engine);
});

test("find/replace matrix: scoped and deck-wide engine replacement", async () => {
  const engine = await loadEngine("simple-edit.pptx");
  await engine.updateParagraphText(0, 0, 0, "First first FIRST");

  const scopedCount = await engine.replaceText("First", "one", {
    matchCase: true,
    slideIndex: 0,
    shapeIndex: 0,
  });
  assert.equal(scopedCount, 1, "scoped replace must respect case and shape scope");

  const deckWideCount = await engine.replaceText("first", "all");
  assert.equal(deckWideCount, 2, "deck-wide replace must find remaining case-insensitive matches");

  const renderedText = engine.renderSlide(0).svg.replace(/<[^>]+>/g, "");
  assert.ok(renderedText.includes("one all all"), "rendered text must contain scoped and deck-wide replacements");
  assert.ok(!renderedText.includes("First"), "original text must be removed");

  await assertExportRoundTrips("find/replace matrix", engine);
});
