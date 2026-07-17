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

test("clipboard matrix: engine copies and pastes a multi-selection atomically", async () => {
  const engine = await loadEngine("features.pptx");
  const baselineCount = renderedShapeCount(engine);
  assert.ok(baselineCount >= 2, "fixture must provide two independently selectable objects");

  const clipboard = await engine.copyShapes(0, [2, 0, 2]);
  assert.deepEqual(clipboard.shapeIndexes, [0, 2]);

  const pastedIndexes = await engine.pasteShapes(clipboard, 0);
  assert.equal(pastedIndexes.length, 2, "engine must return both pasted renderer indexes");
  assert.equal(renderedShapeCount(engine), baselineCount + 2, "paste must add every copied object");

  await assertExportRoundTrips("multi-object clipboard matrix", engine);
});

test("object matrix: engine deletes a multi-selection in one package reload", async () => {
  const engine = await loadEngine("features.pptx");
  const baselineCount = renderedShapeCount(engine);
  assert.ok(baselineCount >= 3, "fixture must provide three independently selectable objects");

  const originalReload = engine.reloadFromBuffer;
  let reloadCount = 0;
  engine.reloadFromBuffer = async (...args) => {
    reloadCount += 1;
    return originalReload.apply(engine, args);
  };

  await engine.deleteShapes(0, [2, 0, 2]);

  assert.equal(reloadCount, 1, "multi-delete must reload the package once");
  assert.equal(renderedShapeCount(engine), baselineCount - 2, "every selected object must be deleted");
  await assertExportRoundTrips("multi-object delete matrix", engine);
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
