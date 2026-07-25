import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertShapeIndex,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";

const SLIDE_INDEX = 0;

function requireEngineApi(t, engine, api) {
  if (typeof engine[api] === "function") return true;
  t.skip(`${api} is not exposed by PresentationEngine`);
  return false;
}

test("nudge skips when no engine API exists", async (t) => {
  const engine = await loadEngine();
  if (!requireEngineApi(t, engine, "nudgeShapes")) return;

  await engine.nudgeShapes(SLIDE_INDEX, [0], 1, 0);
  await assertExportRoundTrips("nudge", engine);
});

test("align skips when no engine API exists", async (t) => {
  const engine = await loadEngine();
  if (!requireEngineApi(t, engine, "alignShapes")) return;

  await engine.alignShapes(SLIDE_INDEX, [0, 1], "left");
  await assertExportRoundTrips("align", engine);
});

test("distribute skips when no engine API exists", async (t) => {
  const engine = await loadEngine();
  if (!requireEngineApi(t, engine, "distributeShapes")) return;

  await engine.distributeShapes(SLIDE_INDEX, [0, 1], "horizontal");
  await assertExportRoundTrips("distribute", engine);
});

for (const mode of ["front", "forward", "backward", "back"]) {
  test(`reorder ${mode} commits through PresentationEngine`, async (t) => {
    const engine = await loadEngine();
    if (!requireEngineApi(t, engine, "reorderShapes")) return;

    const result = await engine.reorderShapes(SLIDE_INDEX, [0], mode);
    assert.deepEqual(result.length, 1, `${mode}: expected one returned shape index`);
    assertShapeIndex(`${mode} reorder`, result[0]);
    await assertExportRoundTrips(`reorder ${mode}`, engine);
  });
}

test("group and ungroup commit through PresentationEngine", async (t) => {
  const engine = await loadEngine();
  if (!requireEngineApi(t, engine, "groupShapes")) return;
  if (!requireEngineApi(t, engine, "ungroupShapes")) return;

  const groupIndex = await engine.groupShapes(SLIDE_INDEX, [0, 1]);
  assertShapeIndex("group", groupIndex);

  const childIndexes = await engine.ungroupShapes(SLIDE_INDEX, groupIndex);
  assert.deepEqual(childIndexes.length, 2, "ungroup should restore both selected shapes");
  childIndexes.forEach((shapeIndex) => assertShapeIndex("ungroup", shapeIndex));
  await assertExportRoundTrips("group and ungroup", engine);
});

test("resize and rotate commit through updateShapeTransform", async (t) => {
  const engine = await loadEngine();
  if (!requireEngineApi(t, engine, "updateShapeTransform")) return;

  const fragment = await engine.updateShapeTransform(SLIDE_INDEX, 0, {
    x: 123456,
    y: 234567,
    cx: 3456789,
    cy: 456789,
    rot: 2700000,
  });

  const slideXml = engine.getSlideXml(SLIDE_INDEX);
  assert.match(slideXml, /x="123456"/);
  assert.match(slideXml, /y="234567"/);
  assert.match(slideXml, /cx="3456789"/);
  assert.match(slideXml, /cy="456789"/);
  assert.match(slideXml, /rot="2700000"/);
  assert.match(fragment ?? "", /<g\b[^>]*data-ooxml-shape-idx="0"/);
  await assertExportRoundTrips("resize and rotate", engine);
});
