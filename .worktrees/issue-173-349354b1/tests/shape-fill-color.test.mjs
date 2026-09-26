import assert from "node:assert/strict";
import { test } from "node:test";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

async function loadEngine() {
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(toArrayBuffer(await readDeck("features.pptx")));
}

test("shape fill color survives export and reload", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await loadEngine();
  const shapeIndex = Array.from({ length: 20 }, (_, index) => index)
    .find((index) => engine.canSetShapeFillColor(0, index));

  assert.notEqual(shapeIndex, undefined, "fixture must contain a fill-capable shape");
  await engine.setShapeFillColor(0, shapeIndex, "#12ab34");
  assert.equal(engine.getShapeVisualStyle(0, shapeIndex)?.fill, "12AB34");

  const reloaded = await PresentationEngine.load(await engine.export());
  assert.equal(reloaded.getShapeVisualStyle(0, shapeIndex)?.fill, "12AB34");
});

test("shape fill rejects invalid colors and unsupported objects", async () => {
  const engine = await loadEngine();
  const supported = Array.from({ length: 20 }, (_, index) => index)
    .find((index) => engine.canSetShapeFillColor(0, index));
  const unsupported = Array.from({ length: 20 }, (_, index) => index)
    .find((index) => engine.getShapeVisualStyle(0, index) !== null && !engine.canSetShapeFillColor(0, index));

  assert.notEqual(supported, undefined);
  await assert.rejects(() => engine.setShapeFillColor(0, supported, "not-a-color"), /6-digit RRGGBB/);
  if (unsupported !== undefined) {
    await assert.rejects(() => engine.setShapeFillColor(0, unsupported, "ABCDEF"), /does not support a fill color/);
  }
});

test("text box detection excludes ordinary labeled auto shapes", async () => {
  const engine = await loadEngine();
  const textBoxIndex = await engine.addTextBox(0);
  const autoShapeIndex = await engine.addShapeGeometry(0, "rect");

  assert.equal(engine.isTextBoxShape(0, textBoxIndex), true);
  assert.equal(engine.isTextBoxShape(0, autoShapeIndex), false);
});
