import assert from "node:assert/strict";
import { test } from "node:test";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";
import { createRenderer, readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

// Audit coverage for the "PPTX · Slides: navigate & manage" feature group.
// Exercises the PresentationEngine slide-mutation methods that back the UI
// thumbnail filmstrip (add / duplicate / delete / reorder / new-slide layouts)
// and confirms each mutation keeps slideCount honest, renders, and round-trips
// through an export that reloads with the expected slide count.

async function loadEngine(name = "features.pptx") {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck(name);
  const engine = await PresentationEngine.load(toArrayBuffer(input));
  return engine;
}

test("adding a slide increments slideCount and the new slide renders", async () => {
  const engine = await loadEngine();
  assert.equal(engine.slideCount, 1);

  const result = await engine.addSlide(0);
  assert.equal(result.slideCount, 2);
  assert.equal(engine.slideCount, 2);
  // The engine reports where the new slide landed; it must be a real index.
  assert.ok(result.slideIndex >= 0 && result.slideIndex < engine.slideCount);
  assert.match(engine.renderSlide(result.slideIndex).svg, /^<svg\b/);
});

test("duplicating a slide increments slideCount and renders the copy", async () => {
  const engine = await loadEngine();
  assert.equal(engine.slideCount, 1);

  const result = await engine.duplicateSlide(0);
  assert.equal(result.slideCount, 2);
  assert.equal(engine.slideCount, 2);
  assert.ok(result.slideIndex >= 0 && result.slideIndex < engine.slideCount);
  assert.match(engine.renderSlide(result.slideIndex).svg, /^<svg\b/);
});

test("deleting a slide decrements slideCount and is guarded at one slide", async () => {
  const engine = await loadEngine();
  await engine.addSlide(0);
  assert.equal(engine.slideCount, 2);

  const result = await engine.deleteSlide(1);
  assert.equal(result.slideCount, 1);
  assert.equal(engine.slideCount, 1);
  assert.match(engine.renderSlide(0).svg, /^<svg\b/);

  // Delete-last-slide guard: the engine must refuse to drop the final slide.
  await assert.rejects(() => engine.deleteSlide(0), /at least one slide/i);
  assert.equal(engine.slideCount, 1);
});

test("new-slide layouts (blank/title/titleBody) each add a slide", async () => {
  const engine = await loadEngine();
  let expected = engine.slideCount;

  for (const layout of ["blank", "title", "titleBody"]) {
    const result = await engine.addSlideWithLayout(engine.slideCount - 1, layout);
    expected += 1;
    assert.equal(result.slideCount, expected, `layout ${layout} should add one slide`);
    assert.equal(engine.slideCount, expected);
    assert.match(engine.renderSlide(result.slideIndex).svg, /^<svg\b/);
  }
});

test("reorder changes slide order and export round-trips to a loadable deck", async () => {
  // The large deck ships with content-distinguishable slides ("Large deck
  // slide N"), so a reorder is observable in the exported XML. Newly added or
  // duplicated slides start blank, hence the dedicated fixture here.
  const engine = await loadEngine("large-deck.pptx");
  assert.equal(engine.slideCount, 160);

  // Capture the first three slides' OOXML before reordering. The deck's slides
  // are content-distinguishable, so slide identity is observable in the XML.
  const before = await createRenderer(new Uint8Array(await engine.export()));
  assert.equal(before.getSlideCount(), 160);
  const before0 = before.getSlideOoxml(0);
  const before1 = before.getSlideOoxml(1);
  const before2 = before.getSlideOoxml(2);
  assert.notEqual(before0, before1, "fixture slides must be distinguishable for this assertion");

  // Move slide 0 down one position; the first two slides must swap.
  const moved = await engine.moveSlide(0, 1);
  assert.equal(moved.slideCount, 160);
  assert.equal(moved.slideIndex, 1);

  const after = await createRenderer(new Uint8Array(await engine.export()));
  assert.equal(after.getSlideCount(), 160, "exported deck keeps the expected slide count");
  assert.equal(after.getSlideOoxml(0), before1, "reorder moved the second slide to the front");
  assert.equal(after.getSlideOoxml(1), before0, "reorder pushed the first slide back");
  assert.equal(after.getSlideOoxml(2), before2, "slides past the swap stay put");
  assert.match(after.renderSlideSvg(0), /^<svg\b/);
});

test("a full add/duplicate/delete sequence exports to a still-loadable deck", async () => {
  const engine = await loadEngine();
  assert.equal((await engine.addSlide(0)).slideCount, 2);
  assert.equal((await engine.duplicateSlide(0)).slideCount, 3);
  assert.equal((await engine.deleteSlide(1)).slideCount, 2);

  const reloaded = await createRenderer(new Uint8Array(await engine.export()));
  assert.equal(reloaded.getSlideCount(), 2);
  assert.match(reloaded.renderSlideSvg(0), /^<svg\b/);
  assert.match(reloaded.renderSlideSvg(1), /^<svg\b/);
});
