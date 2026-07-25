import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

test("PPTX text edit and formatting matrix round-trips", async () => {
  const engine = await loadEngine("features.pptx");
  const slideIndex = 0;
  const shapeIndex = 0;
  const paragraphIndex = 0;
  const text = "Matrix text";

  await engine.updateParagraphText(slideIndex, shapeIndex, paragraphIndex, text);
  assert.match(engine.renderSlide(slideIndex).svg, />Matrix\s*<\/tspan>/);
  assert.match(engine.renderSlide(slideIndex).svg, />text<\/tspan>/);

  const styleCases = [
    ["bold", true],
    ["fontFamily", "Georgia"],
    ["fontSizePt", 31],
    ["color", "112233"],
  ];
  for (const [property, value] of styleCases) {
    await engine.setRunStyle(
      slideIndex,
      shapeIndex,
      { paragraphIndex, runIndex: 0 },
      { [property]: value },
    );
    assert.equal(
      engine.getRunStyle(slideIndex, shapeIndex, paragraphIndex, 0)?.[property],
      value,
      `${property} was not applied`,
    );
  }

  await engine.setParagraphAlignment(slideIndex, shapeIndex, paragraphIndex, "ctr");
  assert.equal(
    engine.getRunStyle(slideIndex, shapeIndex, paragraphIndex, 0)?.alignment,
    "ctr",
  );

  const exported = await assertExportRoundTrips("text matrix", engine);
  const { PresentationEngine } = await loadPresentationEngineModule();
  const reloaded = await PresentationEngine.load(exported);
  const style = reloaded.getRunStyle(slideIndex, shapeIndex, paragraphIndex, 0);

  assert.equal(style?.bold, true);
  assert.equal(style?.fontFamily, "Georgia");
  assert.equal(style?.fontSizePt, 31);
  assert.equal(style?.color, "112233");
  assert.equal(style?.alignment, "ctr");
  assert.match(reloaded.renderSlide(slideIndex).svg, />Matrix\s*<\/tspan>/);
  assert.match(reloaded.renderSlide(slideIndex).svg, />text<\/tspan>/);
});

test("repeating an identical font family is a true slide-XML no-op", async () => {
  const engine = await loadEngine("features.pptx");
  const text = engine.getParagraphRunText(0, 0, 0);
  const ranges = [{ paragraphIndex: 0, start: 0, end: text.length }];

  assert.equal(await engine.setRunStyleForRanges(0, 0, ranges, { fontFamily: "Georgia" }), true);
  const afterFirstApply = engine.getSlideXml(0);

  assert.equal(await engine.setRunStyleForRanges(0, 0, ranges, { fontFamily: "Georgia" }), false);
  assert.equal(
    engine.getSlideXml(0),
    afterFirstApply,
    "a repeated font pick must not rewrite the slide XML",
  );
});
