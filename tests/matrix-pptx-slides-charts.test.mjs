import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertSlideCount,
  loadEngine,
} from "./helpers/pptx-action-harness.mjs";

function chartShapeIndices(engine, slideIndex = 0) {
  return [...engine.renderSlide(slideIndex).svg.matchAll(
    /data-ooxml-shape-type="chart"[^>]*data-ooxml-shape-idx="(\d+)"|data-ooxml-shape-idx="(\d+)"[^>]*data-ooxml-shape-type="chart"/g,
  )]
    .map((match) => Number(match[1] ?? match[2]))
    .filter(Number.isInteger);
}

test("chart data matrix: edit an editable chart when the engine exposes one", async (t) => {
  const engine = await loadEngine("features.pptx");
  assert.equal(typeof engine.getChartDataGrid, "function", "chart data read API must exist");
  assert.equal(typeof engine.updateChartData, "function", "chart data edit API must exist");

  const shapeIndex = chartShapeIndices(engine).find((index) => engine.getChartDataGrid(0, index));
  assert.notEqual(shapeIndex, undefined, "features fixture must expose a chart data descriptor");
  const grid = engine.getChartDataGrid(0, shapeIndex);
  assert.ok(grid, "chart descriptor must resolve");

  if (!grid.editable) {
    t.skip(`fixture chart is read-only: ${grid.reason}`);
    return;
  }

  await engine.updateChartData(0, shapeIndex, {
    categories: grid.categories,
    series: grid.series.map(({ values, pointLabels }) => ({ values, pointLabels })),
  });
  await assertExportRoundTrips("chart data matrix", engine);
});

test("slide matrix: add, duplicate, delete, reorder, and round-trip", async () => {
  const engine = await loadEngine();
  assertSlideCount(engine, 1);

  const added = await engine.addSlide(0);
  assert.equal(added.slideCount, 2);
  assertSlideCount(engine, 2);

  const duplicated = await engine.duplicateSlide(0);
  assert.equal(duplicated.slideCount, 3);
  assertSlideCount(engine, 3);

  const moved = await engine.moveSlide(0, 1);
  assert.equal(moved.slideIndex, 1);
  assert.equal(moved.slideCount, 3);
  assertSlideCount(engine, 3);

  const deleted = await engine.deleteSlide(2);
  assert.equal(deleted.slideCount, 2);
  assertSlideCount(engine, 2);

  await assertExportRoundTrips("slide matrix", engine, { expectedSlides: 2 });
});
