import { test } from "node:test";
import {
  assertExportRoundTrips,
  assertShapeIndex,
  loadEngine,
  readFixtureImage,
} from "./helpers/pptx-action-harness.mjs";

const SLIDE_INDEX = 0;

const actions = [
  {
    label: "shape",
    insert: (engine) => engine.addShapeGeometry(SLIDE_INDEX, "rect"),
  },
  {
    label: "image",
    insert: async (engine) =>
      engine.addImage(SLIDE_INDEX, await readFixtureImage(), "image/png"),
  },
  {
    label: "table",
    insert: (engine) => engine.addTable(SLIDE_INDEX, 3, 4),
  },
  {
    label: "chart",
    insert: (engine) => engine.addChart(SLIDE_INDEX),
  },
  {
    label: "text",
    insert: (engine) => engine.addTextBox(SLIDE_INDEX),
  },
  {
    label: "list",
    insert: async (engine) => {
      const shapeIndex = await engine.addTextBox(SLIDE_INDEX);
      await engine.applyListStyle(SLIDE_INDEX, shapeIndex, 0, "bullet");
      return shapeIndex;
    },
  },
];

for (const { label, insert } of actions) {
  test(`insert ${label} action round-trips`, async () => {
    const engine = await loadEngine();
    const shapeIndex = await insert(engine);
    assertShapeIndex(label, shapeIndex);
    await assertExportRoundTrips(label, engine);
  });
}
