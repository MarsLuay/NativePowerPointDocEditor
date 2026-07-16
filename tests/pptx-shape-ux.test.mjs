import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { test } from "node:test";
import { loadNativePowerPointViewModule } from "./helpers/load-plugin-modules.mjs";

function createView(NativePowerPointView) {
  return new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
}

test("text boxes keep rendered text unchanged during resize preview", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const appliedTransforms = [];
  const previewElement = {
    setAttribute: (name, value) => appliedTransforms.push([name, value]),
  };
  const startTransform = { x: 10, y: 20, cx: 100, cy: 50, rot: 0 };

  view.engine = { getSlideScale: () => 1 };
  view.svgEl = {};
  view.dragState = {
    mode: "resize",
    handle: "se",
    startTransform,
    previewElement,
    previewOriginalTransform: null,
    freezeShapeDuringResize: true,
  };

  view.updateShapeTransformPreview({ x: 10, y: 20, cx: 150, cy: 75, rot: 0 });
  assert.deepEqual(appliedTransforms, [], "resize must not scale a text-box SVG group");

  view.dragState.mode = "move";
  view.updateShapeTransformPreview({ x: 20, y: 30, cx: 100, cy: 50, rot: 0 });
  assert.equal(appliedTransforms.length, 1, "moving a text box should retain the existing live preview");

  appliedTransforms.length = 0;
  view.dragState.mode = "resize";
  view.dragState.freezeShapeDuringResize = false;
  view.updateShapeTransformPreview({ x: 10, y: 20, cx: 150, cy: 75, rot: 0 });
  assert.equal(appliedTransforms.length, 1, "non-text shapes should retain live resize preview");

  let outlineUpdates = 0;
  appliedTransforms.length = 0;
  view.engine = { getSlideScale: () => 1, pxToEmu: (value) => value };
  view.dragState = {
    mode: "resize",
    handle: "se",
    pointerId: 7,
    startPoint: { x: 0, y: 0 },
    startClientX: 0,
    startClientY: 0,
    startBox: { left: 0, top: 0, width: 100, height: 50 },
    startTransform,
    latestTransform: startTransform,
    previewElement,
    previewOriginalTransform: null,
    freezeShapeDuringResize: true,
  };
  view.getSvgPoint = () => ({ x: 25, y: 15 });
  view.getSelectedShapeElement = () => previewElement;
  view.isPictureShape = () => false;
  view.shapeHasRotation = () => false;
  view.updateSelectionOverlayDuringDrag = () => { outlineUpdates += 1; };

  view.handleDragMove({ pointerId: 7, clientX: 25, clientY: 15 });
  assert.equal(outlineUpdates, 1, "outline and resize handles must continue to track the drag");
  assert.deepEqual(appliedTransforms, [], "the text box itself must remain unchanged until commit");
});

test("every rendered text shape freezes during resize preview", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const textShape = { querySelector: (selector) => selector === "text" ? {} : null };
  const nonTextShape = { querySelector: () => null };

  assert.equal(view.shouldFreezeTextDuringResize("resize", textShape), true);
  assert.equal(view.shouldFreezeTextDuringResize("resize", nonTextShape), false);
  assert.equal(view.shouldFreezeTextDuringResize("move", textShape), false);
});

test("selected PowerPoint shapes use an outline without a glow", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  const selectedShapeRule = css.match(/\.native-powerpoint-slide-svg \.native-powerpoint-shape-selected\s*\{([\s\S]*?)\}/);

  assert.ok(selectedShapeRule, "the selected-shape rule must remain explicit for the selection state");
  assert.match(selectedShapeRule[1], /filter:\s*none/);
  assert.doesNotMatch(selectedShapeRule[1], /drop-shadow/);
});

test("inline text edits restore a canvas position changed by native caret reveal", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const restored = [];
  view.canvasPane = { scrollLeft: 12, scrollTop: 34 };
  view.restoreCanvasScrollSoon = (position) => restored.push(position);

  view.canvasPane.scrollLeft = 2;
  view.canvasPane.scrollTop = 3;
  view.preserveCanvasScrollAfterInlineTextEdit({ left: 12, top: 34 }, "deleteContentBackward");

  assert.deepEqual(restored, [{ left: 12, top: 34 }]);
});

test("PowerPoint canvas disables browser scroll anchoring during text preview updates", async () => {
  const css = await readFile(resolve("styles.css"), "utf8");
  const canvasPaneRule = css.match(/\.native-powerpoint-canvas-pane\s*\{([\s\S]*?)\}/);

  assert.ok(canvasPaneRule, "the canvas pane must remain an explicit scroll owner");
  assert.match(canvasPaneRule[1], /overflow-anchor:\s*none/);
});

test("shape fill menu uses the shared color popover and excludes unsupported objects", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = createView(NativePowerPointView);
  const items = [];
  const menu = {
    addSeparator() { return menu; },
    addItem(build) {
      const item = {
        disabled: false,
        icon: null,
        title: null,
        action: null,
        setDisabled(value) { this.disabled = value; return this; },
        setIcon(value) { this.icon = value; return this; },
        setTitle(value) { this.title = value; return this; },
        onClick(value) { this.action = value; return this; },
      };
      build(item);
      items.push(item);
      return menu;
    },
  };
  const opened = [];

  view.t = (key) => key === "powerpoint:contextMenu.shapeFillColor" ? "Shape fill color…" : key;
  view.engine = { canSetShapeFillColor: () => true };
  view.openShapeFillColorPicker = (shapeIndex) => opened.push(shapeIndex);
  view.addShapeFillColorMenuItem(menu, 4, true);

  assert.equal(items.length, 1);
  assert.equal(items[0].title, "Shape fill color…");
  assert.equal(items[0].icon, "palette");
  assert.equal(items[0].disabled, false);
  items[0].action();
  assert.deepEqual(opened, [4]);

  view.engine = { canSetShapeFillColor: () => false };
  view.addShapeFillColorMenuItem(menu, 5, true);
  assert.equal(items.length, 1, "pictures, charts, tables, groups, and connectors should not offer fill color");
});
