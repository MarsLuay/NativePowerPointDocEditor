import assert from "node:assert/strict";
import test from "node:test";
import {
  loadNativePowerPointViewModule,
  loadTextToolbarControllerModule,
} from "./helpers/load-plugin-modules.mjs";

test("color formatting keeps context after selection clears while the popover is open", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  const shape = {
    querySelector(selector) {
      if (selector === "tspan[data-ooxml-run-idx]") {
        return {
          getAttribute: () => "0",
          closest: () => ({ getAttribute: () => "0" }),
        };
      }
      return null;
    },
    closest() {
      return null;
    },
  };
  view.engine = {};
  view.canEdit = () => true;
  view.svgEl = {
    querySelector(selector) {
      if (selector === 'g[data-ooxml-shape-idx="53"]') return shape;
      return null;
    },
    querySelectorAll() {
      return [];
    },
  };
  view.session = { clearSelection() {} };
  view.getElementBox = () => ({ left: 10, top: 20, width: 100, height: 40 });
  view.getSelectedBox = () => null;
  view.shapeHasEditableText = () => true;
  view.removeSelectionOverlay = () => {};
  view.removeMultiSelectionBoxes = () => {};
  view.removeMarqueeSelectionPreview = () => {};
  view.snapController = { clearSnapGuides() {} };
  view.renderInspector = () => {};
  view.updateObjectClipboardAvailability = () => {};
  view.updateTextToolbar = () => {};

  view.selectedShapeIndex = 53;
  view.selectedShapeIndices = new Set([53]);
  view.textToolbarController.toolbarFormattingSnapshot = {
    shapeIndex: 53,
    run: { paragraphIndex: 0, runIndex: 0 },
    anchor: { left: 10, top: 20, width: 100, height: 40 },
    ranges: [{ paragraphIndex: 0, start: 0, end: 5 }],
  };
  view.textToolbarController.activeToolbarPopover = { isConnected: true };

  view.clearSelection({ skipTextCommit: true });

  assert.equal(view.selectedShapeIndex, null);
  assert.ok(view.textToolbarController.getFormattingSnapshot());
  const context = view.getTextStyleContext();
  assert.equal(context?.shapeIndex, 53);
  assert.equal(context?.run?.runIndex, 0);
});

test("preserveFormattingContext stores the live text target before the editor closes", async () => {
  const { TextToolbarController } = await loadTextToolbarControllerModule();
  const host = {
    engine: {},
    activeEditor: { selectionStart: 0, selectionEnd: 4, value: "test" },
    activeTextStyleTarget: {
      shapeIndex: 7,
      paragraphIndex: 0,
      runIndex: 0,
      element: {},
    },
    getTextStyleContext() {
      return {
        shapeIndex: 7,
        run: { paragraphIndex: 0, runIndex: 0 },
        anchor: { left: 1, top: 2, width: 3, height: 4 },
      };
    },
    getStoredInlineSelectionRanges() {
      return null;
    },
  };
  const controller = new TextToolbarController(host);
  controller.preserveFormattingContext();

  const snapshot = controller.getFormattingSnapshot();
  assert.equal(snapshot?.shapeIndex, 7);
  assert.deepEqual(snapshot?.ranges, [{ paragraphIndex: 0, start: 0, end: 4 }]);
});

test("closing the inline editor does not wipe a preserved formatting snapshot", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  view.toolbarFormattingSnapshot = {
    shapeIndex: 9,
    run: { paragraphIndex: 0, runIndex: 0 },
    anchor: { left: 0, top: 0, width: 1, height: 1 },
    ranges: [{ paragraphIndex: 0, start: 1, end: 3 }],
  };
  view.clearWholeShapeInlineSelection();
  assert.equal(view.toolbarFormattingSnapshot?.shapeIndex, 9);
});
