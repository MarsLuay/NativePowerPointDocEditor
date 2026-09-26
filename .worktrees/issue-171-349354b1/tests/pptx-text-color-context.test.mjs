import assert from "node:assert/strict";
import test from "node:test";
import {
  loadInsertControllerModule,
  loadNativePowerPointViewModule,
  loadTextToolbarControllerModule,
} from "./helpers/load-plugin-modules.mjs";

test("bullet toolbar toggles matching native style and removes a legacy literal marker", async () => {
  const { InsertController } = await loadInsertControllerModule();
  const commands = [];
  const host = {
    t: (key) => key,
    engine: {
      getParagraphListStyle: () => "bullet",
      getParagraphRunText: () => "• Existing list item",
    },
    app: {},
    layoutEl: null,
    selectedShapeIndex: 7,
    activeEditorTarget: null,
    currentSlide: 3,
    ensureEditable: () => true,
    captureHistoryEntry: async () => ({}),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderEditedShape: async () => true,
    renderThumbnails: async () => {},
    syncCurrentThumbnailShape: () => true,
    selectShape: () => {},
    selectShapeForTextEditing: () => {},
    startTextEditor: () => {},
    createEditIconButton: () => ({}),
    getTextEditTarget: () => null,
    getListStyleTarget: () => ({ shapeIndex: 7, paragraphIndex: 2, ranges: [] }),
    finishInlineTextEditing: async () => {},
    session: { applyCommand: async (command) => commands.push(command) },
  };
  const controller = new InsertController(host);

  await controller.applyListStyle("bullet");

  assert.deepEqual(commands, [{
    type: "apply-list-style",
    slideIndex: 3,
    shapeIndex: 7,
    paragraphIndex: 2,
    style: "none",
    stripLeadingManualBullet: true,
  }]);
});

test("bullet toolbar removes a legacy literal bullet in one click", async () => {
  const { InsertController } = await loadInsertControllerModule();
  const commands = [];
  const host = {
    t: (key) => key,
    engine: {
      getParagraphListStyle: () => "none",
      getParagraphRunText: () => "• Legacy list item",
    },
    app: {},
    layoutEl: null,
    selectedShapeIndex: 7,
    activeEditorTarget: null,
    currentSlide: 3,
    ensureEditable: () => true,
    captureHistoryEntry: async () => ({}),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderEditedShape: async () => true,
    renderThumbnails: async () => {},
    syncCurrentThumbnailShape: () => true,
    selectShape: () => {},
    selectShapeForTextEditing: () => {},
    startTextEditor: () => {},
    createEditIconButton: () => ({}),
    getTextEditTarget: () => null,
    getListStyleTarget: () => ({ shapeIndex: 7, paragraphIndex: 2, ranges: [] }),
    finishInlineTextEditing: async () => {},
    session: { applyCommand: async (command) => commands.push(command) },
  };
  const controller = new InsertController(host);

  await controller.applyListStyle("bullet");

  assert.deepEqual(commands, [{
    type: "apply-list-style",
    slideIndex: 3,
    shapeIndex: 7,
    paragraphIndex: 2,
    style: "none",
    stripLeadingManualBullet: true,
  }]);
});

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

test("list formatting captures the active paragraph and selected sentence before editor teardown", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  view.selectedShapeIndex = 2;
  view.activeShapeTextTarget = {
    kind: "shape-paragraph",
    shapeIndex: 2,
    paragraphIndex: 2,
    runIndex: 0,
    text: "Prefix selected sentence suffix",
    element: {},
    runElements: [],
  };
  view.activeEditor = { selectionStart: 7, selectionEnd: 24 };
  view.mapRangesToOoxmlOffsets = (_shapeIndex, ranges) => ranges;

  assert.deepEqual(view.getListStyleTarget(), {
    shapeIndex: 2,
    paragraphIndex: 2,
    ranges: [{ paragraphIndex: 2, start: 7, end: 24 }],
  });
});

test("list formatting preserves every selected paragraph before editor teardown", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  view.selectedShapeIndex = 2;
  view.activeShapeTextTarget = {
    kind: "shape-paragraph",
    shapeIndex: 2,
    paragraphIndex: 2,
    runIndex: 0,
    text: "Third paragraph",
    element: {},
    runElements: [],
  };
  view.getActiveInlineSelectionRanges = () => [
    { paragraphIndex: 0, start: 0, end: 5 },
    { paragraphIndex: 1, start: 0, end: 6 },
    { paragraphIndex: 2, start: 0, end: 7 },
  ];
  view.mapRangesToOoxmlOffsets = (_shapeIndex, ranges) => ranges;

  assert.deepEqual(view.getListStyleTarget(), {
    shapeIndex: 2,
    paragraphIndex: 2,
    ranges: [
      { paragraphIndex: 0, start: 0, end: 5 },
      { paragraphIndex: 1, start: 0, end: 6 },
      { paragraphIndex: 2, start: 0, end: 7 },
    ],
  });
});

test("list formatting expands a selected text box to every paragraph", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  view.selectedShapeIndex = 2;
  view.toolbarFormattingSnapshot = {
    shapeIndex: 9,
    run: { paragraphIndex: 0, runIndex: 0 },
    anchor: { left: 0, top: 0, width: 1, height: 1 },
    ranges: [{ paragraphIndex: 0, start: 0, end: 4 }],
  };
  view.getShapeTextRanges = () => [
    { paragraphIndex: 0, start: 0, end: 5 },
    { paragraphIndex: 1, start: 0, end: 6 },
    { paragraphIndex: 2, start: 0, end: 7 },
  ];
  view.mapRangesToOoxmlOffsets = (_shapeIndex, ranges) => ranges;

  assert.deepEqual(view.getListStyleTarget(), {
    shapeIndex: 2,
    paragraphIndex: 0,
    ranges: [
      { paragraphIndex: 0, start: 0, end: 5 },
      { paragraphIndex: 1, start: 0, end: 6 },
      { paragraphIndex: 2, start: 0, end: 7 },
    ],
  });
});

test("text-box list ranges combine visual wrap-lines into their OOXML paragraphs", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const lines = [
    { paragraphIndex: 0, length: 5 },
    { paragraphIndex: 0, length: 7 },
    { paragraphIndex: 1, length: 3 },
    { paragraphIndex: 2, length: 4 },
    { paragraphIndex: 2, length: 8 },
  ];

  view.svgEl = { querySelector: () => ({}) };
  view.getShapeTextParagraphs = () => lines;
  view.getParagraphIndexFromInlineElement = (line) => line.paragraphIndex;
  view.collectParagraphRuns = () => [{}];
  view.getRunCharInfo = (line) => ({ total: line.length });

  assert.deepEqual(view.getShapeTextRanges(2), [
    { paragraphIndex: 0, start: 0, end: 12 },
    { paragraphIndex: 1, start: 0, end: 3 },
    { paragraphIndex: 2, start: 0, end: 12 },
  ]);
});

test("bullet toolbar adds a bullet to every paragraph in a selected text box", async () => {
  const { InsertController } = await loadInsertControllerModule();
  const commands = [];
  const texts = ["First paragraph", "Second paragraph", "Third paragraph"];
  const host = {
    t: (key) => key,
    engine: {
      getParagraphListStyle: () => "none",
      getParagraphRunText: (_slide, _shape, paragraphIndex) => texts[paragraphIndex],
    },
    app: {},
    layoutEl: null,
    selectedShapeIndex: 7,
    activeEditorTarget: null,
    currentSlide: 3,
    ensureEditable: () => true,
    captureHistoryEntry: async () => ({}),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderEditedShape: async () => true,
    renderThumbnails: async () => {},
    syncCurrentThumbnailShape: () => true,
    selectShape: () => {},
    selectShapeForTextEditing: () => {},
    startTextEditor: () => {},
    createEditIconButton: () => ({}),
    getTextEditTarget: () => null,
    getListStyleTarget: () => ({
      shapeIndex: 7,
      paragraphIndex: 0,
      ranges: texts.map((text, paragraphIndex) => ({
        paragraphIndex,
        start: 0,
        end: text.length,
      })),
    }),
    finishInlineTextEditing: async () => {},
    session: { applyCommand: async (command) => commands.push(command) },
  };
  const controller = new InsertController(host);

  await controller.applyListStyle("bullet");

  assert.deepEqual(commands, [{
    type: "apply-list-style-ranges",
    slideIndex: 3,
    shapeIndex: 7,
    ranges: texts.map((text, paragraphIndex) => ({
      paragraphIndex,
      start: 0,
      end: text.length,
    })),
    style: "bullet",
    stripLeadingManualBullet: false,
  }]);
});

test("bullet toolbar removes one matching style across every selected paragraph", async () => {
  const { InsertController } = await loadInsertControllerModule();
  const commands = [];
  const host = {
    t: (key) => key,
    engine: {
      getParagraphListStyle: () => "bullet",
      getParagraphRunText: () => "Existing list item",
    },
    app: {},
    layoutEl: null,
    selectedShapeIndex: 7,
    activeEditorTarget: null,
    currentSlide: 3,
    ensureEditable: () => true,
    captureHistoryEntry: async () => ({}),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderEditedShape: async () => true,
    renderThumbnails: async () => {},
    syncCurrentThumbnailShape: () => true,
    selectShape: () => {},
    selectShapeForTextEditing: () => {},
    startTextEditor: () => {},
    createEditIconButton: () => ({}),
    getTextEditTarget: () => null,
    getListStyleTarget: () => ({
      shapeIndex: 7,
      paragraphIndex: 2,
      ranges: [
        { paragraphIndex: 0, start: 0, end: 5 },
        { paragraphIndex: 1, start: 0, end: 6 },
        { paragraphIndex: 2, start: 0, end: 7 },
      ],
    }),
    finishInlineTextEditing: async () => {},
    session: { applyCommand: async (command) => commands.push(command) },
  };
  const controller = new InsertController(host);

  await controller.applyListStyle("bullet");

  assert.deepEqual(commands, [{
    type: "apply-list-style-ranges",
    slideIndex: 3,
    shapeIndex: 7,
    ranges: [
      { paragraphIndex: 0, start: 0, end: "Existing list item".length },
      { paragraphIndex: 1, start: 0, end: "Existing list item".length },
      { paragraphIndex: 2, start: 0, end: "Existing list item".length },
    ],
    style: "none",
    stripLeadingManualBullet: false,
  }]);
});

test("bullet toolbar removes whole paragraphs when a cross-paragraph drag starts mid-line", async () => {
  const { InsertController } = await loadInsertControllerModule();
  const commands = [];
  const texts = ["First paragraph", "Second paragraph", "Third paragraph"];
  const host = {
    t: (key) => key,
    engine: {
      getParagraphListStyle: () => "bullet",
      getParagraphRunText: (_slide, _shape, paragraphIndex) => texts[paragraphIndex],
    },
    app: {},
    layoutEl: null,
    selectedShapeIndex: 7,
    activeEditorTarget: null,
    currentSlide: 3,
    ensureEditable: () => true,
    captureHistoryEntry: async () => ({}),
    recordHistoryEntry: () => {},
    markDirty: () => {},
    renderCurrentSlide: async () => true,
    renderEditedShape: async () => true,
    renderThumbnails: async () => {},
    syncCurrentThumbnailShape: () => true,
    selectShape: () => {},
    selectShapeForTextEditing: () => {},
    startTextEditor: () => {},
    createEditIconButton: () => ({}),
    getTextEditTarget: () => null,
    getListStyleTarget: () => ({
      shapeIndex: 7,
      paragraphIndex: 2,
      ranges: [
        { paragraphIndex: 0, start: 5, end: texts[0].length },
        { paragraphIndex: 1, start: 0, end: texts[1].length },
        { paragraphIndex: 2, start: 0, end: 5 },
      ],
    }),
    finishInlineTextEditing: async () => {},
    session: { applyCommand: async (command) => commands.push(command) },
  };
  const controller = new InsertController(host);

  await controller.applyListStyle("bullet");

  assert.deepEqual(commands, [{
    type: "apply-list-style-ranges",
    slideIndex: 3,
    shapeIndex: 7,
    ranges: texts.map((text, paragraphIndex) => ({
      paragraphIndex,
      start: 0,
      end: text.length,
    })),
    style: "none",
    stripLeadingManualBullet: false,
  }]);
});

test("clicking off a text box hides formatting context when no popover is open", async () => {
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
      if (selector === 'g[data-ooxml-shape-idx="11"]') return shape;
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

  view.selectedShapeIndex = 11;
  view.selectedShapeIndices = new Set([11]);
  view.textToolbarShapeIndex = 11;
  view.textToolbarController.textToolbarShapeIndex = 11;
  view.textToolbarController.toolbarFormattingSnapshot = {
    shapeIndex: 11,
    run: { paragraphIndex: 0, runIndex: 0 },
    anchor: { left: 10, top: 20, width: 100, height: 40 },
    ranges: null,
  };
  // No active popover — click-off must drop context so the floating toolbar hides.
  view.textToolbarController.activeToolbarPopover = null;

  view.clearSelection({ skipTextCommit: true });

  assert.equal(view.selectedShapeIndex, null);
  assert.equal(view.textToolbarController.getFormattingSnapshot(), null);
  assert.equal(view.getTextStyleContext(), null);
});
