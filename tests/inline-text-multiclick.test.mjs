import assert from "node:assert/strict";
import { test } from "node:test";
import {
  loadNativePowerPointViewModule,
  loadTextUtilsModule,
} from "./helpers/load-plugin-modules.mjs";

test("double-click word ranges include the word under the caret", async () => {
  const { getInlineWordRange } = await loadTextUtilsModule();

  assert.deepEqual(getInlineWordRange("hello world", 8), { start: 6, end: 11 });
  assert.deepEqual(getInlineWordRange("can't stop", 3), { start: 0, end: 5 });
  assert.deepEqual(getInlineWordRange("naïve café", 9), { start: 6, end: 10 });
  assert.deepEqual(getInlineWordRange("hello", 5), { start: 0, end: 5 });
});

test("inline text multi-click selects a word, then the whole text box", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const editor = {
    value: "hello world",
    selectionStart: 0,
    selectionEnd: 0,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const element = {};
  let wholeTextBoxSelections = 0;
  let toolbarRefreshes = 0;

  view.activeShapeTextTarget = { shapeIndex: 7, paragraphIndex: 2 };
  view.stopInlineSelectionDrag = () => {};
  view.clearWholeShapeInlineSelection = () => {};
  view.updateInlineCaret = () => {};
  view.updateTextToolbar = () => {
    toolbarRefreshes += 1;
  };
  view.selectAllInlineText = () => {
    wholeTextBoxSelections += 1;
  };

  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 2), true);
  assert.deepEqual(
    { start: editor.selectionStart, end: editor.selectionEnd },
    { start: 6, end: 11 },
  );
  assert.deepEqual(view.inlineRangeSelection, {
    shapeIndex: 7,
    ranges: [{ paragraphIndex: 2, start: 6, end: 11 }],
  });
  assert.equal(toolbarRefreshes, 1);

  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 3), true);
  assert.equal(wholeTextBoxSelections, 1);
  assert.equal(toolbarRefreshes, 2);
  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 1), false);
});

test("inline caret is created by the owning SVG document", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const attributes = new Map();
  const caret = {
    classList: { add() {} },
    setAttribute(name, value) {
      attributes.set(name, value);
    },
  };
  const calls = [];
  let appended = null;
  view.svgEl = {
    ownerDocument: {
      win: {
        createSvg(name) {
          calls.push({ name });
          return caret;
        },
      },
    },
    appendChild(node) {
      appended = node;
    },
  };

  assert.equal(view.createInlineCaret(), caret);
  assert.deepEqual(calls, [{ name: 'line' }]);
  assert.equal(appended, caret);
  assert.equal(attributes.get('aria-hidden'), 'true');
});

test("inline text selection captures and releases its pointer on the SVG canvas", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const captured = [];
  const released = [];
  view.svgEl = {
    setPointerCapture(pointerId) {
      captured.push(pointerId);
    },
    hasPointerCapture(pointerId) {
      return captured.includes(pointerId);
    },
    releasePointerCapture(pointerId) {
      released.push(pointerId);
    },
  };

  assert.equal(view.captureInlineSelectionPointer(19), true);
  view.releaseInlineSelectionPointer(19);

  assert.deepEqual(captured, [19]);
  assert.deepEqual(released, [19]);
});

test("a pointer-captured inline click cannot clear the active text editor", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  let prevented = false;
  let stopped = false;
  view.suppressNextTextClick = true;

  const consumed = view.consumeInlineTextClick({
    detail: 1,
    preventDefault() { prevented = true; },
    stopPropagation() { stopped = true; },
  }, null);

  assert.equal(consumed, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.equal(view.suppressNextTextClick, false);
});

test("a pointer-captured double-click resolves back to its active text target", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const activeTarget = { shapeIndex: 7, paragraphIndex: 2, element: {} };
  let selection = null;
  view.suppressNextTextClick = true;
  view.activeShapeTextTarget = activeTarget;
  view.applyInlineMultiClickSelectionAtPoint = (...args) => {
    selection = args;
  };

  view.consumeInlineTextClick({
    detail: 2,
    clientX: 100,
    clientY: 200,
    preventDefault() {},
    stopPropagation() {},
  }, null);

  assert.deepEqual(selection, [activeTarget, 100, 200, 2]);
});

test("live preview rebinds detached paragraph runs before deleting text", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const detachedRun = { textContent: "stale text" };
  const liveRun = { textContent: "visible text" };
  const line = { classList: { add() {}, remove() {} } };
  const target = {
    kind: "shape-paragraph",
    shapeIndex: 7,
    paragraphIndex: 2,
    runIndex: 0,
    text: "visible text",
    element: {},
    runElements: [detachedRun],
  };

  view.getRunLineContainers = () => [line];
  view.collectParagraphRuns = () => [liveRun];
  view.reflowShapeParagraphPreview = () => false;
  view.getParagraphPlainText = () => liveRun.textContent;

  view.syncShapeParagraphPreview(target, "visible tex");

  assert.equal(liveRun.textContent, "visible tex");
  assert.equal(detachedRun.textContent, "stale text");
  assert.equal(target.element, line);
  assert.deepEqual(target.runElements, [liveRun]);
});

test("live preview repairs an SVG reflow mismatch before text commit", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const liveRun = { textContent: "visible text" };
  const line = { classList: { add() {}, remove() {} } };
  const target = {
    kind: "shape-paragraph",
    shapeIndex: 7,
    paragraphIndex: 2,
    runIndex: 0,
    text: "visible text",
    element: line,
    runElements: [liveRun],
  };

  view.getRunLineContainers = () => [line];
  view.collectParagraphRuns = () => [liveRun];
  view.getParagraphPlainText = () => liveRun.textContent;
  view.reflowShapeParagraphPreview = () => true;

  view.syncShapeParagraphPreview(target, "visible tex");

  assert.equal(liveRun.textContent, "visible tex");
});

test("full selection replacement repaints the owning SVG text frame immediately", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const liveRun = { textContent: "old text" };
  const line = { classList: { add() {}, remove() {} } };
  const target = {
    kind: "shape-paragraph",
    shapeIndex: 7,
    paragraphIndex: 2,
    runIndex: 0,
    text: "old text",
    element: line,
    runElements: [liveRun],
  };
  let textFrameReplacements = 0;

  view.getRunLineContainers = () => [line];
  view.collectParagraphRuns = () => [liveRun];
  view.getParagraphPlainText = () => liveRun.textContent;
  view.reflowShapeParagraphPreview = () => false;
  view.replaceLiveShapeTextFrame = () => {
    textFrameReplacements += 1;
    return true;
  };

  const replaced = view.syncShapeParagraphPreview(target, "new text", {
    replaceTextFrame: true,
  });

  assert.equal(liveRun.textContent, "new text");
  assert.equal(textFrameReplacements, 1);
  assert.equal(replaced, true);
});

test("live preview replaces connected SVG runs so deletion paints before commit", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const replacement = { textContent: "old text" };
  let replacedWith = null;
  const liveRun = {
    textContent: "old text",
    isConnected: true,
    parentNode: {},
    cloneNode() {
      return replacement;
    },
    replaceWith(next) {
      replacedWith = next;
    },
  };

  const result = view.replaceLiveParagraphRunText(liveRun, "visible text");

  assert.equal(replacedWith, replacement);
  assert.equal(replacement.textContent, "visible text");
  assert.equal(result, replacement);
});

test("flushing an unedited inline editor does not rewrite rendered paragraph text", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const applied = [];
  const editor = { value: "Heading\nBody" };
  const target = { text: "HeadingBody" };

  view.activeEditor = editor;
  view.activeTextStyleTarget = target;
  view.activeEditorTextDirty = false;
  view.removeActiveEditor = () => {};
  view.applyTextValue = (...args) => {
    applied.push(args);
    return Promise.resolve();
  };

  view.flushActiveEditor();

  assert.deepEqual(applied, []);
});

test("Backspace at a paragraph start requests structural removal of an empty predecessor", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "Current paragraph",
    selectionStart: 0,
    selectionEnd: 0,
  };
  const target = { element, shapeIndex: 4, paragraphIndex: 2 };
  let requested = null;
  let prevented = false;
  let stopped = false;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.engine = {
    getParagraphRunText: () => "",
    hasEmptyPrecedingParagraph: () => true,
  };
  view.startInlineEmptyParagraphRemoval = (receivedEditor, receivedTarget) => {
    requested = { editor: receivedEditor, target: receivedTarget };
  };

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    preventDefault: () => { prevented = true; },
    stopPropagation: () => { stopped = true; },
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.deepEqual(requested, { editor, target });
});

test("Backspace at a paragraph start merges a non-empty predecessor", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "Current paragraph",
    selectionStart: 0,
    selectionEnd: 0,
  };
  const target = { element, shapeIndex: 4, paragraphIndex: 2 };
  let requested = null;
  let prevented = false;
  let stopped = false;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.engine = { hasEmptyPrecedingParagraph: () => false };
  view.startInlinePrecedingParagraphMerge = (receivedEditor, receivedTarget) => {
    requested = { editor: receivedEditor, target: receivedTarget };
  };

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    preventDefault: () => { prevented = true; },
    stopPropagation: () => { stopped = true; },
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.deepEqual(requested, { editor, target });
});

test("Delete routes a visual cross-paragraph selection to the OOXML range mutation", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "Current paragraph",
    selectionStart: 0,
    selectionEnd: 0,
  };
  const target = { element, shapeIndex: 4, paragraphIndex: 2 };
  let requested = null;
  let prevented = false;
  let stopped = false;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.inlineRangeSelection = {
    shapeIndex: 4,
    ranges: [
      { paragraphIndex: 1, start: 12, end: 30 },
      { paragraphIndex: 2, start: 0, end: 8 },
    ],
  };
  view.startInlineRangeDeletion = (receivedEditor, receivedTarget) => {
    requested = { editor: receivedEditor, target: receivedTarget };
  };

  const handled = view.handleInlineDeleteKey({
    key: "Delete",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    preventDefault: () => { prevented = true; },
    stopPropagation: () => { stopped = true; },
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.deepEqual(requested, { editor, target });
});

test("cross-paragraph drag keeps the anchor paragraph's wrapped visual lines selected", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const line = (paragraphIndex, length, text) => ({ paragraphIndex, length, textContent: text });
  const paragraphs = [
    line(0, 5, 'Title'),
    line(1, 55, 'A'.repeat(55)),
    line(1, 47, 'B'.repeat(47)),
    line(1, 22, 'C'.repeat(22)),
    line(2, 59, 'D'.repeat(59)),
    line(2, 58, 'E'.repeat(58)),
    line(2, 56, 'F'.repeat(56)),
    line(2, 15, 'G'.repeat(15)),
  ];
  const renderedRects = [];

  view.activeShapeTextTarget = { shapeIndex: 42 };
  view.svgEl = { querySelector: () => ({}) };
  view.getShapeTextParagraphs = () => paragraphs;
  view.getParagraphIndexFromInlineElement = (element) => element.paragraphIndex;
  view.collectParagraphRuns = (element) => [{ textContent: element.textContent }];
  view.getRunCharInfo = (element) => ({ total: element.length });
  view.runOffsetToGeometryIndex = (_element, offset) => offset;
  view.removeInlineSelection = () => {};
  view.renderInlineSelectionRects = (element, start, end) => {
    renderedRects.push({ paragraphIndex: element.paragraphIndex, start, end });
  };
  view.updateTextToolbar = () => {};

  // The editor anchors against paragraph 2's first visual tspan, but the
  // flat editor offset belongs near its final wrapped line. Dragging backward
  // into paragraph 1 must retain all paragraph-2 lines through that offset.
  view.renderCrossParagraphSelection(paragraphs, 4, 187, 1, 0);

  assert.deepEqual(view.inlineRangeSelection, {
    shapeIndex: 42,
    ranges: [
      { paragraphIndex: 1, start: 0, end: 124 },
      { paragraphIndex: 2, start: 0, end: 187 },
    ],
  });
  assert.deepEqual(renderedRects, [
    { paragraphIndex: 1, start: 0, end: 55 },
    { paragraphIndex: 1, start: 0, end: 47 },
    { paragraphIndex: 1, start: 0, end: 22 },
    { paragraphIndex: 2, start: 0, end: 59 },
    { paragraphIndex: 2, start: 0, end: 58 },
    { paragraphIndex: 2, start: 0, end: 56 },
    { paragraphIndex: 2, start: 0, end: 14 },
  ]);
});

test("Delete keeps a selected inline paragraph and its SVG preview in sync", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "MONet",
    selectionStart: 0,
    selectionEnd: 5,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  let preview = null;
  let prevented = false;
  let stopped = false;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.recordInlineEditSnapshot = () => {};
  view.syncShapeParagraphPreview = (_target, text, options) => {
    preview = { text, options };
    return true;
  };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};

  const handled = view.handleInlineDeleteKey({
    key: "Delete",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault: () => { prevented = true; },
    stopPropagation: () => { stopped = true; },
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.equal(editor.value, "");
  assert.deepEqual({ start: editor.selectionStart, end: editor.selectionEnd }, { start: 0, end: 0 });
  assert.deepEqual(preview, {
    text: "",
    options: { replaceTextFrame: true },
  });
  assert.equal(view.activeEditorTextDirty, true);
});

test("native full deletion and a final-character deletion request the SVG text-frame repaint", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));

  assert.equal(
    view.shouldRefreshInlineTextFrameForNativeInput("deleteContentBackward", "MONet", 0, 5),
    true,
  );
  assert.equal(
    view.shouldRefreshInlineTextFrameForNativeInput("deleteContentBackward", "MONet", 2, 2),
    false,
  );
  assert.equal(
    view.shouldRefreshInlineTextFrameForNativeInput("deleteContentBackward", "M", 1, 1, ""),
    true,
  );
  // Electron can omit beforeinput/inputType for an otherwise normal native
  // textarea deletion. The value shrink itself must still force a frame repaint.
  assert.equal(
    view.shouldRefreshInlineTextFrameForNativeInput(null, "MONet", 2, 2, "MOet"),
    true,
  );
  assert.equal(
    view.shouldRefreshInlineTextFrameForNativeInput("insertText", "MONet", 0, 5),
    true,
  );
});

test("metadata-free native input repaints the SVG frame before blur", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const listeners = new Map();
  const editor = {
    value: "What is MONet?",
    selectionStart: 0,
    selectionEnd: 0,
    addEventListener(type, listener) {
      listeners.set(type, listener);
    },
    setCssProps() {},
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const run = { textContent: "What is MONet?" };
  const element = {
    isConnected: true,
    classList: { add() {}, remove() {} },
  };
  const target = {
    kind: "shape-paragraph",
    shapeIndex: 42,
    paragraphIndex: 1,
    runIndex: 0,
    text: "What is MONet?",
    element,
    runElements: [run],
  };
  const previousWindow = globalThis.window;
  let preview = null;

  globalThis.window = {
    getComputedStyle: () => ({
      fill: "#000",
      fontFamily: "Arial",
      fontSize: "12px",
      fontStyle: "normal",
      fontWeight: "400",
      textAnchor: "start",
    }),
    requestAnimationFrame(callback) {
      callback();
      return 0;
    },
  };
  try {
    view.canvasPane = { createEl: () => editor };
    view.selectedShapeIndex = 42;
    view.ensureEditable = () => true;
    view.getSelectedShapeElement = () => ({});
    view.getElementBox = () => ({ left: 0, top: 0, width: 100, height: 20 });
    view.getScreenFontSize = () => 12;
    view.getRunLineContainers = () => [element];
    view.getParagraphPlainText = () => "";
    view.positionTextRunEditor = () => {};
    view.updateSelectionOverlay = () => {};
    view.createInlineCaret = () => null;
    view.focusEditorWithoutCanvasScroll = () => {};
    view.placeInlineCaret = () => {};
    view.updateTextToolbar = () => {};
    view.rememberCollapsedInlineCaretPlacement = () => {};
    view.updateInlineCaret = () => {};
    view.preserveCanvasScrollAfterInlineTextEdit = () => {};
    view.syncShapeParagraphPreview = (_target, text, options) => {
      preview = { text, options };
      return true;
    };

    view.startTextEditor(target);
    editor.value = "";
    listeners.get("input")();

    assert.deepEqual(preview, {
      text: "",
      options: { replaceTextFrame: true },
    });
    assert.equal(view.activeEditorTextDirty, true);
  } finally {
    globalThis.window = previousWindow;
  }
});

test("Backspace clears the final inline character through the owning SVG text frame", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "M",
    selectionStart: 1,
    selectionEnd: 1,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  let preview = null;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.recordInlineEditSnapshot = () => {};
  view.syncShapeParagraphPreview = (_target, text, options) => {
    preview = { text, options };
    return true;
  };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {},
    stopPropagation() {},
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(editor.value, "");
  assert.deepEqual(preview, {
    text: "",
    options: { replaceTextFrame: true },
  });
});

test("Backspace on an already-empty first paragraph keeps the text box in the edit session", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "",
    selectionStart: 0,
    selectionEnd: 0,
    setSelectionRange() {},
  };
  const target = { element, shapeIndex: 37, paragraphIndex: 0 };
  let prevented = false;
  let stopped = false;
  let deletedShape = false;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.selectedShapeIndex = 37;
  view.deleteSelectedShape = async () => {
    deletedShape = true;
  };

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {
      prevented = true;
    },
    stopPropagation() {
      stopped = true;
    },
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(prevented, true);
  assert.equal(stopped, true);
  assert.equal(editor.value, "");
  assert.equal(view.activeEditor, editor);
  assert.equal(deletedShape, false);
});

test("empty inline preview keeps a render anchor so the SVG text frame stays alive", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  view.activeEditor = { value: "" };

  assert.equal(view.resolveInlinePreviewDomText(""), "\u200B");
  assert.equal(view.resolveInlinePreviewDomText("hi"), "hi");
  view.activeEditor = null;
  assert.equal(view.resolveInlinePreviewDomText(""), "");
});

test("Backspace repaints a partial inline deletion through the owning SVG text frame", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "MONet",
    selectionStart: 3,
    selectionEnd: 3,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  let preview = null;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.recordInlineEditSnapshot = () => {};
  view.syncShapeParagraphPreview = (_target, text, options) => {
    preview = { text, options };
    return true;
  };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {},
    stopPropagation() {},
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(editor.value, "MOet");
  assert.equal(view.activeEditorCanonicalText, "MOet");
  assert.deepEqual(preview, {
    text: "MOet",
    options: { replaceTextFrame: true },
  });
});

test("an active-paragraph Ctrl/Cmd+A range deletes synchronously before OOXML commit", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "MONet",
    selectionStart: 0,
    selectionEnd: 5,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  let preview = null;
  let rangeDeletionStarts = 0;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.inlineRangeSelection = {
    shapeIndex: 42,
    ranges: [{ paragraphIndex: 1, start: 0, end: 5 }],
  };
  view.recordInlineEditSnapshot = () => {};
  view.syncShapeParagraphPreview = (_target, text, options) => {
    preview = { text, options };
    return true;
  };
  view.startInlineRangeDeletion = () => { rangeDeletionStarts += 1; };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};

  const handled = view.handleInlineDeleteKey({
    key: "Delete",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {},
    stopPropagation() {},
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(rangeDeletionStarts, 0);
  assert.equal(editor.value, "");
  assert.deepEqual(preview, {
    text: "",
    options: { replaceTextFrame: true },
  });
});

test("deleting a full text-box selection stays in the inline replacement session", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "soil + water",
    selectionStart: 0,
    selectionEnd: 12,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  const ranges = [
    { paragraphIndex: 0, start: 0, end: 5 },
    { paragraphIndex: 1, start: 0, end: 12 },
    { paragraphIndex: 2, start: 0, end: 4 },
  ];
  let replacementStarts = 0;
  let structuralDeletes = 0;
  let preview = null;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.inlineRangeSelection = { shapeIndex: 42, ranges };
  view.getShapeTextRanges = () => ranges;
  view.beginInlineWholeShapeReplacement = () => {
    replacementStarts += 1;
    return true;
  };
  view.startInlineRangeDeletion = () => { structuralDeletes += 1; };
  view.recordInlineEditSnapshot = () => {};
  view.syncInlineTextPreviewSafely = (_target, text, options) => {
    preview = { text, options };
    return true;
  };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};
  view.getElementBox = () => null;

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {},
    stopPropagation() {},
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(replacementStarts, 1);
  assert.equal(structuralDeletes, 0);
  assert.equal(editor.value, "");
  assert.deepEqual(preview, {
    text: "",
    options: { replaceTextFrame: true },
  });
  assert.equal(view.activeEditorTextDirty, true);
});

test("inline deletion retains a visible preview when incremental sync throws", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const element = {};
  const editor = {
    value: "M",
    selectionStart: 1,
    selectionEnd: 1,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element, shapeIndex: 42, paragraphIndex: 1 };
  let recovery = null;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.recordInlineEditSnapshot = () => {};
  view.syncShapeParagraphPreview = () => {
    throw new Error("test preview failure");
  };
  view.recoverInlineTextPreviewAfterSyncFailure = (_target, text) => {
    recovery = text;
    return true;
  };
  view.rememberInlineCaretPlacement = () => {};
  view.resetInlineEditorScroll = () => {};
  view.updateInlineCaret = () => {};
  view.preserveCanvasScrollAfterInlineTextEdit = () => {};

  const handled = view.handleInlineDeleteKey({
    key: "Backspace",
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    isComposing: false,
    preventDefault() {},
    stopPropagation() {},
  }, editor, element);

  assert.equal(handled, true);
  assert.equal(editor.value, "");
  assert.equal(recovery, "");
});

test("inline undo restores through preview recovery without losing its snapshot", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const editor = {
    value: "after delete",
    selectionStart: 0,
    selectionEnd: 0,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const target = { element: {}, shapeIndex: 42, paragraphIndex: 1 };
  let recovered = null;

  view.activeEditor = editor;
  view.activeShapeTextTarget = target;
  view.inlineUndoStack = [{ value: "before delete", selectionStart: 2, selectionEnd: 8 }];
  view.inlineRedoStack = [];
  view.syncShapeParagraphPreview = () => {
    throw new Error("test preview failure");
  };
  view.recoverInlineTextPreviewAfterSyncFailure = (_target, text) => {
    recovered = text;
    return true;
  };
  view.clearWholeShapeInlineSelection = () => {};
  view.historyController.updateAvailability = () => {};

  assert.equal(view.undoInlineEdit(editor), true);
  assert.equal(recovered, "before delete");
  assert.equal(editor.value, "before delete");
  assert.deepEqual(
    { start: editor.selectionStart, end: editor.selectionEnd },
    { start: 2, end: 8 },
  );
  assert.equal(view.inlineUndoStack.length, 0);
  assert.deepEqual(view.inlineRedoStack, [{ value: "after delete", selectionStart: 0, selectionEnd: 0 }]);
});

test("failed inline undo restoration keeps the history entry retryable", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const editor = {
    value: "after delete",
    selectionStart: 0,
    selectionEnd: 0,
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
  };
  const snapshot = { value: "before delete", selectionStart: 2, selectionEnd: 8 };

  view.activeEditor = editor;
  view.inlineUndoStack = [snapshot];
  view.inlineRedoStack = [];
  view.applyInlineSnapshot = () => {
    throw new Error("unrecoverable test failure");
  };
  view.historyController.updateAvailability = () => {};

  assert.equal(view.undoInlineEdit(editor), false);
  assert.deepEqual(view.inlineUndoStack, [snapshot]);
  assert.deepEqual(view.inlineRedoStack, []);

  let commits = 0;
  let documentUndoCalls = 0;
  view.finishInlineTextEditing = async () => { commits += 1; };
  view.session.undo = () => { documentUndoCalls += 1; return true; };

  await view.requestHistoryAction("undo", "toolbar");
  assert.equal(commits, 0);
  assert.equal(documentUndoCalls, 0);
});

test("Edit menu history follows the active inline editor before document history", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const editor = { value: "after delete" };
  let request = null;
  let documentUndoCalls = 0;

  view.activeEditor = editor;
  view.inlineUndoStack = [{ value: "before delete", selectionStart: 0, selectionEnd: 0 }];
  view.canEdit = () => true;
  view.session.undo = () => { documentUndoCalls += 1; };
  view.requestHistoryAction = (action, source) => {
    request = { action, source };
    return Promise.resolve();
  };

  const [undo] = view.getEditMenuItems();
  assert.equal(undo.disabled, false);
  undo.onClick();

  assert.deepEqual(request, { action: "undo", source: "menu" });
  assert.equal(documentUndoCalls, 0);
});

test("inline commit keeps the last transaction text if the hidden textarea regresses", async () => {
  const { NativePowerPointView } = await loadNativePowerPointViewModule();
  const view = new NativePowerPointView({ app: { vault: {} } }, () => ({
    autosaveEnabled: false,
    yoloMode: false,
  }));
  const editor = { value: "original text" };

  view.activeEditor = editor;
  view.activeEditorCanonicalText = "deleted text";

  assert.equal(view.resolveInlineTextForCommit(editor), "deleted text");
});
