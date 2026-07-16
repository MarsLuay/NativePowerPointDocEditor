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
