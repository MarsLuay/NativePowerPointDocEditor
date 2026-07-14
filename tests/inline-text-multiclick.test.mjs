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

  view.stopInlineSelectionDrag = () => {};
  view.clearWholeShapeInlineSelection = () => {};
  view.updateInlineCaret = () => {};
  view.selectAllInlineText = () => {
    wholeTextBoxSelections += 1;
  };

  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 2), true);
  assert.deepEqual(
    { start: editor.selectionStart, end: editor.selectionEnd },
    { start: 6, end: 11 },
  );

  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 3), true);
  assert.equal(wholeTextBoxSelections, 1);
  assert.equal(view.applyInlineMultiClickSelection(editor, element, 8, 1), false);
});
