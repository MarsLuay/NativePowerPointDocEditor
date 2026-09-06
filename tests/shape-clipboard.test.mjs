import assert from "node:assert/strict";
import { test } from "node:test";
import { loadShapeClipboardModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

test("createSlideObjectsClipboard should create a clipboard object with normalized shape indexes", async () => {
  const { createSlideObjectsClipboard } = await loadShapeClipboardModule();

  const buffer = new ArrayBuffer(8);
  const slideIndex = 0;
  const shapeIndexes = [2, 0, 1, 2];

  const clipboard = createSlideObjectsClipboard(buffer, slideIndex, shapeIndexes);

  assert.equal(clipboard.slideIndex, 0);
  assert.deepEqual(clipboard.shapeIndexes, [0, 1, 2]);
  assert.equal(clipboard.shapeIndex, 0);
});

test("createSlideObjectClipboard should create a clipboard object for a single shape", async () => {
  const { createSlideObjectClipboard } = await loadShapeClipboardModule();

  const buffer = new ArrayBuffer(8);
  const slideIndex = 1;
  const shapeIndex = 3;

  const clipboard = createSlideObjectClipboard(buffer, slideIndex, shapeIndex);

  assert.equal(clipboard.slideIndex, 1);
  assert.deepEqual(clipboard.shapeIndexes, [3]);
  assert.equal(clipboard.shapeIndex, 3);
});

test("createSlideObjectsClipboard throws error if shape indexes are empty", async () => {
  const { createSlideObjectsClipboard } = await loadShapeClipboardModule();

  const buffer = new ArrayBuffer(8);
  assert.throws(
    () => createSlideObjectsClipboard(buffer, 0, []),
    /Select at least one PowerPoint object to copy/
  );
  assert.throws(
    () => createSlideObjectsClipboard(buffer, 0, [-1]),
    /Select at least one PowerPoint object to copy/
  );
});

test("pasteSlideObject pastes a single shape", async () => {
  const { pasteSlideObject, createSlideObjectClipboard } = await loadShapeClipboardModule();

  const buffer = toArrayBuffer(await readDeck("simple-edit.pptx"));
  const clipboard = createSlideObjectClipboard(buffer, 0, 0);

  const result = await pasteSlideObject(buffer, clipboard, 0);

  assert.ok(result.buffer instanceof ArrayBuffer);
  assert.ok(Number.isInteger(result.shapeIndex));
});

test("pasteSlideObjects pastes multiple shapes", async () => {
  const { pasteSlideObjects, createSlideObjectsClipboard } = await loadShapeClipboardModule();

  const buffer = toArrayBuffer(await readDeck("features.pptx"));
  const clipboard = createSlideObjectsClipboard(buffer, 0, [0, 2]);

  const result = await pasteSlideObjects(buffer, clipboard, 0);

  assert.ok(result.buffer instanceof ArrayBuffer);
  assert.equal(result.shapeIndexes.length, 2);
  assert.ok(result.shapeIndexes.every(idx => Number.isInteger(idx)));
});
