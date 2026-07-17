import assert from "node:assert/strict";
import { test } from "node:test";
import { extractZip } from "pptx-svg";
import {
  loadPresentationEngineModule,
  loadShapeClipboardModule,
} from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

// Counts top-level slide objects by the per-shape <g data-ooxml-shape-idx>
// groups the renderer emits, so a successful paste shows up as exactly one more
// group than the source deck.
function countRenderedShapes(svg) {
  return (svg.match(/data-ooxml-shape-idx=/g) ?? []).length;
}

function sameBytes(left, right) {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((byte, index) => byte === right[index]);
}

// PART 2(a) — ShapeClipboard is pure-ish (createSlideObjectClipboard copies the
// source buffer; pasteSlideObject reads source + destination and returns a fresh
// buffer). This exercises the copy/paste/duplicate path the view wires through
// PresentationEngine.copyShape/pasteShape/duplicateShape.
test("pasting one clipboard twice yields independent, round-trippable decks", async () => {
  const { createSlideObjectClipboard, pasteSlideObject } = await loadShapeClipboardModule();
  const { PresentationEngine } = await loadPresentationEngineModule();

  const source = toArrayBuffer(await readDeck("simple-edit.pptx"));
  const clipboard = createSlideObjectClipboard(source, 0, 0);

  // createSlideObjectClipboard must snapshot the source (buffer.slice(0)), not
  // alias it, or later edits to the open deck would corrupt the clipboard.
  assert.notEqual(clipboard.buffer, source, "clipboard must not alias the source ArrayBuffer");
  assert.equal(clipboard.buffer.byteLength, source.byteLength);
  const clipboardSnapshot = new Uint8Array(clipboard.buffer.slice(0));

  // Destroy the source backing store. An aliasing clipboard would now be a
  // zeroed (invalid) ZIP; an independent copy stays a valid "PK" archive.
  new Uint8Array(source).fill(0);
  const clipboardBytes = new Uint8Array(clipboard.buffer);
  assert.equal(clipboardBytes[0], 0x50, "clipboard buffer should still start with the ZIP 'PK' signature");
  assert.equal(clipboardBytes[1], 0x4b, "clipboard buffer should still start with the ZIP 'PK' signature");

  // Paste the SAME clipboard onto the SAME (independently re-read) destination
  // twice. If any state leaked between pastes (shared cache/workbook or a
  // mutated destination), the second paste would observe an extra shape and
  // report a different index.
  const destination = toArrayBuffer(await readDeck("simple-edit.pptx"));
  const destinationSnapshot = new Uint8Array(destination.slice(0));

  const firstPaste = await pasteSlideObject(destination, clipboard, 0);
  const secondPaste = await pasteSlideObject(destination, clipboard, 0);

  assert.ok(
    sameBytes(new Uint8Array(destination), destinationSnapshot),
    "pasteSlideObject must not mutate the destination buffer it was handed",
  );
  assert.ok(
    sameBytes(new Uint8Array(clipboard.buffer), clipboardSnapshot),
    "pasteSlideObject must not mutate the clipboard buffer",
  );
  assert.notEqual(firstPaste.buffer, secondPaste.buffer, "each paste must produce a distinct output buffer");
  assert.equal(
    firstPaste.shapeIndex,
    secondPaste.shapeIndex,
    "two pastes of the same clipboard onto the same deck must append at the same index (no accumulation)",
  );

  // Deck round-trips: each pasted buffer reloads, keeps its slide count, renders,
  // and carries exactly one more shape than the source deck.
  const sourceEngine = await PresentationEngine.load(toArrayBuffer(await readDeck("simple-edit.pptx")));
  const baselineShapeCount = countRenderedShapes(sourceEngine.renderSlide(0).svg);

  for (const result of [firstPaste, secondPaste]) {
    const engine = await PresentationEngine.load(result.buffer);
    assert.equal(engine.slideCount, sourceEngine.slideCount, "paste must not change the slide count");
    const rendered = engine.renderSlide(0);
    assert.match(rendered.svg, /^<svg\b/, "pasted deck must still render to SVG");
    assert.equal(
      countRenderedShapes(rendered.svg),
      baselineShapeCount + 1,
      "paste must add exactly one shape to the slide",
    );
    const reexported = await engine.export();
    assert.ok(reexported.byteLength > 0, "pasted deck must re-export without error");
  }

  // Independence at the part level: a second paste of the same clipboard must
  // not have grafted the first paste's relationship/media parts. Both outputs
  // should hold the identical set of part paths.
  const firstParts = [...(await extractZip(firstPaste.buffer)).textFiles.keys()].sort();
  const secondParts = [...(await extractZip(secondPaste.buffer)).textFiles.keys()].sort();
  assert.deepEqual(secondParts, firstParts, "independent pastes must produce identical part layouts");
});

test("a multi-object clipboard pastes every selected object in one round trip", async () => {
  const { createSlideObjectsClipboard, pasteSlideObjects } = await loadShapeClipboardModule();
  const { PresentationEngine } = await loadPresentationEngineModule();
  const source = toArrayBuffer(await readDeck("features.pptx"));
  const sourceEngine = await PresentationEngine.load(source.slice(0));
  const baselineShapeCount = countRenderedShapes(sourceEngine.renderSlide(0).svg);
  assert.ok(baselineShapeCount >= 2, "fixture must provide two independently selectable objects");

  const clipboard = createSlideObjectsClipboard(source, 0, [2, 0, 2]);
  assert.deepEqual(
    clipboard.shapeIndexes,
    [0, 2],
    "clipboard must deduplicate selection and preserve source z-order",
  );

  const pasted = await pasteSlideObjects(source, clipboard, 0);
  assert.equal(pasted.shapeIndexes.length, 2, "paste must return one renderer index per copied object");
  assert.deepEqual(
    pasted.shapeIndexes,
    [...pasted.shapeIndexes].sort((left, right) => left - right),
    "pasted objects must retain their source stacking order",
  );

  const pastedEngine = await PresentationEngine.load(pasted.buffer);
  assert.equal(
    countRenderedShapes(pastedEngine.renderSlide(0).svg),
    baselineShapeCount + 2,
    "pasting a two-object clipboard must add both objects",
  );
  assert.ok((await pastedEngine.export()).byteLength > 0, "multi-object paste must remain exportable");

  const slideXml = (await extractZip(pasted.buffer)).textFiles.get("ppt/slides/slide1.xml") ?? "";
  const nonVisualIds = [...slideXml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => match[1]);
  assert.equal(
    new Set(nonVisualIds).size,
    nonVisualIds.length,
    "every pasted object must receive fresh non-visual ids",
  );
});

// PART 2(b) — The FindReplaceController's search (collectFindMatches), match
// navigation (moveFindMatch), and highlighting (applyFindHighlight) are all
// `private` and bound to a live DOM: they read `this.findInputEl.value`, parse
// `engine.renderSlide(...).svg` with `new DOMParser()`, run
// `querySelectorAll('g[data-ooxml-shape-idx]')`, and mutate `host.svgEl`. The
// only entry points (`createPanel`, `open`) call Obsidian DOM helpers
// (`activeDocument.body.createDiv`, `createEl`, `requestAnimationFrame`, focus),
// so driving find headlessly would require a full Obsidian DOM. Find is
// therefore host/DOM-bound and verified manually.
//
// What IS unit-testable is the replacement engine the controller delegates to:
// replaceCurrentMatch/replaceAllMatches both call PresentationEngine.replaceText
// (scoped vs. deck-wide), so its match counting and mutation are the substance
// of "replace / replace all" correctness.
test("replaceText counts and mutates deck-wide and scoped, as find/replace delegates", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(toArrayBuffer(await readDeck("simple-edit.pptx")));

  await engine.updateParagraphText(0, 0, 0, "alpha beta alpha gamma alpha");

  // Empty query is a no-op (guard in replaceText).
  assert.equal(await engine.replaceText("", "x"), 0, "empty query must replace nothing");

  // Deck-wide (Replace all) reports the true occurrence count and mutates.
  const replacedAll = await engine.replaceText("alpha", "omega");
  assert.equal(replacedAll, 3, "Replace all must report every occurrence");
  const afterAll = engine.renderSlide(0).svg;
  assert.ok(afterAll.includes("omega"), "replacement text must appear after Replace all");
  assert.ok(!afterAll.includes("alpha"), "original text must be gone after Replace all");

  // Missing term reports zero (the controller surfaces 'No matches to replace').
  assert.equal(await engine.replaceText("does-not-exist-xyz", "x"), 0, "absent query must report zero matches");

  // Scoped (single-match Replace path) restricts to one shape and still counts.
  const replacedScoped = await engine.replaceText("omega", "delta", { slideIndex: 0, shapeIndex: 0 });
  assert.equal(replacedScoped, 3, "scoped replace must report matches within the targeted shape");
  const afterScoped = engine.renderSlide(0).svg;
  assert.ok(afterScoped.includes("delta"), "scoped replacement text must appear");
  assert.ok(!afterScoped.includes("omega"), "scoped replace must consume the prior text");

  // Case-insensitive by default (matches collectFindMatches' lower-cased compare).
  await engine.updateParagraphText(0, 0, 0, "Title TITLE title");
  assert.equal(await engine.replaceText("title", "word"), 3, "default replace must be case-insensitive");
});

// PART 2(c) — collectFindMatches searches over the shape's CONCATENATED text, so
// a query that visually spans two runs (e.g. "alpha" + "beta" rendered as
// "alphabeta") is FOUND. Before the fix, replaceText substituted only within
// each individual <a:t> run, so such a query returned 0 ("No matches to
// replace") — a confusing find/replace mismatch. This proves the spanning match
// is now both counted and replaced while untouched runs keep their formatting.
test("replaceText replaces a match spanning two runs in one paragraph", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(toArrayBuffer(await readDeck("simple-edit.pptx")));

  // Build a single paragraph whose text is "alphabeta", then split it into two
  // adjacent runs by styling only the first five characters bold. The result is
  // run0="alpha" (bold) + run1="beta": the word "alphabeta" lives only in the
  // concatenation of the two runs, not in any single <a:t>.
  await engine.updateParagraphText(0, 0, 0, "alphabeta");
  await engine.setRunStyleForRange(0, 0, 0, 0, 5, { bold: true });

  // Confirm the paragraph really has two runs and that the split applied the
  // formatting we rely on (so this exercises the multi-run path, not a fast
  // single-run replace).
  const firstRun = engine.getRunStyle(0, 0, 0, 0);
  const secondRun = engine.getRunStyle(0, 0, 0, 1);
  assert.ok(firstRun?.bold, "first run must be the bold 'alpha' run after the split");
  assert.ok(secondRun, "paragraph must contain a second run ('beta') after the split");
  assert.equal(engine.getRunStyle(0, 0, 0, 2), null, "paragraph must contain exactly two runs");

  // The cross-run query is replaced once and the substitution is visible.
  const replaced = await engine.replaceText("alphabeta", "wrapped");
  assert.equal(replaced, 1, "a match spanning two runs must be counted exactly once");
  const svg = engine.renderSlide(0).svg;
  assert.ok(svg.includes("wrapped"), "the cross-run replacement text must appear");
  assert.ok(!svg.includes("alphabeta"), "the original cross-run text must be gone");

  // The replacement is anchored in the first run, which keeps its bold styling;
  // the now-empty second run is preserved without corrupting the paragraph.
  const anchorRun = engine.getRunStyle(0, 0, 0, 0);
  assert.ok(anchorRun?.bold, "the run anchoring the replacement must keep its formatting");

  // A partial match that starts in run0 and ends in run1 is also handled,
  // leaving the unmatched characters of each run intact.
  await engine.updateParagraphText(0, 0, 0, "alphabeta");
  await engine.setRunStyleForRange(0, 0, 0, 0, 5, { bold: true });
  const partial = await engine.replaceText("phabe", "-");
  assert.equal(partial, 1, "a partial match crossing the run boundary must be counted once");
  // Strip tags so the assertion does not depend on whether the renderer emits the
  // two surviving runs ("al-" and "ta") as separate tspans.
  const partialText = engine.renderSlide(0).svg.replace(/<[^>]+>/g, "");
  assert.ok(partialText.includes("al-ta"), "unmatched characters on both sides of the boundary must survive");
});
