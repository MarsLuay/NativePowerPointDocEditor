import assert from "node:assert/strict";
import test from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";
import { loadPresentationEngineModule, loadTextUtilsModule } from "./helpers/load-plugin-modules.mjs";

const { mapEditorOffsetToOoxmlOffset, alignRunTspansToOoxml, mapEditorRangeToOoxml } = await loadTextUtilsModule();

test("mapEditorRangeToOoxml reproduces start/end semantics from aligned tiles", () => {
  const ooxml = "alpha beta gamma";
  // Three wrapped lines, two dropped spaces.
  const tiles = alignRunTspansToOoxml(["alpha", "beta", "gamma"], ooxml).spans;

  // Full selection covers the whole paragraph.
  assert.deepEqual(mapEditorRangeToOoxml(tiles, 0, "alphabetagamma".length), { start: 0, end: ooxml.length });

  // Selecting editor "beta" (offsets 5..9) maps to exactly OOXML "beta": START
  // skips the dropped space before it, END stops at the next dropped space.
  const r = mapEditorRangeToOoxml(tiles, 5, 9);
  assert.equal(r.start, ooxml.indexOf("beta"));
  assert.equal(ooxml.slice(r.start, r.end).replace(/\s+$/, ""), "beta");

  // Cross-boundary selection of editor "phabe" still lands exactly.
  const cross = mapEditorRangeToOoxml(tiles, 2, 9);
  assert.equal(ooxml.slice(cross.start, cross.end).replace(/\s+$/, ""), "pha beta");
});

test("mapEditorRangeToOoxml matches the per-call mapper, tiling dropped spaces to one line", () => {
  const ooxml = "Connect, and manage all your Samsung and SmartThings appliances";
  const lines = ["Connect, and manage all your Samsung and", "SmartThings", "appliances"];
  const editor = lines.join("");
  const tiles = alignRunTspansToOoxml(lines, ooxml).spans;
  // Wrap-boundary editor offsets: the tile mapper intentionally tiles the dropped
  // space onto the preceding line's END, so a START there skips it (lands one past
  // the old false-semantics result). Everywhere else the two mappers must agree.
  const boundaries = new Set();
  let cursor = 0;
  for (const line of lines) {
    cursor += line.length;
    boundaries.add(cursor);
  }
  for (let i = 0; i <= editor.length; i++) {
    for (let j = i; j <= editor.length; j++) {
      const viaTiles = mapEditorRangeToOoxml(tiles, i, j);
      const oldStart = mapEditorOffsetToOoxmlOffset(editor, ooxml, i, false);
      const end = mapEditorOffsetToOoxmlOffset(editor, ooxml, j, true);
      assert.equal(viaTiles.end, end, `end ${i}..${j}`);
      if (boundaries.has(i) && i !== editor.length) {
        // Tile START skips the dropped space; old mapper stopped on it.
        assert.ok(viaTiles.start >= oldStart, `boundary start ${i}`);
        assert.ok(viaTiles.start - oldStart <= 1, `boundary start drift ${i}`);
      } else {
        assert.equal(viaTiles.start, oldStart, `start ${i}..${j}`);
      }
    }
  }
});

test("alignRunTspansToOoxml: identity when no whitespace is dropped", () => {
  const ooxml = "Hello world";
  const result = alignRunTspansToOoxml(["Hello ", "world"], ooxml);
  assert.equal(result.reconciled, true);
  assert.deepEqual(result.spans.map((s) => [s.charStart, s.charEnd]), [[0, 6], [6, 11]]);
});

test("alignRunTspansToOoxml: dropped wrap space is absorbed into the preceding line's end", () => {
  // OOXML keeps the space at index 5; the two wrapped run tspans dropped it.
  const ooxml = "alpha beta";
  const result = alignRunTspansToOoxml(["alpha", "beta"], ooxml);
  assert.equal(result.reconciled, true);
  // Line 0 charEnd absorbs the trailing space (6), line 1 starts after it.
  assert.deepEqual(result.spans.map((s) => [s.charStart, s.charEnd]), [[0, 6], [6, 10]]);
  // The whole paragraph is tiled with no gap and the dropped space is covered.
  assert.equal(result.spans[result.spans.length - 1].charEnd, ooxml.length);
});

test("alignRunTspansToOoxml: multiple wrapped lines tile the full OOXML text", () => {
  const ooxml = "Connect, automate, and manage all your Samsung and SmartThings-compatible appliances";
  // Simulate three wrapped lines (spaces dropped at the two wrap boundaries).
  const lines = ["Connect, automate, and manage all your Samsung and", "SmartThings-compatible", "appliances"];
  const result = alignRunTspansToOoxml(lines, ooxml);
  assert.equal(result.reconciled, true);
  assert.equal(result.spans[0].charStart, 0);
  assert.equal(result.spans[result.spans.length - 1].charEnd, ooxml.length);
  // Each line's OOXML slice (trimmed of the absorbed trailing space) is the line text.
  for (let i = 0; i < lines.length; i++) {
    assert.equal(ooxml.slice(result.spans[i].charStart, result.spans[i].charEnd).replace(/\s+$/, ""), lines[i]);
  }
});

test("alignRunTspansToOoxml: empty paragraph reconciles", () => {
  assert.equal(alignRunTspansToOoxml([], "").reconciled, true);
});

// The canonical soft-wrapping paragraph the user kept hitting: a bullet line
// plus a long run that wraps across several visual lines. OOXML keeps the
// spaces swallowed at wrap boundaries; the SVG-derived editor text drops them.
const fixtureTitleParagraph =
  '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>';

const bulletParagraph =
  '<a:p>' +
  '<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="\u25CF"/></a:pPr>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>Because sm</a:t></a:r>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>art home automate, and manage all your Samsung and SmartThings-compatible appliances and electronics with a single, easy-to-use app.</a:t></a:r>' +
  '<a:endParaRPr lang="en-US"/>' +
  '</a:p>';

async function loadEngineWithBulletParagraph() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  assert.ok(slideXml);
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, bulletParagraph)]])
  );
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(patched);
}

// Reconstruct the inline editor's paragraph text exactly the way the view does:
// concatenate the textContent of every run-bearing tspan (data-ooxml-run-idx)
// under the paragraph, in document order. This is the SVG-space string that the
// editor selection offsets are measured against.
function editorTextFromSvg(svg, shapeIndex, paragraphIndex) {
  const doc = new globalThis.DOMParser().parseFromString(svg, "text/xml");
  // Run text lives in tspans carrying data-ooxml-run-idx, nested under a
  // line-container tspan (data-ooxml-para-idx) and a shape group
  // (data-ooxml-shape-idx). The view selects them with a descendant selector.
  // Concatenate the target shape+paragraph's run text in document order --
  // exactly the inline editor's paragraph string.
  const ancestorAttr = (el, attr) => {
    for (let node = el.parentNode; node && node.getAttribute; node = node.parentNode) {
      const value = node.getAttribute(attr);
      if (value !== null) return value;
    }
    return null;
  };
  return Array.from(doc.getElementsByTagName("tspan"))
    .filter(
      (el) =>
        el.getAttribute("data-ooxml-run-idx") !== null &&
        ancestorAttr(el, "data-ooxml-para-idx") === String(paragraphIndex) &&
        ancestorAttr(el, "data-ooxml-shape-idx") === String(shapeIndex)
    )
    .map((el) => el.textContent || "")
    .join("");
}

function paragraphHighlightSpan(engine, slide, shapeIndex, paragraphIndex) {
  const hl = engine
    .getSlideRunHighlights(slide)
    .filter((h) => h.shapeIndex === shapeIndex && h.paragraphIndex === paragraphIndex && h.end > h.start);
  if (!hl.length) return null;
  return { start: Math.min(...hl.map((h) => h.start)), end: Math.max(...hl.map((h) => h.end)), runs: hl.length };
}

test("trailing-gap mapping absorbs dropped wrap whitespace but never real characters", () => {
  // Normal soft-wrap case: OOXML keeps a space the editor dropped at the END.
  const ooxml = "Samsung and SmartThings";
  const editor = "Samsung andSmartThings"; // the wrap space was swallowed
  const editorEnd = editor.indexOf("SmartThings"); // caret right before the wrapped word
  // END consumes the dropped space so the selection spans the wrap boundary.
  assert.equal(mapEditorOffsetToOoxmlOffset(editor, ooxml, editorEnd, true), ooxml.indexOf("SmartThings"));

  // Divergence guard: if the editor string is NOT a clean subsequence (a real
  // letter differs), the END must not gobble the rest of the paragraph.
  const ooxml2 = "alpha beta gamma";
  const editor2 = "alpha Xeta gamma"; // 'b' -> 'X' (non-whitespace divergence)
  const end = mapEditorOffsetToOoxmlOffset(editor2, ooxml2, "alpha ".length, true);
  assert.ok(end <= "alpha ".length + 1, "END stops at the divergence instead of consuming real text");
});

test("soft-break newlines are editor-only glyphs: split offsets skip them in OOXML run text", () => {
  // Enter-splits-at-wrong-character repro. `getParagraphRunText` concatenates
  // <a:r> run text only; a soft break (<a:br/>) contributes ZERO characters, so
  // the OOXML string has no `\n` while the editor value does. Each soft break
  // before the caret must be consumed from the editor side without advancing
  // OOXML, otherwise the split target overshoots by one glyph per break.
  const editor = "AB\nCD"; // one soft break after "B"
  const ooxml = "ABCD"; // run text: the break is not a character
  // Caret after "C" (editor offset 4) -> OOXML offset 3 (split keeps "ABC").
  assert.equal(mapEditorOffsetToOoxmlOffset(editor, ooxml, 4, false), 3);
  // Caret right after the break (editor offset 3, before "C") -> OOXML offset 2.
  assert.equal(mapEditorOffsetToOoxmlOffset(editor, ooxml, 3, false), 2);
  // Caret before the break (offset 2) is unaffected.
  assert.equal(mapEditorOffsetToOoxmlOffset(editor, ooxml, 2, false), 2);

  // Multiple soft breaks accumulate: two breaks before the caret => offset - 2.
  const multi = "A\nB\nCD";
  assert.equal(mapEditorOffsetToOoxmlOffset(multi, "ABCD", "A\nB\nC".length, false), 3);

  // Pending-edit split maps editor-space -> editor-space text: BOTH strings
  // carry the break, so the equality branch matches and the newline is NOT
  // skipped one-sidedly (no regression to the pending path).
  const editorSpace = "AB\nCD";
  assert.equal(mapEditorOffsetToOoxmlOffset(editorSpace, editorSpace, 4, false), 4);
});

test("end-to-end: clearing the full editor selection removes every highlight (no residual)", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const shapeIndex = 0;
  const paragraphIndex = 0;

  const ooxmlText = engine.getParagraphRunText(slide, shapeIndex, paragraphIndex);
  const editorText = editorTextFromSvg(engine.renderSlide(slide).svg, shapeIndex, paragraphIndex);
  assert.ok(ooxmlText && ooxmlText.length > 0);
  assert.ok(editorText.length > 0);
  // The SVG editor text drops the wrap-boundary space(s); it is a strict
  // subsequence shorter than the OOXML run text.
  assert.ok(editorText.length <= ooxmlText.length, "editor text never exceeds OOXML text");

  // Highlight the whole paragraph (the deck-author's pre-highlight equivalent).
  await engine.setRunStyleForRanges(slide, shapeIndex,
    [{ paragraphIndex, start: 0, end: ooxmlText.length }], { highlight: "FFFF00" });
  assert.ok(paragraphHighlightSpan(engine, slide, shapeIndex, paragraphIndex), "highlight applied");

  // The user drags across the entire visible text and hits "No color". The view
  // maps the editor-space END through the wrap gap to the true OOXML end.
  const mappedStart = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, 0, false);
  const mappedEnd = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, editorText.length, true);
  assert.equal(mappedStart, 0);
  assert.equal(mappedEnd, ooxmlText.length, "full editor selection maps to the true OOXML paragraph end");

  await engine.setRunStyleForRanges(slide, shapeIndex,
    [{ paragraphIndex, start: mappedStart, end: mappedEnd }], { highlight: null });

  assert.equal(
    paragraphHighlightSpan(engine, slide, shapeIndex, paragraphIndex),
    null,
    "no residual highlight after clearing the full selection"
  );
});

test("end-to-end: clearing a wrapped sub-word clears exactly that word, not its neighbors", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const shapeIndex = 0;
  const paragraphIndex = 0;

  const ooxmlText = engine.getParagraphRunText(slide, shapeIndex, paragraphIndex);
  const editorText = editorTextFromSvg(engine.renderSlide(slide).svg, shapeIndex, paragraphIndex);

  await engine.setRunStyleForRanges(slide, shapeIndex,
    [{ paragraphIndex, start: 0, end: ooxmlText.length }], { highlight: "FFFF00" });

  // Select "easy-to-use" in editor space and clear only it.
  const word = "easy-to-use";
  const eStart = editorText.indexOf(word);
  assert.ok(eStart >= 0, "word present in editor text");
  const mappedStart = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, eStart, false);
  const mappedEnd = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, eStart + word.length, true);
  assert.equal(ooxmlText.slice(mappedStart, mappedEnd).replace(/\s+$/, ""), word);

  await engine.setRunStyleForRanges(slide, shapeIndex,
    [{ paragraphIndex, start: mappedStart, end: mappedEnd }], { highlight: null });

  const span = paragraphHighlightSpan(engine, slide, shapeIndex, paragraphIndex);
  assert.ok(span, "the rest of the paragraph stays highlighted");
  // The cleared hole is exactly [mappedStart, mappedEnd): text before and after stays styled.
  assert.equal(engine.isRangeStyled(slide, shapeIndex, paragraphIndex, 0, mappedStart, "underline"), false);
  const beforeStillHighlighted = engine
    .getSlideRunHighlights(slide)
    .some((h) => h.paragraphIndex === paragraphIndex && h.start < mappedStart);
  const afterStillHighlighted = engine
    .getSlideRunHighlights(slide)
    .some((h) => h.paragraphIndex === paragraphIndex && h.end > mappedEnd);
  assert.ok(beforeStillHighlighted, "text before the word stays highlighted");
  assert.ok(afterStillHighlighted, "text after the word stays highlighted");
  // Nothing inside the cleared word remains highlighted.
  const insideCleared = engine
    .getSlideRunHighlights(slide)
    .some((h) => h.paragraphIndex === paragraphIndex && h.start < mappedEnd && h.end > mappedStart && h.start >= mappedStart && h.end <= mappedEnd);
  assert.ok(!insideCleared, "the selected word is fully cleared");
});

// Inspect the live OOXML run list for the paragraph whose concatenated run text
// contains `matchSubstring`. Returns the run count and the joined text.
function paragraphRuns(engine, slide, matchSubstring) {
  const ooxml = engine.renderer.getSlideOoxml(slide);
  const doc = new globalThis.DOMParser().parseFromString(ooxml, "application/xml");
  const para = Array.from(doc.getElementsByTagName("a:p")).find((p) =>
    Array.from(p.getElementsByTagName("a:r"))
      .map((r) => r.getElementsByTagName("a:t")[0]?.textContent || "")
      .join("")
      .includes(matchSubstring)
  );
  if (!para) return null;
  const runs = Array.from(para.getElementsByTagName("a:r"));
  return {
    count: runs.length,
    text: runs.map((r) => r.getElementsByTagName("a:t")[0]?.textContent || "").join(""),
  };
}

test("highlight fidelity: moving a highlighted shape preserves <a:highlight> (no renderer drop)", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;

  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: 20 }], { highlight: "FFFF00" });
  const before = engine.getSlideRunHighlights(slide).filter((h) => h.end > h.start);
  assert.ok(before.length > 0, "highlight applied");

  // A Wasm-primitive edit (move/resize) re-serializes the slide; without the
  // preservation wrapper the renderer's model drops every <a:highlight>.
  await engine.updateShapeTransform(slide, 0, { x: 120000, y: 120000, cx: 5000000, cy: 1200000, rot: 0 });

  const after = engine.getSlideRunHighlights(slide).filter((h) => h.end > h.start);
  assert.equal(after.length, before.length, "highlight survives the shape transform");
  assert.deepEqual(
    after.map((h) => [h.paragraphIndex, h.start, h.end, h.color]),
    before.map((h) => [h.paragraphIndex, h.start, h.end, h.color]),
    "same runs/color stay highlighted after the move"
  );

  // And it round-trips through export/reopen.
  const exported = await engine.export();
  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(exported);
  assert.equal(
    reopened.getSlideRunHighlights(slide).filter((h) => h.end > h.start).length,
    before.length,
    "highlight persists after export + reopen following a transform"
  );
});

// A run wearing properties the pptx-svg model drops on re-serialize: a run
// hyperlink (hlinkClick), an underline fill (uFill), normalize-heights, plus a
// highlight (handled by the offset path) as a control.
const decoratedRunParagraph =
  "<a:p><a:r>" +
  '<a:rPr lang="en-US" sz="1800" normalizeH="1">' +
  '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>' +
  '<a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill>' +
  '<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/>' +
  "</a:rPr><a:t>Linked decorated run</a:t></a:r>" +
  '<a:endParaRPr lang="en-US"/></a:p>';

async function loadEngineWithDecoratedRun() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  assert.ok(slideXml);
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, decoratedRunParagraph)]])
  );
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(patched);
}

function runRprFlags(ooxml, runText) {
  const doc = new globalThis.DOMParser().parseFromString(ooxml, "application/xml");
  const run = Array.from(doc.getElementsByTagName("a:r")).find(
    (r) => (r.getElementsByTagName("a:t")[0]?.textContent || "") === runText
  );
  const rPr = run?.getElementsByTagName("a:rPr")[0];
  if (!rPr) return null;
  const hasChild = (ln) => Array.from(rPr.childNodes).some((n) => n.nodeType === 1 && n.localName === ln);
  return {
    normalizeH: rPr.getAttribute("normalizeH"),
    hlinkClick: hasChild("hlinkClick"),
    uFill: hasChild("uFill"),
    highlight: hasChild("highlight"),
  };
}

async function reopen(engine) {
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(await engine.export());
}

function decoratedShapeIndex(engine, slide, text) {
  for (let i = 0; i < 32; i++) {
    if (engine.getParagraphRunText(slide, i, 0)?.includes(text)) return i;
  }
  return -1;
}

const sig = (h) => [h.shapeIndex, h.paragraphIndex, h.start, h.end, h.color];
const liveHighlights = (engine, slide) =>
  engine.getSlideRunHighlights(slide).filter((h) => h.end > h.start);

// The renderer's addShape return value is its own index space; styling APIs use
// document order (getShapeElement). Resolve a shape by its first paragraph text.
function shapeIndexByParagraphText(engine, slide, text) {
  for (let i = 0; i < 32; i++) {
    if (engine.getParagraphRunText(slide, i, 0) === text) return i;
  }
  return -1;
}

// The union of highlighted character offsets per paragraph, so an assertion can
// ignore how many runs the coverage is split across (e.g. a bold edit splits a
// highlighted run without changing what is highlighted).
function highlightCoverage(engine, slide) {
  const cover = new Map();
  for (const h of liveHighlights(engine, slide)) {
    const key = `${h.shapeIndex}:${h.paragraphIndex}:${h.color}`;
    const set = cover.get(key) ?? new Set();
    for (let i = h.start; i < h.end; i++) set.add(i);
    cover.set(key, set);
  }
  return [...cover.entries()]
    .map(([key, set]) => `${key}@${[...set].sort((a, b) => a - b).join(",")}`)
    .sort();
}

test("highlight fidelity: adding a text box keeps the existing shape's <a:highlight>", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;

  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: 20 }], { highlight: "FFFF00" });
  const before = liveHighlights(engine, slide);
  assert.ok(before.length > 0, "highlight applied");

  // Adding a shape re-serializes the slide via a Wasm primitive, which drops
  // every <a:highlight> from the renderer model.
  const newIndex = await engine.addTextBox(slide);
  assert.ok(newIndex > 0, "the text box was appended after the existing shape");

  assert.deepEqual(
    liveHighlights(engine, slide).map(sig),
    before.map(sig),
    "the existing highlight is unchanged after adding a text box"
  );

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(await engine.export());
  assert.equal(
    liveHighlights(reopened, slide).length,
    before.length,
    "highlight persists after export + reopen following an add"
  );
});

test("highlight fidelity: deleting a lower shape remaps the surviving highlight index", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;

  // Highlight the bullet shape (index 0), then add a second text box and
  // highlight it too with a distinct color.
  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: 20 }], { highlight: "FFFF00" });
  await engine.addTextBox(slide);
  const boxIndex = shapeIndexByParagraphText(engine, slide, "New text");
  assert.ok(boxIndex > 0, "the new text box was found in document order");
  await engine.setRunStyleForRanges(slide, boxIndex, [{ paragraphIndex: 0, start: 0, end: 3 }], { highlight: "00FF00" });

  const before = liveHighlights(engine, slide);
  const boxHighlight = before.find((h) => h.shapeIndex === boxIndex && h.color === "00FF00");
  assert.ok(boxHighlight, "the new text box is highlighted");
  assert.ok(before.some((h) => h.shapeIndex === 0 && h.color === "FFFF00"), "the bullet shape is highlighted");

  // Delete the bullet (index 0); the renderer renumbers every higher shape down
  // by one, so the cached highlight must follow the text box to boxIndex-1.
  await engine.deleteShape(slide, 0);

  const after = liveHighlights(engine, slide);
  assert.ok(!after.some((h) => h.color === "FFFF00"), "the deleted shape's highlight is gone");
  const moved = after.find((h) => h.color === "00FF00");
  assert.ok(moved, "the surviving highlight is preserved");
  const textBoxIndex = shapeIndexByParagraphText(engine, slide, "New text");
  assert.equal(moved.shapeIndex, textBoxIndex, "the surviving highlight stays on the text box");

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(await engine.export());
  const reHl = liveHighlights(reopened, slide);
  assert.ok(reHl.some((h) => h.color === "00FF00"), "surviving highlight persists after export + reopen");
  assert.ok(!reHl.some((h) => h.color === "FFFF00"), "deleted highlight stays gone after export + reopen");
});

test("highlight fidelity: duplicating a highlighted shape highlights the copy too", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;

  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: 20 }], { highlight: "FFFF00" });
  const before = liveHighlights(engine, slide);
  assert.ok(before.length > 0, "highlight applied");

  const dupIndex = await engine.duplicateShape(slide, 0);
  assert.ok(Number.isInteger(dupIndex), "duplicate returns the new shape index");

  const after = liveHighlights(engine, slide);
  const highlightedShapes = new Set(after.map((h) => h.shapeIndex));
  assert.ok(highlightedShapes.size >= 2, "both the original shape and its duplicate are highlighted");

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(await engine.export());
  assert.ok(
    new Set(liveHighlights(reopened, slide).map((h) => h.shapeIndex)).size >= 2,
    "both highlighted shapes persist after export + reopen"
  );
});

test("highlight fidelity: an OOXML edit after a shape move keeps the highlight (re-graft path)", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;

  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: 20 }], { highlight: "FFFF00" });
  const beforeCoverage = highlightCoverage(engine, slide);
  assert.ok(beforeCoverage.length > 0, "highlight applied");

  // A Wasm-primitive move strips the highlights from the renderer model...
  await engine.updateShapeTransform(slide, 0, { x: 100000, y: 100000, cx: 4000000, cy: 1000000, rot: 0 });
  // ...then a separate OOXML-path edit reads that now-lossy slide. editSlideShape
  // re-grafts the cached highlights before applying the bold; the bold splits the
  // highlighted run in two but the highlighted characters are unchanged.
  await engine.setRunStyleForRange(slide, 0, 0, 0, 5, { bold: true });

  assert.deepEqual(
    highlightCoverage(engine, slide),
    beforeCoverage,
    "the same characters stay highlighted after a move followed by an unrelated OOXML edit"
  );

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(await engine.export());
  assert.deepEqual(
    highlightCoverage(reopened, slide),
    beforeCoverage,
    "highlighted characters persist after export + reopen following move + edit"
  );
});

test("run-property fidelity: a shape move preserves hyperlink, underline fill, and normalizeH", async () => {
  const engine = await loadEngineWithDecoratedRun();
  const slide = 0;
  const text = "Linked decorated run";

  const before = runRprFlags(engine.renderer.getSlideOoxml(slide), text);
  assert.ok(before, "decorated run is present");
  assert.equal(before.normalizeH, "1", "normalizeH authored");
  assert.ok(before.hlinkClick && before.uFill && before.highlight, "hyperlink, underline fill, and highlight authored");

  // A Wasm-primitive move strips every un-modeled run property from the model.
  const shapeIndex = decoratedShapeIndex(engine, slide, text);
  await engine.updateShapeTransform(slide, shapeIndex, { x: 120000, y: 120000, cx: 4000000, cy: 1000000, rot: 0 });

  // The export funnel re-grafts them from the cache.
  const after = runRprFlags((await reopen(engine)).renderer.getSlideOoxml(slide), text);
  assert.ok(after, "decorated run survives export + reopen");
  assert.equal(after.normalizeH, "1", "normalizeH preserved across the move");
  assert.ok(after.hlinkClick, "run hyperlink preserved across the move");
  assert.ok(after.uFill, "underline fill preserved across the move");
  assert.ok(after.highlight, "highlight preserved across the move");
});

test("run-property fidelity: an OOXML edit after a move keeps dropped run properties (re-graft path)", async () => {
  const engine = await loadEngineWithDecoratedRun();
  const slide = 0;
  const text = "Linked decorated run";
  const shapeIndex = decoratedShapeIndex(engine, slide, text);

  // Move (model goes lossy) then a separate OOXML edit reads that lossy slide.
  // Without the editSlideShape re-graft, the bold commit would persist the loss
  // and even the export funnel could not recover it (the cache would refresh to
  // the lossy state). The re-graft restores the props before committing.
  await engine.updateShapeTransform(slide, shapeIndex, { x: 120000, y: 120000, cx: 4000000, cy: 1000000, rot: 0 });
  // Bold the whole run so it stays a single run (a partial range would split it).
  const runLength = engine.getParagraphRunText(slide, shapeIndex, 0).length;
  await engine.setRunStyleForRange(slide, shapeIndex, 0, 0, runLength, { bold: true });

  const after = runRprFlags((await reopen(engine)).renderer.getSlideOoxml(slide), text);
  assert.ok(after, "decorated run survives the move + edit");
  assert.equal(after.normalizeH, "1", "normalizeH survives move + OOXML edit");
  assert.ok(after.hlinkClick, "run hyperlink survives move + OOXML edit");
  assert.ok(after.uFill, "underline fill survives move + OOXML edit");
  assert.ok(after.highlight, "highlight survives move + OOXML edit");
});

test("single-slide edit path: text edits reuse the renderer instead of reloading the deck", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const match = "automate";

  // The renderer is build-patched with the single-slide entry point.
  assert.equal(typeof engine.renderer.loadSlideXml, "function", "loadSlideXml entry point present");

  // Capture the live renderer; the fast path must mutate it in place (no new
  // PptxRenderer, no loadPptx teardown).
  const rendererBefore = engine.renderer;
  const before = paragraphRuns(engine, slide, match);

  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 8, end: 20 }], { highlight: "FFFF00" });
  assert.equal(engine.renderer, rendererBefore, "edit did not tear down / replace the renderer");

  // Edit is observable through the renderer and a second edit stacks on it.
  assert.ok(
    paragraphHighlightSpan(engine, slide, 0, 0),
    "highlight applied via the single-slide path"
  );
  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 8, end: 20 }], { highlight: null });
  assert.equal(engine.renderer, rendererBefore, "second edit still reuses the renderer");
  assert.equal(paragraphHighlightSpan(engine, slide, 0, 0), null, "highlight cleared via the single-slide path");

  // Export still captures the in-place edits, and the reopened deck round-trips:
  // text preserved and the highlight we cleared is gone.
  const exported = await engine.export();
  assert.ok(exported.byteLength > 0, "deck still exports after in-place edits");

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reopened = await PresentationEngine.load(exported);
  const reopenedRuns = paragraphRuns(reopened, slide, match);
  assert.equal(reopenedRuns.text, before.text, "text preserved across export/reopen");
  assert.equal(paragraphHighlightSpan(reopened, slide, 0, 0), null, "cleared highlight stayed cleared after reopen");
});

test("run normalization: clearing a highlight collapses the split runs back to identical rPr", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const match = "automate";
  const before = paragraphRuns(engine, slide, match);
  assert.ok(before, "paragraph found");
  const originalText = before.text;

  // Highlight a middle sub-range (forces boundary splits), then clear it. After
  // clearing, every run carries identical rPr again, so they must coalesce.
  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 12, end: 28 }], { highlight: "FFFF00" });
  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 12, end: 28 }], { highlight: null });

  const after = paragraphRuns(engine, slide, match);
  assert.equal(after.text, originalText, "text is preserved exactly across the round trip");
  assert.equal(after.count, 1, "identical-rPr runs are merged into a single run");
});

test("run normalization: repeated range edits do not grow the run count without bound", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const match = "automate";
  const originalText = paragraphRuns(engine, slide, match).text;

  // Hammer many disjoint sub-ranges with highlight, then clear them all.
  for (const [start, end] of [[2, 6], [10, 18], [22, 30], [33, 41], [44, 52]]) {
    await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start, end }], { highlight: "FFFF00" });
  }
  await engine.setRunStyleForRanges(slide, 0, [{ paragraphIndex: 0, start: 0, end: originalText.length }], { highlight: null });

  const after = paragraphRuns(engine, slide, match);
  assert.equal(after.text, originalText, "text preserved after many edits");
  assert.equal(after.count, 1, "all identical runs collapse back to one after clearing");
});

// --- Incremental single-shape re-render contract ----------------------------
// The view's incremental render swaps just the edited `<g data-ooxml-shape-idx>`
// using engine.renderShape() instead of rebuilding the whole slide SVG. That is
// only correct if the single-shape render is byte-identical to the same shape's
// group inside a full slide render -- both before and after an edit. These lock
// that invariant in at the engine level (the view DOM swap itself is exercised
// by hand, but this guards the data the swap depends on).

const SVG_FRAGMENT_WRAPPER_OPEN =
  '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">';

function topLevelShapeGroup(svg, shapeIndex) {
  const doc = new globalThis.DOMParser().parseFromString(svg, "image/svg+xml");
  return (
    Array.from(doc.getElementsByTagName("g")).find((g) => {
      if (g.getAttribute("data-ooxml-shape-idx") !== String(shapeIndex)) return false;
      for (let n = g.parentNode; n && n.getAttribute; n = n.parentNode) {
        if (n.getAttribute("data-ooxml-shape-idx") !== null) return false;
      }
      return true;
    }) ?? null
  );
}

// Parse the bare `<g>` fragment renderShape returns exactly the way the view does
// (wrapped in a namespaced <svg>, then the group extracted), so the serialized
// form is comparable to the group lifted out of a full slide render.
function shapeGroupFromFragment(fragment, shapeIndex) {
  return topLevelShapeGroup(`${SVG_FRAGMENT_WRAPPER_OPEN}${fragment}</svg>`, shapeIndex);
}

function serializeNode(node) {
  return node ? new globalThis.XMLSerializer().serializeToString(node) : null;
}

function topLevelShapeIndices(svg) {
  const doc = new globalThis.DOMParser().parseFromString(svg, "image/svg+xml");
  const indices = [];
  for (const g of Array.from(doc.getElementsByTagName("g"))) {
    const idx = g.getAttribute("data-ooxml-shape-idx");
    if (idx === null) continue;
    let nested = false;
    for (let n = g.parentNode; n && n.getAttribute; n = n.parentNode) {
      if (n.getAttribute("data-ooxml-shape-idx") !== null) nested = true;
    }
    if (!nested) indices.push(Number(idx));
  }
  return indices;
}

test("incremental render: renderShape reproduces each slide shape group verbatim", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const slideSvg = engine.renderSlide(slide).svg;
  const indices = topLevelShapeIndices(slideSvg);
  assert.ok(indices.length > 0, "slide has top-level shapes");

  for (const idx of indices) {
    const fragment = engine.renderShape(slide, idx);
    const fromShape = shapeGroupFromFragment(fragment, idx);
    assert.ok(fromShape, `renderShape(${idx}) yields a top-level shape group`);
    assert.equal(
      fromShape.getAttribute("data-ooxml-shape-idx"),
      String(idx),
      "fragment root is the requested shape group"
    );
    assert.equal(
      serializeNode(fromShape),
      serializeNode(topLevelShapeGroup(slideSvg, idx)),
      `single-shape render matches the full-slide group for shape ${idx}`
    );
  }
});

test("currentBuffer sync: a fast-path edit is folded in before applyListStyle reads currentBuffer", async () => {
  // Fast-path (`loadSlideXml`) commits update only the renderer model and leave
  // `currentBuffer` behind. `applyListStyle` reconciles a fresh renderer export
  // against `currentBuffer` (it merges graphic frames / package parts / extension
  // lists from it), so the engine must fold the pending fast-path slide XML into
  // `currentBuffer` first. This guards that the sync runs and the funnel stays
  // lossless: the fast-path edit, the list edit, and the renderer-dropped chart
  // graphic frame + extension list all survive together.
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const shapeIndex = 0;
  const paragraphIndex = 0;

  // Fast-path edit: records pending slide XML, does not touch currentBuffer.
  const runLength = engine.getParagraphRunText(slide, shapeIndex, paragraphIndex).length;
  await engine.setRunStyleForRange(slide, shapeIndex, 0, 0, runLength, { bold: true });

  // currentBuffer consumer: must sync (fold pending XML, preserve dropped content)
  // before its merges read currentBuffer.
  await engine.applyListStyle(slide, shapeIndex, paragraphIndex, "number");

  const exported = await engine.export();
  const zip = await extractZip(exported);
  const slideXml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.ok(slideXml, "slide XML present after the funnel");
  assert.match(slideXml, /b="1"/, "fast-path bold survived the fold -> applyListStyle funnel");
  assert.match(slideXml, /buAutoNum/, "numbered list applied");
  assert.match(slideXml, /extLst/, "slide extension list preserved through the fold");
  assert.match(slideXml, /graphicFrame/, "graphic frames preserved through the fold");
  assert.ok(
    [...zip.textFiles.keys()].some((part) => part.startsWith("ppt/charts/")),
    "chart part preserved"
  );

  const reopened = await reopen(engine);
  assert.equal(reopened.slideCount, 1, "deck still round-trips after the funnel");
  assert.match(reopened.renderSlide(slide).svg, /^<svg\b/, "reopened slide still renders");
});

test("incremental render: an edited shape's single render still matches the full slide", async () => {
  const engine = await loadEngineWithBulletParagraph();
  const slide = 0;
  const shapeIndex = 0;
  const paragraphIndex = 0;

  const beforeFragment = serializeNode(shapeGroupFromFragment(engine.renderShape(slide, shapeIndex), shapeIndex));

  // Underline renders into the SVG (text-decoration), so the shape group must
  // visibly change -- proving the swapped-in node carries the edit.
  const ooxmlText = engine.getParagraphRunText(slide, shapeIndex, paragraphIndex);
  await engine.setRunStyleForRanges(
    slide,
    shapeIndex,
    [{ paragraphIndex, start: 0, end: ooxmlText.length }],
    { underline: true }
  );

  const slideSvg = engine.renderSlide(slide).svg;
  const fromSlide = serializeNode(topLevelShapeGroup(slideSvg, shapeIndex));
  const fromShape = serializeNode(shapeGroupFromFragment(engine.renderShape(slide, shapeIndex), shapeIndex));

  assert.notEqual(fromShape, beforeFragment, "the edit changed the single-shape render");
  assert.equal(fromShape, fromSlide, "post-edit single-shape render equals the full-slide shape group");

  // Other shapes are untouched: their single render is unchanged by the edit, so
  // an incremental swap of shape 0 leaves them byte-identical.
  for (const idx of topLevelShapeIndices(slideSvg)) {
    if (idx === shapeIndex) continue;
    assert.equal(
      serializeNode(shapeGroupFromFragment(engine.renderShape(slide, idx), idx)),
      serializeNode(topLevelShapeGroup(slideSvg, idx)),
      `unedited shape ${idx} renders identically`
    );
  }
});
