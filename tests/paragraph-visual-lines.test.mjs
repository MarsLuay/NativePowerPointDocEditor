import assert from "node:assert/strict";
import test from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

// Mirrors joinParagraphVisualLines in src/powerpoint/textUtils.ts.
function joinParagraphVisualLines(lineTexts) {
  return lineTexts.join("");
}

// Mirrors mapFlatOffsetToRunLine in src/powerpoint/textUtils.ts.
function mapFlatOffsetToRunLine(runLineCharCounts, flatOffset) {
  if (runLineCharCounts.length === 0) {
    return { lineIndex: 0, localOffset: Math.max(0, flatOffset) };
  }
  const clamped = Math.max(0, flatOffset);
  let offset = 0;
  for (let lineIndex = 0; lineIndex < runLineCharCounts.length; lineIndex++) {
    const count = Math.max(0, runLineCharCounts[lineIndex] ?? 0);
    if (clamped <= offset + count) {
      return { lineIndex, localOffset: clamped - offset };
    }
    offset += count;
  }
  const lastIndex = runLineCharCounts.length - 1;
  return { lineIndex: lastIndex, localOffset: Math.max(0, runLineCharCounts[lastIndex] ?? 0) };
}

// Mirrors mapFlatRangeToRunLineSegments in src/powerpoint/textUtils.ts.
function mapFlatRangeToRunLineSegments(runLineCharCounts, flatStart, flatEnd) {
  const start = Math.max(0, Math.min(flatStart, flatEnd));
  const end = Math.max(flatStart, flatEnd);
  const segments = [];
  let offset = 0;
  for (let lineIndex = 0; lineIndex < runLineCharCounts.length; lineIndex++) {
    const count = Math.max(0, runLineCharCounts[lineIndex] ?? 0);
    const localStart = Math.max(0, start - offset);
    const localEnd = Math.min(count, end - offset);
    if (localEnd > localStart) {
      segments.push({ lineIndex, localStart, localEnd });
    }
    offset += count;
  }
  return segments;
}

test("joinParagraphVisualLines matches OOXML flat offsets for soft-wrapped lines", () => {
  const flat = joinParagraphVisualLines([
    "Because sm",
    "art home automate, and manage all your Samsung and ",
    "SmartThings-compatible appliances"
  ]);
  assert.equal(flat, "Because smart home automate, and manage all your Samsung and SmartThings-compatible appliances");

  const automateStart = flat.indexOf("automate");
  assert.ok(automateStart >= 0);
  assert.equal(flat.slice(automateStart, automateStart + "automate".length), "automate");

  const legacyEditor = [
    "Because sm",
    "art home automate, and manage all your Samsung and ",
    "SmartThings-compatible appliances"
  ].join("\n");
  const legacyOffset = legacyEditor.indexOf("automate");
  assert.notEqual(legacyOffset, automateStart, "synthetic newlines shift selection offsets");
});

test("flat offset maps onto the correct wrapped run line", () => {
  // Run-only visual line lengths (bullet line excluded by the caller).
  const lines = ["Because sm", "art home automate, ", "and manage all."];
  const flat = lines.join("");
  const counts = lines.map((line) => line.length);

  const automateStart = flat.indexOf("automate");
  const mapped = mapFlatOffsetToRunLine(counts, automateStart);
  // "automate" starts on the second visual line.
  assert.equal(mapped.lineIndex, 1);
  assert.equal(lines[mapped.lineIndex].slice(mapped.localOffset, mapped.localOffset + 8), "automate");
});

test("flat range splits into per-line highlight segments without bullet drift", () => {
  const lines = ["Because sm", "art home automate, ", "and manage all."];
  const flat = lines.join("");
  const counts = lines.map((line) => line.length);

  const start = flat.indexOf("automate");
  const end = start + "automate".length;
  const segments = mapFlatRangeToRunLineSegments(counts, start, end);

  // The whole word lives on line index 1; reconstruct the highlighted text.
  const highlighted = segments
    .map((segment) => lines[segment.lineIndex].slice(segment.localStart, segment.localEnd))
    .join("");
  assert.equal(highlighted, "automate");
});

test("selection that wraps across two lines highlights exactly the selected text", () => {
  const lines = ["Because sm", "art home automate"];
  const flat = lines.join(""); // "Because smart home automate"
  const counts = lines.map((line) => line.length);

  // Select "smart" which straddles the soft wrap between line 0 and line 1.
  const start = flat.indexOf("smart");
  const end = start + "smart".length;
  const segments = mapFlatRangeToRunLineSegments(counts, start, end);

  const highlighted = segments
    .map((segment) => lines[segment.lineIndex].slice(segment.localStart, segment.localEnd))
    .join("");
  assert.equal(highlighted, "smart");
  assert.equal(segments.length, 2, "selection spans both wrapped lines");
});

const fixtureTitleParagraph =
  '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>';

const bulletParagraph =
  '<a:p>' +
  '<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="\u25CF"/></a:pPr>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>Because sm</a:t></a:r>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>art home automate, and manage all your Samsung and SmartThings-compatible appliances.</a:t></a:r>' +
  '<a:endParaRPr lang="en-US"/>' +
  '</a:p>';

async function renderBulletedParagraphSvg() {
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
  const engine = await PresentationEngine.load(patched);
  return engine.renderSlide(0).svg;
}

test("renderer keeps the bullet glyph in a run-less line container", async () => {
  // The fix relies on bullets being rendered as their own data-ooxml-para-idx
  // tspan with no data-ooxml-run-idx child. If the renderer ever inlines the
  // bullet into a run-bearing container, the offset math must be revisited.
  const svg = await renderBulletedParagraphSvg();
  assert.ok(svg.includes("\u25CF"), "bullet glyph should be rendered");

  const bulletContainer = svg.match(
    /<tspan\b[^>]*data-ooxml-para-idx="0"[^>]*data-ooxml-bu-font="[^"]*"[^>]*>\s*<tspan[^>]*>\u25CF/
  );
  assert.ok(bulletContainer, "bullet should live in a dedicated para container with a bullet font marker");

  // The bullet's inner tspan must not carry a run index.
  const bulletInner = svg.match(/data-ooxml-bu-font="[^"]*"[^>]*>\s*<tspan([^>]*)>\u25CF/);
  assert.ok(bulletInner, "bullet inner tspan present");
  assert.ok(
    !/data-ooxml-run-idx/.test(bulletInner[1]),
    "bullet glyph tspan must not be a run (would desync offsets)"
  );
});

test("setRunStyleForRanges underlines exactly the run-only range on a bulleted paragraph", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, bulletParagraph)]])
  );
  const engine = await PresentationEngine.load(patched);

  // Run-only flat text (bullet excluded), exactly what the inline editor holds.
  const flat = "Because smart home automate, and manage all your Samsung and SmartThings-compatible appliances.";
  const start = flat.indexOf("automate");
  const end = start + "automate".length;

  await engine.setRunStyleForRanges(0, 0, [{ paragraphIndex: 0, start, end }], { underline: true });

  // Only the "automate" word should be underlined; "Because " before it stays plain.
  assert.equal(engine.isRangeStyled(0, 0, 0, start, end, "underline"), true);
  assert.equal(engine.isRangeStyled(0, 0, 0, 0, start, "underline"), false);
});

test("geometry index skips bullet prefix when mapping to run-only offsets", () => {
  // Simulates a visual line whose SVG leaves are: bullet (2 chars) + run text.
  const leafSpans = [
    { isRun: false, count: 2 },
    { isRun: true, count: 50 },
  ];

  function geometryIndexToRunOffset(leafSpans, geometryIndex) {
    const clamped = Math.max(0, geometryIndex);
    let runOffset = 0;
    let leafOffset = 0;
    for (const span of leafSpans) {
      const count = Math.max(0, span.count);
      if (clamped <= leafOffset + count) {
        if (span.isRun) runOffset += Math.max(0, clamped - leafOffset);
        return runOffset;
      }
      if (span.isRun) runOffset += count;
      leafOffset += count;
    }
    return runOffset;
  }

  function runOffsetToGeometryIndex(leafSpans, runOffset) {
    const clamped = Math.max(0, runOffset);
    let runSeen = 0;
    let leafOffset = 0;
    for (const span of leafSpans) {
      const count = Math.max(0, span.count);
      if (span.isRun) {
        if (clamped <= runSeen + count) {
          return leafOffset + Math.max(0, clamped - runSeen);
        }
        runSeen += count;
      }
      leafOffset += count;
    }
    return leafOffset;
  }

  // Geometry index includes a 2-char bullet prefix that is not in the editor string.
  assert.equal(geometryIndexToRunOffset(leafSpans, 14), 12);
  assert.equal(geometryIndexToRunOffset(leafSpans, 16), 14);
  assert.equal(runOffsetToGeometryIndex(leafSpans, 12), 14);
  assert.equal(runOffsetToGeometryIndex(leafSpans, 14), 16);
  assert.equal(geometryIndexToRunOffset(leafSpans, 2), 0);
});

test("underlining 'om' inside automate does not drift to 'at' on a bulleted paragraph", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, bulletParagraph)]])
  );
  const engine = await PresentationEngine.load(patched);

  // Run-only flat text (bullet excluded), exactly what the inline editor holds.
  const flat = "Because smart home automate, and manage all your Samsung and SmartThings-compatible appliances.";
  const automate = flat.indexOf("automate");
  // The "om" inside "automate" (a-u-t-o-m...).
  const omStart = automate + 3;
  const omEnd = omStart + 2;
  assert.equal(flat.slice(omStart, omEnd), "om");

  await engine.setRunStyleForRanges(0, 0, [{ paragraphIndex: 0, start: omStart, end: omEnd }], { underline: true });

  assert.equal(engine.isRangeStyled(0, 0, 0, omStart, omEnd, "underline"), true);
  // The "at" two characters to the right (the old +bullet drift target) must stay plain.
  assert.equal(engine.isRangeStyled(0, 0, 0, omStart + 2, omEnd + 2, "underline"), false);
});

// Mirrors mapEditorOffsetToOoxmlOffset in src/powerpoint/textUtils.ts.
function mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, editorOffset, consumeTrailingGap = false) {
  const target = Math.max(0, Math.min(editorOffset, editorText.length));
  if (editorText === ooxmlText) return target;
  let e = 0;
  let o = 0;
  while (e < target && o < ooxmlText.length) {
    if (editorText[e] === ooxmlText[o]) {
      e++;
      o++;
    } else {
      o++;
    }
  }
  if (consumeTrailingGap) {
    while (o < ooxmlText.length && (e >= editorText.length || editorText[e] !== ooxmlText[o])) {
      o++;
    }
  }
  return o;
}

test("editor offsets map across wrap-dropped whitespace into OOXML offsets", () => {
  // OOXML keeps the spaces swallowed at two soft-wrap boundaries; the editor
  // (SVG) text drops them, so editor offsets run 2 short of OOXML offsets.
  const ooxml = "manage all your Samsung and SmartThings-compatible appliances and electronics with a single, easy-to-use app.";
  const editor = ooxml.replace("Samsung and", "Samsungand").replace("with a", "witha");
  assert.equal(editor.length, ooxml.length - 2);

  // A character that sits after both dropped spaces is shifted by +2.
  const editorIdx = editor.indexOf("easy-to-use");
  const ooxmlIdx = ooxml.indexOf("easy-to-use");
  assert.equal(ooxmlIdx - editorIdx, 2);
  assert.equal(mapEditorOffsetToOoxmlOffset(editor, ooxml, editorIdx), ooxmlIdx);
});

test("regression: clearing through 'to' includes the trailing run in OOXML space", () => {
  // The reported bug: a visual selection that reaches "...easy-to" left "o" (and
  // the wrap-space) highlighted because the editor end offset was passed raw as
  // an OOXML offset, landing 2 chars short of the real runs.
  const ooxml = "and manage all your Samsung and SmartThings-compatible appliances and electronics with a single, easy-to.";
  const editor = ooxml.replace("Samsung and", "Samsungand").replace("with a", "witha");
  assert.equal(editor.length, ooxml.length - 2);

  // Visual selection covers everything up to and including the final "to".
  const editorEnd = editor.length; // user selected to the very end
  const mappedEnd = mapEditorOffsetToOoxmlOffset(editor, ooxml, editorEnd, true);
  assert.equal(mappedEnd, ooxml.length, "END maps to the true OOXML paragraph end");

  // The "o" in the trailing "to" lives within [start, mappedEnd) now.
  const oIndex = ooxml.lastIndexOf("o"); // the o in "to."
  assert.ok(oIndex < mappedEnd, "the trailing 'o' is inside the cleared OOXML range");
});

test("identity mapping when editor and OOXML text are equal", () => {
  const text = "no soft wrap here";
  for (let i = 0; i <= text.length; i++) {
    assert.equal(mapEditorOffsetToOoxmlOffset(text, text, i), i);
    assert.equal(mapEditorOffsetToOoxmlOffset(text, text, i, true), i);
  }
});
