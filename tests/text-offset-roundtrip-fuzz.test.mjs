// Property/fuzz coverage: offset round-trip across real decks.
//
// For every text paragraph the renderer emits, stamp offsets (the same path the
// view uses after render) and assert:
//   - alignment reconciles
//   - stamped attributes match alignRunTspansToOoxml tiles
//   - mapEditorRangeToOoxml on stamped tiles matches the reference subsequence walk
//
// Runs headlessly in Node against rendered SVG parsed with @xmldom/xmldom.

import { test } from "node:test";
import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";
import {
  loadAnnotateTextOffsetsModule,
  loadPresentationEngineModule,
  loadTextUtilsModule,
} from "./helpers/load-plugin-modules.mjs";
import { parseSvgForStamping } from "./helpers/svg-xml-dom.mjs";
import {
  assertStampsMatchAlignment,
  fuzzEditorRangeRoundTrip,
} from "./helpers/text-offset-verify.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const DECK_CASES = [
  { name: "features.pptx" },
  { name: "simple-edit.pptx" },
  { name: "large-deck.pptx" },
];

const SAMPLE_DECK = path.join(projectRoot, "test_files/10MB-Sample-PPT-File.pptx");
if (existsSync(SAMPLE_DECK)) {
  DECK_CASES.push({ name: SAMPLE_DECK, slides: [0, 1, 2], external: true });
}

function sampleSlideIndices(slideCount) {
  if (slideCount <= 0) return [];
  if (slideCount === 1) return [0];
  const picks = new Set([
    0,
    Math.floor(slideCount / 4),
    Math.floor(slideCount / 2),
    Math.floor((slideCount * 3) / 4),
    slideCount - 1,
  ]);
  return [...picks].filter((index) => index >= 0 && index < slideCount).sort((a, b) => a - b);
}

async function loadEngineForDeck(deckCase) {
  const bytes = deckCase.external
    ? await import("node:fs/promises").then((fs) => fs.readFile(deckCase.name))
    : await readDeck(deckCase.name);
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(toArrayBuffer(bytes));
}

function collectParagraphs(svg, engine, slideIndex) {
  const {
    annotateSlideTextOffsets,
    collectRunSpansByParagraph,
    readStampedTiles,
  } = annotateModule;
  const { alignRunTspansToOoxml } = textUtilsModule;

  annotateSlideTextOffsets(svg, (shapeIndex, paragraphIndex) =>
    engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex),
  );

  const paragraphs = [];
  for (const shapeGroup of Array.from(svg.querySelectorAll('g[data-ooxml-shape-idx]'))) {
    const shapeIndex = Number(shapeGroup.getAttribute("data-ooxml-shape-idx"));
    for (const [paragraphIndex, runSpans] of collectRunSpansByParagraph(shapeGroup)) {
      const ooxmlText = engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex);
      if (ooxmlText === null) continue;
      const runTexts = runSpans.map((span) => span.textContent || "");
      const editorText = runTexts.join("");
      const alignment = alignRunTspansToOoxml(runTexts, ooxmlText);
      const tiles = readStampedTiles(runSpans);
      paragraphs.push({
        deck: deckCaseLabel,
        slideIndex,
        shapeIndex,
        paragraphIndex,
        editorText,
        ooxmlText,
        alignment,
        tiles,
        runSpans,
      });
    }
  }
  return paragraphs;
}

let annotateModule;
let textUtilsModule;
let deckCaseLabel = "";

test("text offset round-trip fuzz across real decks", async () => {
  annotateModule = await loadAnnotateTextOffsetsModule();
  textUtilsModule = await loadTextUtilsModule();
  const { mapEditorRangeToOoxml, mapEditorOffsetToOoxmlOffset } = textUtilsModule;

  const failures = [];
  let paragraphTotal = 0;

  for (const deckCase of DECK_CASES) {
    deckCaseLabel = deckCase.external ? path.basename(deckCase.name) : deckCase.name;
    const engine = await loadEngineForDeck(deckCase);
    const slideIndices = deckCase.slides ?? sampleSlideIndices(engine.slideCount);

    for (const slideIndex of slideIndices) {
      const svg = parseSvgForStamping(engine.renderSlide(slideIndex).svg);
      const paragraphs = collectParagraphs(svg, engine, slideIndex);
      paragraphTotal += paragraphs.length;

      for (const paragraph of paragraphs) {
        if (!paragraph.alignment.reconciled) {
          failures.push({
            kind: "reconcile",
            deck: paragraph.deck,
            slide: paragraph.slideIndex,
            shape: paragraph.shapeIndex,
            para: paragraph.paragraphIndex,
            editorLen: paragraph.editorText.length,
            ooxmlLen: paragraph.ooxmlText.length,
          });
          continue;
        }

        try {
          assertStampsMatchAlignment(paragraph.runSpans, paragraph.alignment);
        } catch (error) {
          failures.push({
            kind: "stamp-mismatch",
            deck: paragraph.deck,
            slide: paragraph.slideIndex,
            shape: paragraph.shapeIndex,
            para: paragraph.paragraphIndex,
            message: error.message,
          });
          continue;
        }

        assert.ok(paragraph.tiles, `missing stamped tiles: ${deckCaseLabel} slide ${slideIndex}`);
        const fuzz = fuzzEditorRangeRoundTrip({
          editorText: paragraph.editorText,
          ooxmlText: paragraph.ooxmlText,
          tiles: paragraph.tiles,
          mapEditorRangeToOoxml,
          mapEditorOffsetToOoxmlOffset,
          iterations: 40,
          seed: slideIndex * 997 + paragraph.shapeIndex * 31 + paragraph.paragraphIndex,
        });
        if (fuzz.mismatches.length > 0) {
          failures.push({
            kind: "range-roundtrip",
            deck: paragraph.deck,
            slide: paragraph.slideIndex,
            shape: paragraph.shapeIndex,
            para: paragraph.paragraphIndex,
            mismatches: fuzz.mismatches.slice(0, 3),
          });
        }
      }
    }
  }

  assert.ok(paragraphTotal > 0, "expected at least one stamped paragraph across decks");
  assert.deepEqual(
    failures,
    [],
    `offset round-trip failures (${failures.length}):\n${JSON.stringify(failures.slice(0, 8), null, 2)}`,
  );
});
