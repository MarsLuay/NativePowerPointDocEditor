// Headless DOM test for annotateSlideTextOffsets end-to-end.
//
// Renders real engine SVG, mounts it in Electron (Obsidian-matched Chromium),
// runs the extracted stamping module against live DOM tspans, and verifies:
//   1. every paragraph reconciles (editor text is a subsequence of OOXML)
//   2. stamped data-ooxml-char-* attributes match alignRunTspansToOoxml
//   3. mapEditorRangeToOoxml on stamped tiles matches the reference walk
//
// Usage:
//   node scripts/smoke-annotate-text-offsets.mjs
//   node scripts/smoke-annotate-text-offsets.mjs --deck=path/to.pptx --slide=1

import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readDeck, toArrayBuffer } from "../tests/helpers/renderer.mjs";
import { loadPresentationEngineModule } from "../tests/helpers/load-plugin-modules.mjs";
import {
  buildOoxmlTextMap,
  bundleAnnotateModule,
  bundleTextUtilsModule,
  loadDeckEngine,
  projectRoot,
  resolveElectronBinary,
  runHeadlessHarness,
  writeHarnessHtml,
} from "./lib/text-offset-harness.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve("scripts/visual-output");

function parseArgs(argv) {
  const args = { deck: null, slide: 0 };
  for (const token of argv) {
    const deck = token.match(/^--deck=(.+)$/);
    const slide = token.match(/^--slide=(\d+)$/);
    if (deck) args.deck = deck[1];
    else if (slide) args.slide = Number(slide[1]);
  }
  return args;
}

const cli = parseArgs(process.argv.slice(2));

const driverScript = String.raw`
(function () {
  const svg = document.querySelector('svg');
  const {
    annotateSlideTextOffsets,
    collectRunSpansByParagraph,
    readStampedTiles,
  } = AnnotateTextOffsetsNS;
  const {
    alignRunTspansToOoxml,
    mapEditorRangeToOoxml,
    mapEditorOffsetToOoxmlOffset,
  } = TextUtilsNS;

  const ooxmlMap = window.OOXML_TEXT;
  annotateSlideTextOffsets(svg, (shapeIdx, paraIdx) => {
    const shape = ooxmlMap[String(shapeIdx)];
    return shape ? shape[String(paraIdx)] ?? null : null;
  });

  const paragraphs = [];
  let rng = 1;
  const next = () => {
    rng = (rng * 1664525 + 1013904223) >>> 0;
    return rng / 0x100000000;
  };

  for (const shapeGroup of Array.from(svg.querySelectorAll('g[data-ooxml-shape-idx]'))) {
    const shapeIndex = Number(shapeGroup.getAttribute('data-ooxml-shape-idx'));
    const runsByParagraph = collectRunSpansByParagraph(shapeGroup);
    for (const [paragraphIndex, runSpans] of runsByParagraph) {
      const ooxmlText = ooxmlMap[String(shapeIndex)]?.[String(paragraphIndex)];
      if (ooxmlText == null) continue;
      const editorText = runSpans.map((s) => s.textContent || '').join('');
      const alignment = alignRunTspansToOoxml(
        runSpans.map((s) => s.textContent || ''),
        ooxmlText,
      );
      const tiles = readStampedTiles(runSpans);
      const entry = {
        shapeIndex,
        paragraphIndex,
        reconciled: alignment.reconciled,
        spanCount: runSpans.length,
        stampMismatch: false,
        rangeMismatches: [],
      };

      for (let i = 0; i < runSpans.length; i++) {
        const tile = alignment.spans[i];
        if (
          runSpans[i].getAttribute('data-ooxml-char-start') !== String(tile.charStart)
          || runSpans[i].getAttribute('data-ooxml-char-end') !== String(tile.charEnd)
        ) {
          entry.stampMismatch = true;
        }
      }

      const editorLen = editorText.length;
      for (let t = 0; t < 32; t++) {
        const a = Math.floor(next() * (editorLen + 1));
        const b = Math.floor(next() * (editorLen + 1));
        const start = Math.min(a, b);
        const end = Math.max(a, b);
        const mapped = tiles ? mapEditorRangeToOoxml(tiles, start, end) : null;
        const refStart = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, start, false);
        const refEnd = mapEditorOffsetToOoxmlOffset(editorText, ooxmlText, end, true);
        if (!mapped || mapped.start !== refStart || mapped.end !== refEnd) {
          entry.rangeMismatches.push({ start, end, mapped, refStart, refEnd });
        }
      }
      paragraphs.push(entry);
    }
  }

  const textParagraphs = paragraphs.filter((p) => p.spanCount > 0);
  const ok = textParagraphs.length > 0
    && textParagraphs.every((p) => p.reconciled && !p.stampMismatch && p.rangeMismatches.length === 0);

  document.body.dataset.metrics = encodeURIComponent(JSON.stringify({
    ok,
    paragraphCount: textParagraphs.length,
    paragraphs: textParagraphs,
    failures: textParagraphs.filter(
      (p) => !p.reconciled || p.stampMismatch || p.rangeMismatches.length > 0,
    ),
  }));
})();
`;

async function loadEngineAndSlide() {
  if (cli.deck) {
    const engine = await loadDeckEngine(cli.deck);
    return { engine, slideIndex: cli.slide, label: `${path.resolve(cli.deck)} slide ${cli.slide}` };
  }

  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(toArrayBuffer(await readDeck("features.pptx")));
  return { engine, slideIndex: 0, label: "features.pptx slide 0" };
}

if (!resolveElectronBinary()) {
  console.log("Electron not installed; skipping annotate-text-offsets smoke (requires real Chromium DOM).");
  process.exit(0);
}

const [{ engine, slideIndex, label }, annotateBundle, textUtilsBundle] = await Promise.all([
  loadEngineAndSlide(),
  bundleAnnotateModule(),
  bundleTextUtilsModule(),
]);

const svg = engine.renderSlide(slideIndex).svg;
const ooxmlMap = await buildOoxmlTextMap(engine, slideIndex);

const htmlPath = await writeHarnessHtml({
  outputDir,
  filename: "annotate-text-offsets.html",
  svg,
  ooxmlMap,
  annotateBundle,
  textUtilsBundle,
  driverScript,
});

const harness = await runHeadlessHarness(htmlPath);
assert.ok(harness, "no headless runtime available");
const { metrics, runtime } = harness;

if (metrics.error) {
  console.error("Harness error:", metrics.error);
  process.exit(1);
}

console.log(`Runtime: ${runtime}`);
console.log(`Source:  ${label}`);
console.log(`HTML:    ${htmlPath}`);
console.log(`Paragraphs stamped: ${metrics.paragraphCount}`);

if (!metrics.ok) {
  console.error("Failures:", JSON.stringify(metrics.failures, null, 2));
  process.exit(1);
}

console.log("\nAnnotate text offsets smoke passed: stamps reconcile and range round-trips hold on every paragraph.");
