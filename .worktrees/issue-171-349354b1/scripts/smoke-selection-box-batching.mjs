// Guards the geometry-batching optimization in getSvgInlineSelectionBoxes.
//
// The view reconstructs a contiguous selection/highlight box for a run tspan
// from only its FIRST and LAST glyph extents (2 getExtentOfChar calls), instead
// of unioning every glyph (O(chars)). That is exact ONLY while Chrome returns a
// uniform y/height for every glyph in a run (the line box, not per-glyph ink).
//
// This harness renders the real engine SVG in the Obsidian-matched Electron
// Chromium and asserts, per run tspan, that the endpoint-union box equals the
// per-glyph-union box. If a future Chromium ever returns per-glyph ink metrics,
// this turns red and the batching must revert to a per-glyph sweep.
//
// Usage: node scripts/smoke-selection-box-batching.mjs

import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildZip, extractZip } from 'pptx-svg';
import { readDeck, toArrayBuffer } from '../tests/helpers/renderer.mjs';
import { loadPresentationEngineModule } from '../tests/helpers/load-plugin-modules.mjs';
import { runHeadlessHarness } from './lib/text-offset-harness.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.resolve('scripts/visual-output');
const htmlPath = path.join(outputDir, 'selection-box-batching.html');

const fixtureTitleParagraph =
  '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>';
const bulletParagraph =
  '<a:p>' +
  '<a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="\u25CF"/></a:pPr>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>Because sm</a:t></a:r>' +
  '<a:r><a:rPr lang="en-US" sz="1800"/><a:t>art home automate, and manage all your Samsung and SmartThings-compatible appliances.</a:t></a:r>' +
  '<a:endParaRPr lang="en-US"/>' +
  '</a:p>';

async function renderFixtureSvg() {
  const input = await readDeck('features.pptx');
  const source = toArrayBuffer(input);
  const zip = await extractZip(source);
  const slidePath = 'ppt/slides/slide1.xml';
  const slideXml = zip.textFiles.get(slidePath);
  assert.ok(slideXml, 'features.pptx slide1.xml not found');
  const patched = await buildZip(source, new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, bulletParagraph)]]));
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(patched);
  return engine.renderSlide(0).svg;
}

const driver = String.raw`
(function () {
  const svg = document.querySelector('svg');
  const rootMatrix = svg.getScreenCTM();
  const rootInverse = rootMatrix.inverse();

  function transform(rect, elementMatrix) {
    const m = rootInverse.multiply(elementMatrix);
    const pts = [
      new DOMPoint(rect.x, rect.y),
      new DOMPoint(rect.x + rect.width, rect.y),
      new DOMPoint(rect.x, rect.y + rect.height),
      new DOMPoint(rect.x + rect.width, rect.y + rect.height)
    ].map((p) => p.matrixTransform(m));
    const xs = pts.map((p) => p.x);
    const ys = pts.map((p) => p.y);
    const x = Math.min.apply(null, xs);
    const y = Math.min.apply(null, ys);
    return { x, y, width: Math.max.apply(null, xs) - x, height: Math.max.apply(null, ys) - y };
  }

  function mergeRow(rows, box) {
    const centerY = box.y + box.height / 2;
    const row = rows.find((c) => Math.abs(centerY - (c.y + c.height / 2)) < Math.max(2, box.height * 0.55));
    if (!row) { rows.push(Object.assign({}, box)); return; }
    const left = Math.min(row.x, box.x);
    const top = Math.min(row.y, box.y);
    const right = Math.max(row.x + row.width, box.x + box.width);
    const bottom = Math.max(row.y + row.height, box.y + box.height);
    row.x = left; row.y = top; row.width = right - left; row.height = bottom - top;
  }

  // Reference: union every glyph (the pre-optimization algorithm).
  function perGlyphBoxes(span) {
    const m = span.getScreenCTM();
    const n = span.getNumberOfChars();
    const rows = [];
    for (let i = 0; i < n; i++) {
      let b = null;
      try { b = transform(span.getExtentOfChar(i), m); } catch (e) { b = null; }
      if (b && b.width >= 0 && b.height > 0) mergeRow(rows, b);
    }
    return rows;
  }

  // Optimized: endpoint union (mirrors view.measureSpanRangeBoxes fast path).
  function endpointBoxes(span) {
    const m = span.getScreenCTM();
    const n = span.getNumberOfChars();
    if (n <= 0) return [];
    const first = transform(span.getExtentOfChar(0), m);
    const last = n === 1 ? first : transform(span.getExtentOfChar(n - 1), m);
    const sameRow = Math.abs((first.y + first.height / 2) - (last.y + last.height / 2)) < Math.max(2, first.height * 0.55);
    if (!sameRow) return perGlyphBoxes(span); // fallback path
    const x = Math.min(first.x, last.x);
    const y = Math.min(first.y, last.y);
    const right = Math.max(first.x + first.width, last.x + last.width);
    const bottom = Math.max(first.y + first.height, last.y + last.height);
    return [{ x, y, width: right - x, height: bottom - y }];
  }

  function approxEqual(a, b) {
    const e = 0.01;
    return a.length === b.length && a.every((box, i) =>
      Math.abs(box.x - b[i].x) < e && Math.abs(box.y - b[i].y) < e &&
      Math.abs(box.width - b[i].width) < e && Math.abs(box.height - b[i].height) < e);
  }

  const runs = Array.from(svg.querySelectorAll('tspan[data-ooxml-run-idx]'))
    .filter((s) => !s.querySelector('tspan') && s.getNumberOfChars() > 0);

  const result = { ok: true, checked: 0, mismatches: [] };
  for (const span of runs) {
    const reference = perGlyphBoxes(span);
    const optimized = endpointBoxes(span);
    result.checked++;
    if (!approxEqual(optimized, reference)) {
      result.ok = false;
      result.mismatches.push({ text: (span.textContent || '').slice(0, 30), glyphs: span.getNumberOfChars(), reference, optimized });
    }
  }

  document.body.dataset.metrics = encodeURIComponent(JSON.stringify(result));
})();
`;

const svg = await renderFixtureSvg();
const html = String.raw`<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Selection Box Batching Smoke</title>
<style>body{margin:0;padding:24px}svg{display:block}</style></head>
<body><div class="native-powerpoint-canvas-pane"><div class="native-powerpoint-slide-surface">${svg}</div></div>
<script>window.addEventListener('load',()=>{try{${driver}}catch(e){document.body.dataset.metrics=encodeURIComponent(JSON.stringify({ok:false,error:String(e&&e.stack||e)}))}});</script>
</body></html>`;

await mkdir(outputDir, { recursive: true });
await writeFile(htmlPath, html, 'utf8');
const harness = await runHeadlessHarness(htmlPath, { chromeTimeoutMs: 25000 });
if (!harness) {
  console.log('No headless browser runtime available; skipping selection-box-batching smoke.');
  process.exit(0);
}
const { metrics, runtime } = harness;

if (metrics.error) {
  console.error('Harness error:', metrics.error);
  process.exit(1);
}

console.log(`Runtime: ${runtime}`);
console.log(`Run tspans checked: ${metrics.checked}`);
assert.ok(metrics.checked > 0, 'expected at least one run tspan to check');
assert.ok(
  metrics.ok,
  `endpoint-union box diverged from per-glyph union (getExtentOfChar is no longer uniform per glyph):\n${JSON.stringify(metrics.mismatches, null, 2)}`
);
console.log('\nSelection box batching smoke passed: endpoint union == per-glyph union on every run tspan.');
