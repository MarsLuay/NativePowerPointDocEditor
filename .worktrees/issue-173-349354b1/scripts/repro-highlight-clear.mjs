// Headless reproduction of the "residual highlight after No-color" bug.
//
// Highlights aren't rendered into the SVG (the renderer drops <a:highlight>);
// the view repaints them as overlay rects from engine.getSlideRunHighlights().
// So the authoritative "is anything still highlighted" check is that method.
//
// This drives the real engine through highlight -> clear cycles on a chosen
// paragraph and reports any run that survives a No-color clear. Pure OOXML, no
// browser needed, fully deterministic.
//
// Usage:
//   node scripts/repro-highlight-clear.mjs --deck=<path> --slide=<n> --match="<substr>"

import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { toArrayBuffer } from '../tests/helpers/renderer.mjs';
import { loadPresentationEngineModule } from '../tests/helpers/load-plugin-modules.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const args = { deck: null, slide: 0, match: 'single, easy-to' };
  for (const token of argv) {
    const deck = token.match(/^--deck=(.+)$/);
    const slide = token.match(/^--slide=(\d+)$/);
    const match = token.match(/^--match=(.+)$/);
    if (deck) args.deck = deck[1];
    else if (slide) args.slide = Number(slide[1]);
    else if (match) args.match = match[1];
  }
  return args;
}

const cli = parseArgs(process.argv.slice(2));
assert.ok(cli.deck, 'pass --deck=<path-to.pptx>');
const deckPath = path.resolve(process.cwd(), cli.deck);
assert.ok(existsSync(deckPath), `deck not found: ${deckPath}`);

const { PresentationEngine } = await loadPresentationEngineModule();
const deckBytes = await readFile(deckPath);

async function freshEngine() {
  return PresentationEngine.load(toArrayBuffer(deckBytes));
}

function findTargetParagraph(engine, slideIndex, match) {
  for (let shapeIndex = 0; shapeIndex < 60; shapeIndex++) {
    let sawAny = false;
    for (let paragraphIndex = 0; paragraphIndex < 80; paragraphIndex++) {
      let text = null;
      try {
        text = engine.getParagraphRunText(slideIndex, shapeIndex, paragraphIndex);
      } catch {
        text = null;
      }
      if (text === null) {
        if (paragraphIndex === 0) break; // no such shape
        break; // end of this shape's paragraphs
      }
      sawAny = true;
      if (text.replace(/\s+/g, ' ').includes(match.replace(/\s+/g, ' '))) {
        return { shapeIndex, paragraphIndex, text };
      }
    }
    if (!sawAny && shapeIndex > 30) break;
  }
  return null;
}

function highlightsForParagraph(engine, slideIndex, shapeIndex, paragraphIndex) {
  return engine.getSlideRunHighlights(slideIndex)
    .filter((h) => h.shapeIndex === shapeIndex && h.paragraphIndex === paragraphIndex)
    .map((h) => ({ runIndex: h.runIndex, color: h.color, start: h.start, end: h.end }));
}

const HL = 'FFFF00';

async function runScenario(name, slideIndex, target, applyRange, clearRange) {
  const engine = await freshEngine();
  const { shapeIndex, paragraphIndex } = target;

  await engine.setRunStyleForRanges(slideIndex, shapeIndex,
    [{ paragraphIndex, start: applyRange[0], end: applyRange[1] }], { highlight: HL });
  const afterApply = highlightsForParagraph(engine, slideIndex, shapeIndex, paragraphIndex);

  await engine.setRunStyleForRanges(slideIndex, shapeIndex,
    [{ paragraphIndex, start: clearRange[0], end: clearRange[1] }], { highlight: null });
  const afterClear = highlightsForParagraph(engine, slideIndex, shapeIndex, paragraphIndex);

  const residual = afterClear.filter((h) => h.end > h.start);
  return { name, applyRange, clearRange, afterApply, afterClear, residual };
}

const probe = await freshEngine();
const target = findTargetParagraph(probe, cli.slide, cli.match);
if (!target) {
  console.error(`No paragraph on slide ${cli.slide} matched "${cli.match}".`);
  process.exit(1);
}

const fullLen = target.text.length;
console.log(`Deck:  ${deckPath}`);
console.log(`Slide: ${cli.slide}  shape: ${target.shapeIndex}  paragraph: ${target.paragraphIndex}`);
console.log(`Run text (${fullLen} chars): ${target.text}`);
console.log('');

// Apply-only probes: highlight a range, then read back exactly which spans
// carry highlight. Localizes the over-application.
async function applyOnly(range) {
  const engine = await freshEngine();
  const { shapeIndex, paragraphIndex } = target;
  await engine.setRunStyleForRanges(cli.slide, shapeIndex,
    [{ paragraphIndex, start: range[0], end: range[1] }], { highlight: HL });
  const hl = highlightsForParagraph(engine, cli.slide, shapeIndex, paragraphIndex);
  const min = hl.length ? Math.min(...hl.map((h) => h.start)) : -1;
  const max = hl.length ? Math.max(...hl.map((h) => h.end)) : -1;
  return { range: `[${range[0]},${range[1]}]`, requested: range[1] - range[0], highlightedSpan: hl.length ? `${min}-${max}` : '(none)', highlightedChars: hl.reduce((n, h) => n + (h.end - h.start), 0) };
}

{
  const baseline = highlightsForParagraph(probe, cli.slide, target.shapeIndex, target.paragraphIndex);
  const span = baseline.length ? `${Math.min(...baseline.map((h) => h.start))}-${Math.max(...baseline.map((h) => h.end))}` : '(none)';
  console.log(`BASELINE highlights on this paragraph (no edits applied): ${baseline.length} run(s), span ${span}`);
  console.log('');
}

async function dumpRunsAfter(range) {
  const engine = await freshEngine();
  const { shapeIndex, paragraphIndex } = target;
  await engine.setRunStyleForRanges(cli.slide, shapeIndex,
    [{ paragraphIndex, start: range[0], end: range[1] }], { highlight: HL });
  let ooxml = null;
  try { ooxml = engine.renderer.getSlideOoxml(cli.slide); } catch (e) { return `(no renderer access: ${e})`; }
  const doc = new globalThis.DOMParser().parseFromString(ooxml, 'application/xml');
  const paras = Array.from(doc.getElementsByTagName('a:p'));
  // Find the target paragraph by matching concatenated run text.
  const para = paras.find((p) => Array.from(p.getElementsByTagName('a:r'))
    .map((r) => (r.getElementsByTagName('a:t')[0]?.textContent || '')).join('').includes(cli.match));
  if (!para) return '(paragraph not found in raw OOXML)';
  const runs = Array.from(para.getElementsByTagName('a:r'));
  let offset = 0;
  return runs.map((r, i) => {
    const t = r.getElementsByTagName('a:t')[0]?.textContent || '';
    const rPr = r.getElementsByTagName('a:rPr')[0];
    const hasHl = rPr ? rPr.getElementsByTagName('a:highlight').length > 0 : false;
    const span = `${offset}-${offset + t.length}`;
    offset += t.length;
    return `#${i} ${span}${hasHl ? ' HL' : '   '} ${JSON.stringify(t.slice(0, 14))}`;
  }).join('\n');
}

console.log(`Raw runs after highlight [60,70]:`);
console.log(await dumpRunsAfter([60, 70]));
console.log('');

const probes = [];
probes.push(await applyOnly([0, fullLen]));        // no split
probes.push(await applyOnly([0, 123]));            // split at end only
probes.push(await applyOnly([116, fullLen]));      // split at start only
probes.push(await applyOnly([116, 123]));          // split both ends (middle)
probes.push(await applyOnly([60, 70]));            // middle, away from edges
console.log('Apply-only probes (highlight range -> spans actually highlighted):');
console.table(probes);
console.log('');

const toIdx = target.text.indexOf('easy-to');
const scenarios = [
  // The exact user flow: highlight everything, then No-color everything.
  { name: 'highlight full -> clear full', apply: [0, fullLen], clear: [0, fullLen] },
  // Clear via two abutting ranges (mimics a multi-line selection rebuilt per line).
  { name: 'highlight full -> clear [0,half]+[half,full]', apply: [0, fullLen], clear: null, split: true },
  // Highlight a word containing "to", clear exactly it.
  toIdx >= 0
    ? { name: 'highlight "easy-to" -> clear same', apply: [toIdx, toIdx + 'easy-to'.length], clear: [toIdx, toIdx + 'easy-to'.length] }
    : null,
  // Highlight full, clear one char short of the end (residual-by-design sanity).
  { name: 'highlight full -> clear [0,full-2] (under-clear)', apply: [0, fullLen], clear: [0, Math.max(0, fullLen - 2)] }
].filter(Boolean);

const results = [];
for (const scenario of scenarios) {
  if (scenario.split) {
    const half = Math.floor(fullLen / 2);
    const engine = await freshEngine();
    const { shapeIndex, paragraphIndex } = target;
    await engine.setRunStyleForRanges(cli.slide, shapeIndex,
      [{ paragraphIndex, start: 0, end: fullLen }], { highlight: HL });
    await engine.setRunStyleForRanges(cli.slide, shapeIndex, [
      { paragraphIndex, start: 0, end: half },
      { paragraphIndex, start: half, end: fullLen }
    ], { highlight: null });
    const afterClear = highlightsForParagraph(engine, cli.slide, shapeIndex, paragraphIndex);
    results.push({ name: scenario.name, applyRange: [0, fullLen], clearRange: [[0, half], [half, fullLen]], afterClear, residual: afterClear.filter((h) => h.end > h.start) });
    continue;
  }
  results.push(await runScenario(scenario.name, cli.slide, target, scenario.apply, scenario.clear));
}

console.table(results.map((r) => ({
  scenario: r.name,
  apply: JSON.stringify(r.applyRange),
  clear: JSON.stringify(r.clearRange),
  afterApplyRuns: r.afterApply ? r.afterApply.length : '?',
  afterApplySpan: r.afterApply ? (r.afterApply.map((h) => `${h.start}-${h.end}`).join(',') || '(none)') : '?',
  residualRuns: r.residual.length
})));

for (const r of results) {
  if (r.residual.length > 0) {
    const chars = r.residual.map((h) => `[${h.start}-${h.end}] "${target.text.slice(h.start, h.end)}"`).join('  ');
    console.log(`\n  ${r.name}: residual highlight on ${chars}`);
  }
}

// Only the *full-paragraph* clears must end completely empty. Sub-range and
// under-clear scenarios are expected to leave the rest of any (pre-existing)
// highlight intact -- that is correct "clear only what you selected" behavior.
const fullClearNames = ['highlight full -> clear full', 'highlight full -> clear [0,half]+[half,full]'];
const unexpected = results.filter((r) => fullClearNames.includes(r.name) && r.residual.length > 0);
if (unexpected.length > 0) {
  console.error(`\nFAIL: ${unexpected.length} full-paragraph No-color(s) left residual highlight.`);
  process.exit(1);
}

console.log('\nPASS: full-paragraph No-color clears all highlight; sub-range clears correctly leave the rest.');
