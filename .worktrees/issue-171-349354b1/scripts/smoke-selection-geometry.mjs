// Headless reproduction harness for inline text *selection geometry*.
//
// Why this exists: the "selection stops short of a soft-wrapped line end"
// class of bug lives in real SVG glyph hit-testing (getExtentOfChar /
// getNumberOfChars / getScreenCTM). jsdom stubs those to zero, so a plain Node
// test cannot reproduce it. This harness renders the *real* engine SVG for the
// known-problematic bulleted paragraph, mounts it in the same canvas-pane DOM
// the plugin uses, bundles the *real* InlineTextGeometry module, and drives it
// with synthetic pointer coordinates inside headless Chrome — which computes
// genuine text layout. That lets us iterate on the geometry bug without
// dragging in Obsidian by hand.
//
// Usage: node scripts/smoke-selection-geometry.mjs

import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { buildZip, extractZip } from 'pptx-svg';
import { readDeck, toArrayBuffer } from '../tests/helpers/renderer.mjs';
import { loadPresentationEngineModule } from '../tests/helpers/load-plugin-modules.mjs';
import { runHeadlessHarness } from './lib/text-offset-harness.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const outputDir = path.resolve('scripts/visual-output');
const htmlPath = path.join(outputDir, 'selection-geometry.html');

// CLI: point the harness at a real deck instead of the synthetic fixture.
//   node scripts/smoke-selection-geometry.mjs --deck=<path> --slide=<n> --match="<substr>"
// --deck   absolute/relative path to a .pptx (omit to use the bundled fixture)
// --slide  0-based slide index (default 0)
// --match  substring that identifies the target paragraph's run text
//          (default: the SmartThings fixture sentence)
function parseArgs(argv) {
  const args = { deck: null, slide: 0, match: null };
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
const DEFAULT_MATCH = 'SmartThings-compatible appliances';
const matchSubstring = cli.match ?? DEFAULT_MATCH;

// Long, multi-run paragraph that soft-wraps across several visual lines. Mirrors
// the fixture used by tests/paragraph-visual-lines.test.mjs so the rendered DOM
// structure (bullet container + run tspans) matches what shipped offsets assume.
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
  assert.ok(slideXml.includes(fixtureTitleParagraph), 'fixture title paragraph not present in slide1.xml');
  const patched = await buildZip(
    source,
    new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, bulletParagraph)]])
  );
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(patched);
  return engine.renderSlide(0).svg;
}

async function renderRealDeckSvg(deckPath, slideIndex) {
  const resolved = path.resolve(process.cwd(), deckPath);
  assert.ok(existsSync(resolved), `deck not found: ${resolved}`);
  const bytes = await readFile(resolved);
  const source = toArrayBuffer(bytes);
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(source);
  const slideCount = typeof engine.getSlideCount === 'function' ? engine.getSlideCount() : null;
  if (slideCount !== null) {
    assert.ok(slideIndex < slideCount, `slide ${slideIndex} out of range (deck has ${slideCount})`);
  }
  return engine.renderSlide(slideIndex).svg;
}

async function bundleGeometryModule() {
  const result = await build({
    entryPoints: [path.join(projectRoot, 'src/powerpoint/inlineTextGeometry.ts')],
    bundle: true,
    format: 'iife',
    globalName: 'InlineGeometryNS',
    platform: 'browser',
    target: 'es2020',
    write: false,
    logLevel: 'silent'
  });
  return result.outputFiles[0].text;
}

const svg = cli.deck
  ? await renderRealDeckSvg(cli.deck, cli.slide)
  : await renderFixtureSvg();
const geometryBundle = await bundleGeometryModule();
const sourceLabel = cli.deck
  ? `${path.resolve(process.cwd(), cli.deck)} (slide ${cli.slide})`
  : 'synthetic fixture (features.pptx, bulleted Samsung paragraph)';

const driverScript = String.raw`
(function () {
  const pane = document.querySelector('.native-powerpoint-canvas-pane');
  const svg = pane.querySelector('svg');
  const geometry = new InlineGeometryNS.InlineTextGeometry(() => pane);

  function collectParagraphRuns(paraContainer) {
    const runs = paraContainer.matches('tspan[data-ooxml-para-idx]')
      ? Array.from(paraContainer.querySelectorAll(':scope > tspan[data-ooxml-run-idx]'))
      : Array.from(paraContainer.querySelectorAll('tspan[data-ooxml-run-idx]'));
    return runs;
  }

  function paragraphLineContainers(textEl) {
    // Direct-child line containers carrying a data-ooxml-para-idx, grouped per
    // paragraph index (one paragraph soft-wraps into several such containers).
    return Array.from(textEl.children).filter(
      (child) => child.tagName === 'tspan' && child.hasAttribute('data-ooxml-para-idx')
    );
  }

  const wanted = ${JSON.stringify(matchSubstring)};
  const normalized = (s) => (s || '').replace(/\s+/g, ' ');

  // Locate the <text> + paragraph index whose run text contains the match.
  let targetText = null;
  let targetParaIdx = null;
  for (const textEl of Array.from(svg.querySelectorAll('text'))) {
    const byPara = new Map();
    for (const lineContainer of paragraphLineContainers(textEl)) {
      const idx = lineContainer.getAttribute('data-ooxml-para-idx');
      const runText = collectParagraphRuns(lineContainer).map((r) => r.textContent || '').join('');
      byPara.set(idx, (byPara.get(idx) || '') + runText);
    }
    for (const [idx, runText] of byPara) {
      if (normalized(runText).includes(normalized(wanted))) {
        targetText = textEl;
        targetParaIdx = idx;
        break;
      }
    }
    if (targetText) break;
  }

  const result = { ok: false, error: null, containers: [], match: wanted, paraIdx: targetParaIdx };
  if (!targetText) {
    result.error = 'no paragraph run text matched "' + wanted + '" in rendered SVG';
    document.body.dataset.metrics = encodeURIComponent(JSON.stringify(result));
    return;
  }

  const lineContainers = paragraphLineContainers(targetText).filter(
    (c) => c.getAttribute('data-ooxml-para-idx') === targetParaIdx
  );
  const runContainers = lineContainers.filter((c) => collectParagraphRuns(c).length > 0);
  result.runOnlyText = runContainers
    .map((c) => collectParagraphRuns(c).map((r) => r.textContent || '').join(''))
    .join('');

  const paneRect = pane.getBoundingClientRect();
  // Replicates the per-container slice of getInlineTextOffsetAtClientPoint:
  //   geometryLocal -> runLocal -> snapWrappedRunLocalToLineEnd
  for (let i = 0; i < runContainers.length; i++) {
    const container = runContainers[i];
    const box = geometry.getElementBox(container);
    const geometryTotal = geometry.getLeafCharInfo(container).total;
    const runTotal = geometry.getRunCharInfo(container).total;
    let rawNumberOfChars = -1;
    try { rawNumberOfChars = container.getNumberOfChars(); } catch (e) { rawNumberOfChars = -2; }
    const leafSpanCount = geometry.getLeafTextSpans(container).length;
    if (!box) {
      result.containers.push({ index: i, error: 'no box' });
      continue;
    }

    // Simulate a drag that lands well past the right edge of this visual line —
    // i.e. the user trying to select to the end of the wrapped line.
    const clientX = paneRect.left + box.left + box.width + 24 - pane.scrollLeft;
    const clientY = paneRect.top + box.top + box.height / 2 - pane.scrollTop;
    const localClientX = clientX - paneRect.left + pane.scrollLeft;

    const geometryLocal = Math.max(0, Math.min(
      geometryTotal,
      geometry.getInlineTextOffsetAtClientPointForElement(container, clientX, clientY, box)
    ));
    const runLocalBeforeSnap = geometry.geometryIndexToRunOffset(container, geometryLocal);
    const runLocalAfterSnap = geometry.snapWrappedRunLocalToLineEnd(container, runLocalBeforeSnap, runTotal, localClientX);

    result.containers.push({
      index: i,
      text: container.textContent || '',
      geometryTotal,
      runTotal,
      rawNumberOfChars,
      leafSpanCount,
      undercounts: rawNumberOfChars >= 0 && rawNumberOfChars < geometryTotal,
      geometryLocal,
      runLocalBeforeSnap,
      runLocalAfterSnap,
      reachedEnd: runLocalAfterSnap === runTotal
    });
  }

  result.ok = runContainers.length > 0 && result.containers.every((c) => c.reachedEnd === true);
  document.body.dataset.metrics = encodeURIComponent(JSON.stringify(result));
})();
`;

const html = String.raw`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Native PowerPoint Selection Geometry Smoke</title>
  <style>
    body { background: #f4f6fb; margin: 0; padding: 24px; font-family: Arial, sans-serif; }
    .native-powerpoint-canvas-pane {
      background: #d9dee9;
      display: inline-block;
      overflow: auto;
      padding: 18px;
      position: relative;
    }
    .native-powerpoint-slide-surface {
      background: white;
      position: relative;
    }
    svg { display: block; }
    svg text, svg tspan { -webkit-user-select: none; user-select: none; }
  </style>
</head>
<body>
  <main>
    <div class="native-powerpoint-canvas-pane">
      <div class="native-powerpoint-slide-surface">${svg}</div>
    </div>
  </main>
  <script>${geometryBundle}</script>
  <script>
    window.addEventListener('load', () => {
      try {
        ${driverScript}
      } catch (error) {
        document.body.dataset.metrics = encodeURIComponent(JSON.stringify({ ok: false, error: String(error && error.stack || error) }));
      }
    });
  </script>
</body>
</html>`;

await mkdir(outputDir, { recursive: true });
await writeFile(htmlPath, html, 'utf8');

const harness = await runHeadlessHarness(htmlPath, { chromeTimeoutMs: 25000 });
assert.ok(harness, 'no headless runtime available');
const { metrics, runtime: runtimeLabel } = harness;

if (metrics.error) {
  console.error('Harness error:', metrics.error);
  process.exit(1);
}

console.log(`Runtime: ${runtimeLabel}`);
console.log(`Source:  ${sourceLabel}`);
console.log(`HTML:    ${htmlPath}`);
console.log(`Paragraph (para-idx ${metrics.paraIdx}) run text: ${metrics.runOnlyText}`);
console.table(metrics.containers.map((c) => ({
  line: c.index,
  leafSpans: c.leafSpanCount,
  rawChars: c.rawNumberOfChars,
  leafTotal: c.geometryTotal,
  undercount: c.undercounts,
  runTotal: c.runTotal,
  beforeSnap: c.runLocalBeforeSnap,
  afterSnap: c.runLocalAfterSnap,
  reachedEnd: c.reachedEnd,
  text: (c.text || '').slice(0, 26)
})));

assert.ok(metrics.containers.length > 0, 'expected at least one run-bearing visual line');
for (const container of metrics.containers) {
  // Independent coverage of both fixes:
  //  - beforeSnap exercises the leaf-total fix in getInlineTextOffsetFromSvgGeometry
  //    (getNumberOfChars() under-counted wrapped lines and capped the offset).
  //  - afterSnap exercises snapWrappedRunLocalToLineEnd (the line-end rescue).
  // Asserting both means a regression in either path turns this red.
  assert.equal(
    container.runLocalBeforeSnap,
    container.runTotal,
    `line ${container.index} ("${(container.text || '').slice(-24)}") geometry path capped at ${container.runLocalBeforeSnap}/${container.runTotal} — getInlineTextOffsetFromSvgGeometry under-counted the wrapped line`
  );
  assert.equal(
    container.reachedEnd,
    true,
    `line ${container.index} ("${(container.text || '').slice(-24)}") capped at ${container.runLocalAfterSnap}/${container.runTotal} — selection cannot reach the wrapped line end`
  );
}

console.log('\nSelection geometry smoke passed: drag to line end reaches the full run on every wrapped line.');
