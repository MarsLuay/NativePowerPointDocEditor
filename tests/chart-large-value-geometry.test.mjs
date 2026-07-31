import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { loadPresentationEngineModule } from './helpers/load-plugin-modules.mjs';
import { toArrayBuffer } from './helpers/renderer.mjs';

const require = createRequire(import.meta.url);
const { DOMParser } = require('@xmldom/xmldom');

function chartBars(root, shapeIndex) {
  const groups = Array.from(root.getElementsByTagName('g'));
  const chart = groups.find(
    (node) =>
      node.getAttribute('data-ooxml-shape-type') === 'chart' &&
      node.getAttribute('data-ooxml-shape-idx') === String(shapeIndex)
  );
  assert.ok(chart, 'chart group missing');
  return Array.from(chart.getElementsByTagName('rect'))
    .filter((rect) => rect.getAttribute('data-native-powerpoint-chart-bar') === 'true')
    .map((rect) => ({
      x: Number(rect.getAttribute('x')),
      y: Number(rect.getAttribute('y')),
      width: Number(rect.getAttribute('width')),
      height: Number(rect.getAttribute('height')),
    }));
}

function axisTickTexts(root) {
  return Array.from(root.getElementsByTagName('text'))
    .filter((node) => node.getAttribute('data-native-powerpoint-axis-tick') === 'true')
    .map((node) => node.textContent);
}

test('column chart preview scales 687789 after pptx-svg i32 overflow', async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const bytes = await readFile(path.join(process.cwd(), 'tests/fixtures/decks/features.pptx'));
  const engine = await PresentationEngine.load(toArrayBuffer(bytes));
  const shapeIndex = await engine.addChart(0, 'column');
  const before = engine.getChartDataGrid(0, shapeIndex);
  assert.ok(before);

  await engine.updateChartData(0, shapeIndex, {
    categories: before.categories,
    series: [{ values: ['687789', '500000', '400000'], pointLabels: null }],
  });

  const rendered = engine.renderSlide(0).svg;
  const document = new DOMParser().parseFromString(rendered, 'image/svg+xml');
  const root = document.documentElement;
  engine.formatChartAxisLabels(root, 0);

  const bars = chartBars(root, shapeIndex);
  assert.equal(bars.length, 3);

  const heights = bars.map((bar) => bar.height).sort((a, b) => a - b);
  assert.ok(heights[0] > 50, `smallest bar too short: ${heights[0]}`);
  assert.ok(heights[2] > heights[0], 'bars should reflect value ranking');
  assert.ok(Math.abs(heights[2] / heights[0] - 687789 / 400000) < 0.15, 'bar height ratio off');
  assert.ok(bars.every((bar) => bar.height > bar.width), 'column bars should be vertical');

  const ticks = axisTickTexts(root);
  assert.ok(ticks.some((tick) => tick.includes('687') || tick.includes('700') || tick === '687789'));
});

test('bar chart preview keeps horizontal bars for 689999', async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const bytes = await readFile(path.join(process.cwd(), 'tests/fixtures/decks/features.pptx'));
  const engine = await PresentationEngine.load(toArrayBuffer(bytes));
  const shapeIndex = await engine.addChart(0, 'bar');
  const before = engine.getChartDataGrid(0, shapeIndex);
  assert.ok(before);

  await engine.updateChartData(0, shapeIndex, {
    categories: before.categories,
    series: [{ values: ['689999', '500000', '400000'], pointLabels: null }],
  });

  const rendered = engine.renderSlide(0).svg;
  const document = new DOMParser().parseFromString(rendered, 'image/svg+xml');
  const root = document.documentElement;
  engine.formatChartAxisLabels(root, 0);

  const bars = chartBars(root, shapeIndex);
  assert.equal(bars.length, 3);
  assert.ok(bars.every((bar) => bar.width > bar.height), 'bar chart bars should stay horizontal');
  const widths = bars.map((bar) => bar.width).sort((a, b) => a - b);
  assert.ok(widths[2] > widths[0], 'bar widths should reflect value ranking');
});
