import assert from "node:assert/strict";
import { test } from "node:test";
import { extractZip } from "pptx-svg";
import {
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const FIXTURE = "features.pptx";
const SLIDE_INDEX = 0;

const CHART_TYPES = [
  { id: "column", fingerprint: /<c:barChart\b/, barDir: /<c:barDir val="col"\/>/ },
  { id: "line", fingerprint: /<c:lineChart\b/ },
  { id: "pie", fingerprint: /<c:pieChart\b/ },
  { id: "bar", fingerprint: /<c:barChart\b/, barDir: /<c:barDir val="bar"\/>/ },
  { id: "area", fingerprint: /<c:areaChart\b/ },
  { id: "scatter", fingerprint: /<c:scatterChart\b/ },
  { id: "stock", fingerprint: /<c:stockChart\b/ },
  { id: "surface", fingerprint: /<c:surfaceChart\b/ },
  { id: "radar", fingerprint: /<c:radarChart\b/ },
  { id: "treemap", fingerprint: /layoutId="treemap"/, frame: /<cx:chart\b/ },
  { id: "sunburst", fingerprint: /layoutId="sunburst"/, frame: /<cx:chart\b/ },
  { id: "histogram", fingerprint: /layoutId="histogram"/, frame: /<cx:chart\b/ },
  { id: "boxWhisker", fingerprint: /layoutId="boxWhisker"/, frame: /<cx:chart\b/ },
  { id: "waterfall", fingerprint: /layoutId="waterfall"/, frame: /<cx:chart\b/ },
  { id: "combo", fingerprint: /<c:barChart\b[\s\S]*<c:lineChart\b/ },
];

async function loadFreshEngine() {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const bytes = await readDeck(FIXTURE);
  return PresentationEngine.load(toArrayBuffer(bytes));
}

test("chart insert catalog covers every PowerPoint-like type without stubs", async () => {
  const { build } = await import("esbuild");
  const { mkdtemp } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const path = await import("node:path");
  const { fileURLToPath, pathToFileURL } = await import("node:url");
  const { createRequire } = await import("node:module");
  const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const require = createRequire(import.meta.url);
  const { DOMParser } = require("@xmldom/xmldom");
  globalThis.DOMParser ??= DOMParser;

  const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-chart-types-"));
  const outfile = path.join(outputDirectory, "chart-types.cjs");
  await build({
    stdin: {
      contents: [
        "export { INSERTABLE_CHART_TYPES, CHART_TEMPLATE_ENTRIES, buildChartInsertParts } from './src/powerpoint/chartInsertTypes.ts';",
      ].join("\n"),
      resolveDir: projectRoot,
      sourcefile: "chart-types-test-entry.ts",
    },
    bundle: true,
    format: "cjs",
    platform: "node",
    outfile,
    logLevel: "silent",
  });
  const mod = require(outfile);
  assert.equal(mod.INSERTABLE_CHART_TYPES.length, CHART_TYPES.length);
  assert.deepEqual(
    mod.INSERTABLE_CHART_TYPES.map((entry) => entry.id),
    CHART_TYPES.map((entry) => entry.id),
  );
  assert.equal(mod.CHART_TEMPLATE_ENTRIES.length, 3);
  for (const entry of mod.INSERTABLE_CHART_TYPES) {
    const parts = mod.buildChartInsertParts(entry.id);
    assert.ok(parts.chartXml.length > 100, `${entry.id} must emit real chart XML`);
    assert.ok(parts.frameXml.includes("graphicFrame"), `${entry.id} must emit a graphic frame`);
    assert.match(parts.chartXml, entry.kind === "chartex" ? /cx:chartSpace/ : /c:chartSpace/);
  }
});

for (const chart of CHART_TYPES) {
  test(`insert ${chart.id} chart renders and round-trips`, async () => {
    const engine = await loadFreshEngine();
    const shapeIndex = await engine.addChart(SLIDE_INDEX, chart.id);
    assert.equal(Number.isInteger(shapeIndex) && shapeIndex >= 0, true);

    const svg = engine.renderSlide(SLIDE_INDEX).svg;
    assert.match(svg, /^<svg\b/);
    assert.match(
      svg,
      /data-ooxml-shape-type="chart"/,
      `${chart.id}: rendered SVG must expose a chart shape`,
    );

    const exported = await engine.export();
    const zip = await extractZip(exported);
    const slideXml = zip.textFiles.get("ppt/slides/slide1.xml") ?? "";
    assert.match(slideXml, chart.frame ?? /<c:chart\b/, `${chart.id}: slide must reference chart`);

    const chartParts = [...zip.textFiles.keys()].filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p));
    assert.ok(chartParts.length >= 1, `${chart.id}: must write a chart part`);
    const chartXml = zip.textFiles.get(chartParts[chartParts.length - 1]) ?? "";
    assert.match(chartXml, chart.fingerprint, `${chart.id}: chart XML fingerprint`);
    if (chart.barDir) {
      assert.match(chartXml, chart.barDir, `${chart.id}: barDir`);
    }
  });
}
