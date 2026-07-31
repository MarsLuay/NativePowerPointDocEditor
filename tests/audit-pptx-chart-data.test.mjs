import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { after, before, test } from "node:test";
import { build } from "esbuild";

// Audit coverage for "PPTX - Tables & charts" chart-data editing
// (src/ChartData.ts: getChartDataDescriptor / updateChartData). ChartData
// operates on raw OOXML package buffers via pptx-svg's extractZip/buildZip, so
// these functions can be exercised fully headless without the wasm renderer.
//
// There is no editable embedded-workbook chart in tests/fixtures/decks/ (the
// only chart fixture, features.pptx, uses literal strLit/numLit data with no
// embedded workbook), so the editable round-trip is verified against a
// self-contained synthetic deck built here, while the real fixture verifies the
// documented read-only-reason path. See the "fixture gap" note at the bottom.

const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

// pptx-svg's index pulls in pptx-renderer.js, whose module scope resolves a
// default wasm URL via import.meta.url. ChartData never instantiates the
// renderer, so we neutralize that URL (mirrors the project's smoke bundlers).
const inlinePptxSvgWasmPlugin = {
  name: "inline-pptx-svg-wasm",
  setup(buildContext) {
    buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async (args) => {
      const source = await readFile(args.path, "utf8");
      return {
        contents: source.replace(
          "const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;",
          "const DEFAULT_WASM_URL = undefined;",
        ),
        loader: "js",
      };
    });
  },
};

let tempDir;
let chartData;
let zipHelpers;

async function bundle(options, outputName) {
  const outfile = path.join(tempDir, outputName);
  await build({
    bundle: true,
    format: "cjs",
    loader: { ".wasm": "binary" },
    logLevel: "silent",
    outfile,
    platform: "node",
    plugins: [inlinePptxSvgWasmPlugin],
    target: "node22",
    ...options,
  });
  return require(outfile);
}

before(async () => {
  tempDir = await mkdtemp(path.join(os.tmpdir(), "native-powerpoint-chart-data-test-"));
  // pptx-svg ships ESM-only (no CJS require entry), so re-export the zip helpers
  // through an esbuild bundle rather than require()-ing the package directly.
  // resolveDir anchors the virtual entry to the project's node_modules.
  zipHelpers = await bundle(
    {
      stdin: {
        contents: "export { extractZip, buildZip } from 'pptx-svg';\n",
        resolveDir: projectRoot,
        sourcefile: "zip-entry.mjs",
        loader: "js",
      },
    },
    "zip-helpers.cjs",
  );
  chartData = await bundle(
    { entryPoints: [path.join(projectRoot, "src/ChartData.ts")] },
    "ChartData.cjs",
  );
});

after(async () => {
  if (tempDir) await rm(tempDir, { recursive: true, force: true });
});

function toArrayBuffer(view) {
  return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength);
}

// A minimal embedded Excel workbook (.xlsx) containing exactly the cells the
// synthetic chart's formulas point at. STORE compression keeps the bytes
// readable by pptx-svg's extractZip without an inflate round-trip.
async function buildEmbeddedWorkbook() {
  const zip = new JSZip();
  zip.file(
    "xl/workbook.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1"' +
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' +
      ' Target="worksheets/sheet1.xml"/></Relationships>',
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>20</v></c></row>' +
      "</sheetData></worksheet>",
  );
  return zip.generateAsync({ type: "uint8array" });
}

// A self-contained chart package: a bar chart with source-backed (numRef/strRef)
// ranges plus an embedded-workbook relationship, which is what makes the grid
// editable (categories present, a numeric series present, single category range).
async function buildEditableChartDeck() {
  const workbook = await buildEmbeddedWorkbook();
  const zip = new JSZip();
  zip.file(
    "ppt/charts/chart1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      "<c:chart><c:plotArea><c:barChart><c:barDir val=\"col\"/>" +
      "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
      "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>" +
      "<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>" +
      "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f>" +
      "<c:strCache><c:ptCount val=\"2\"/>" +
      "<c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
      "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f>" +
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"2\"/>" +
      "<c:pt idx=\"0\"><c:v>10</c:v></c:pt><c:pt idx=\"1\"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>" +
      "</c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart></c:plotArea></c:chart>" +
      '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
  );
  zip.file(
    "ppt/charts/_rels/chart1.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1"' +
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"' +
      ' Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>',
  );
  zip.file("ppt/embeddings/Microsoft_Excel_Worksheet.xlsx", workbook);
  return zip.generateAsync({ type: "arraybuffer" });
}

const CHART_PATH = "ppt/charts/chart1.xml";
const WORKBOOK_PATH = "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx";

test("getChartDataDescriptor exposes an editable grid for a source-backed embedded-workbook chart", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  assert.equal(descriptor.chartPath, CHART_PATH);
  assert.equal(descriptor.workbookPath, WORKBOOK_PATH);
  assert.equal(descriptor.grid.editable, true);
  assert.equal(descriptor.grid.reason, "");
  assert.deepEqual(descriptor.grid.categories, ["Q1", "Q2"]);
  assert.equal(descriptor.grid.series.length, 1);
  assert.equal(descriptor.grid.series[0].name, "Revenue");
  assert.deepEqual(descriptor.grid.series[0].values, ["10", "20"]);
});

test("updateChartData applies edits to chart caches and the embedded workbook, and round-trips", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  const update = {
    categories: ["Q1", "Q3"],
    series: [{ values: ["10", "99"], pointLabels: null }],
  };
  const output = await chartData.updateChartData(buffer, descriptor, update);

  // 1. The descriptor re-read from the exported deck reflects the change.
  const reloaded = chartData.getChartDataDescriptor(await extractZip(output), CHART_PATH);
  assert.deepEqual(reloaded.grid.categories, ["Q1", "Q3"]);
  assert.deepEqual(reloaded.grid.series[0].values, ["10", "99"]);
  assert.equal(reloaded.grid.editable, true);

  // 2. The chart cache XML was rewritten (workbook-sync: cache + workbook move
  // together, the bug class this group is most prone to).
  const reExtracted = await extractZip(output);
  const chartXml = reExtracted.textFiles.get(CHART_PATH);
  assert.match(chartXml, /<c:v>Q3<\/c:v>/);
  assert.match(chartXml, /<c:v>99<\/c:v>/);

  // 3. The embedded workbook cells were updated in lockstep and marked for
  // recalculation so the host app does not show stale formula results.
  const workbookBytes = reExtracted.binaryFiles.get(WORKBOOK_PATH);
  assert.ok(workbookBytes, "embedded workbook should survive the export");
  const workbookFiles = await extractZip(toArrayBuffer(workbookBytes));
  const sheetXml = workbookFiles.textFiles.get("xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A3" t="inlineStr"><is><t>Q3<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B3"><v>99<\/v><\/c>/);
  const workbookXml = workbookFiles.textFiles.get("xl/workbook.xml");
  assert.match(workbookXml, /fullCalcOnLoad="1"/);
  assert.match(workbookXml, /forceFullCalc="1"/);
});

test("updateChartData accepts grouped and overflow-scale numeric values", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  const output = await chartData.updateChartData(buffer, descriptor, {
    categories: descriptor.grid.categories,
    series: [{
      values: ["1,000,000", "1e400"],
      pointLabels: null,
    }],
  });

  const reloaded = chartData.getChartDataDescriptor(await extractZip(output), CHART_PATH);
  assert.equal(reloaded.grid.series[0].values[0], "1000000");
  assert.equal(reloaded.grid.series[0].values[1], String(Number.MAX_VALUE));

  const negative = await chartData.updateChartData(output, reloaded, {
    categories: reloaded.grid.categories,
    series: [{
      values: ["1000000", "-9e999"],
      pointLabels: null,
    }],
  });
  const negativeReloaded = chartData.getChartDataDescriptor(await extractZip(negative), CHART_PATH);
  assert.equal(negativeReloaded.grid.series[0].values[1], String(-Number.MAX_VALUE));

  await assert.rejects(
    () => chartData.updateChartData(buffer, descriptor, {
      categories: descriptor.grid.categories,
      series: [{ values: ["nope", "20"], pointLabels: null }],
    }),
    /Series 1 row 1 must be a number/,
  );
});

test("updateChartData inserts chart rows by extending caches, formulas, and the embedded workbook", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  const output = await chartData.updateChartData(buffer, descriptor, {
    categories: ["Q1", "Q2", "Q3"],
    rowOperation: { type: "insert", index: 2 },
    series: [{ values: ["10", "20", "30"], pointLabels: null }],
  });

  const reExtracted = await extractZip(output);
  const reloaded = chartData.getChartDataDescriptor(reExtracted, CHART_PATH);
  assert.deepEqual(reloaded.grid.categories, ["Q1", "Q2", "Q3"]);
  assert.deepEqual(reloaded.grid.series[0].values, ["10", "20", "30"]);

  const chartXml = reExtracted.textFiles.get(CHART_PATH);
  assert.match(chartXml, /Sheet1!\$A\$2:\$A\$4/);
  assert.match(chartXml, /Sheet1!\$B\$2:\$B\$4/);
  assert.match(chartXml, /<c:ptCount val="3"\/>/);

  const workbookBytes = reExtracted.binaryFiles.get(WORKBOOK_PATH);
  const workbookFiles = await extractZip(toArrayBuffer(workbookBytes));
  const sheetXml = workbookFiles.textFiles.get("xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A4" t="inlineStr"><is><t>Q3<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B4"><v>30<\/v><\/c>/);
});

test("updateChartData deletes chart rows by shrinking formulas and clearing stale workbook cells", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  const output = await chartData.updateChartData(buffer, descriptor, {
    categories: ["Q2"],
    rowOperation: { type: "delete", index: 0 },
    series: [{ values: ["20"], pointLabels: null }],
  });

  const reExtracted = await extractZip(output);
  const reloaded = chartData.getChartDataDescriptor(reExtracted, CHART_PATH);
  assert.deepEqual(reloaded.grid.categories, ["Q2"]);
  assert.deepEqual(reloaded.grid.series[0].values, ["20"]);

  const chartXml = reExtracted.textFiles.get(CHART_PATH);
  assert.match(chartXml, /Sheet1!\$A\$2<\/c:f>/);
  assert.match(chartXml, /Sheet1!\$B\$2<\/c:f>/);
  assert.match(chartXml, /<c:ptCount val="1"\/>/);

  const workbookBytes = reExtracted.binaryFiles.get(WORKBOOK_PATH);
  const workbookFiles = await extractZip(toArrayBuffer(workbookBytes));
  const sheetXml = workbookFiles.textFiles.get("xl/worksheets/sheet1.xml");
  assert.match(sheetXml, /<c r="A2" t="inlineStr"><is><t>Q2<\/t><\/is><\/c>/);
  assert.match(sheetXml, /<c r="B2"><v>20<\/v><\/c>/);
  assert.match(sheetXml, /<c r="A3"\/>/);
  assert.match(sheetXml, /<c r="B3"\/>/);
});

test("updateChartData validates row operations and keeps at least one chart row", async () => {
  const { extractZip } = zipHelpers;
  const buffer = await buildEditableChartDeck();
  const descriptor = chartData.getChartDataDescriptor(await extractZip(buffer), CHART_PATH);

  await assert.rejects(
    () => chartData.updateChartData(buffer, descriptor, {
      categories: ["Q1", "Q2", "Q3"],
      rowOperation: { type: "delete", index: 0 },
      series: [{ values: ["10", "20", "30"], pointLabels: null }],
    }),
    /Deleted chart rows do not match the updated chart grid/,
  );
  await assert.rejects(
    () => chartData.updateChartData(buffer, descriptor, {
      categories: [],
      rowOperation: { type: "delete", index: 0, count: 2 },
      series: [{ values: [], pointLabels: null }],
    }),
    /A chart must keep at least one category row/,
  );
});

test("getChartDataDescriptor returns a documented read-only reason for the literal-data fixture (features.pptx)", async () => {
  const { extractZip } = zipHelpers;
  const fixturePath = path.join(projectRoot, "tests/fixtures/decks/features.pptx");
  const fixtureBuffer = toArrayBuffer(await readFile(fixturePath));
  const descriptor = chartData.getChartDataDescriptor(await extractZip(fixtureBuffer), CHART_PATH);

  // features.pptx's chart uses strLit/numLit with no embedded workbook, so the
  // grid is intentionally read-only with a user-facing reason, and there is no
  // workbook path to edit.
  assert.equal(descriptor.grid.editable, false);
  assert.equal(descriptor.workbookPath, null);
  assert.equal(descriptor.grid.reason, "This chart has no embedded Excel workbook.");

  // The read-only contract is enforced in updateChartData too: it throws the
  // same reason rather than silently no-op'ing.
  await assert.rejects(
    () =>
      chartData.updateChartData(fixtureBuffer, descriptor, {
        categories: descriptor.grid.categories,
        series: [],
      }),
    /This chart has no embedded Excel workbook\./,
  );
});
