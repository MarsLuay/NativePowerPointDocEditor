import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { test } from 'node:test';
import { build } from 'esbuild';
import { buildZip } from 'pptx-svg';

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

async function loadImportModule() {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), 'npde-chart-import-'));
  const outfile = path.join(tempDir, 'chartDataImport.cjs');
  try {
    await build({
      entryPoints: [path.resolve('src/powerpoint/chartDataImport.ts')],
      bundle: true,
      format: 'cjs',
      outfile,
      platform: 'node',
    });
    return { module: require(outfile), tempDir };
  } catch (error) {
    await rm(tempDir, { recursive: true, force: true });
    throw error;
  }
}

test('chart data import parses chart CSV, excel layout, transpose, and xlsx', async () => {
  const { module, tempDir } = await loadImportModule();
  try {
    const {
      matrixToChartDraft,
      parseDelimitedChartMatrix,
      parseExcelChartMatrix,
    } = module;

    assert.deepEqual(parseDelimitedChartMatrix('Category,Samples\n1g,1.04\n2g,2.08\n'), [
      ['Category', 'Samples'],
      ['1g', '1.04'],
      ['2g', '2.08'],
    ]);

    const chartDraft = matrixToChartDraft([
      ['Category', 'Samples'],
      ['1g', '1.04'],
      ['2g', '2.08'],
    ]);
    assert.equal(chartDraft.categoryLabel, 'Category');
    assert.deepEqual(chartDraft.categories, ['1g', '2g']);
    assert.deepEqual(chartDraft.series[0]?.values, ['1.04', '2.08']);

    const excelDraft = matrixToChartDraft([
      ['', 'Samples'],
      ['1g', '687789'],
      ['2g', '2.08'],
    ]);
    assert.deepEqual(excelDraft.categories, ['1g', '2g']);
    assert.deepEqual(excelDraft.series[0]?.values, ['687789', '2.08']);

    const transposed = matrixToChartDraft([
      ['', '1g', '2g', '3g'],
      ['Samples', '1', '2', '3'],
    ]);
    assert.deepEqual(transposed.categories, ['1g', '2g', '3g']);
    assert.deepEqual(transposed.series[0]?.values, ['1', '2', '3']);

    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t></t></is></c>
      <c r="B1" t="inlineStr"><is><t>Samples</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>1g</t></is></c>
      <c r="B2"><v>10</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>2g</t></is></c>
      <c r="B3"><v>20</v></c>
    </row>
  </sheetData>
</worksheet>`;

    const buffer = await buildZip(new ArrayBuffer(0), new Map([
      ['[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>'],
      ['_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml', '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
      ['xl/worksheets/sheet1.xml', sheet],
    ]));

    const matrix = await parseExcelChartMatrix(buffer);
    const fromXlsx = matrixToChartDraft(matrix);
    assert.deepEqual(fromXlsx.categories, ['1g', '2g']);
    assert.deepEqual(fromXlsx.series[0]?.values, ['10', '20']);
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
});
