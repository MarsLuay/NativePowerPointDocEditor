import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { loadDocxTableCellFontSizePreserverModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(projectRoot, "tests/fixtures/docx/table-cell-direct-24pt-font.docx");

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function updateDocumentXml(buffer, transform) {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  zip.file("word/document.xml", transform(documentXml));
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

async function readDocumentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file("word/document.xml").async("string");
}

test("DOCX table font-size fixture keeps 24 pt direct cell-run formatting", async () => {
  const zip = await JSZip.loadAsync(await readFile(fixturePath));
  const documentXml = await zip.file("word/document.xml").async("string");
  const stylesXml = await zip.file("word/styles.xml").async("string");

  assert.match(stylesXml, /<w:docDefaults>[\s\S]*<w:sz w:val="22"\/>[\s\S]*<\/w:docDefaults>/);

  const cellCount = (documentXml.match(/<w:tc>/g) ?? []).length;
  assert.equal(cellCount, 20);

  const directCellFontSizes = documentXml.match(/<w:tc>[\s\S]*?<w:rPr>[\s\S]*?<w:sz w:val="48"\/>[\s\S]*?<w:szCs w:val="48"\/>[\s\S]*?<\/w:rPr>[\s\S]*?<\/w:tc>/g) ?? [];
  assert.equal(directCellFontSizes.length, cellCount);

  assert.doesNotMatch(documentXml, /<w:tc>[\s\S]*?<w:sz w:val="22"\/>[\s\S]*?<\/w:tc>/);
});

test("preserveDocxTableCellFontSizes restores lost direct table-cell font sizes", async () => {
  const { preserveDocxTableCellFontSizes } = await loadDocxTableCellFontSizePreserverModule();
  const source = toArrayBuffer(await readFile(fixturePath));
  const damaged = await updateDocumentXml(source, (documentXml) =>
    documentXml.replace('<w:sz w:val="48"/>\n    <w:szCs w:val="48"/>', ""),
  );

  assert.equal((await readDocumentXml(damaged)).match(/<w:sz w:val="48"\/>/g).length, 19);

  const preserved = await preserveDocxTableCellFontSizes(source, damaged);
  const repairedXml = await readDocumentXml(preserved.buffer);

  assert.equal(preserved.restoredRuns, 1);
  assert.equal(preserved.restoredTags, 2);
  assert.equal(preserved.status, "restored");
  assert.equal(preserved.sourceCellCount, 20);
  assert.equal(preserved.outputCellCount, 20);
  assert.equal(preserved.matchedCellCount, 20);
  assert.equal(preserved.sourceRunsWithDirectSize, 20);
  assert.equal(repairedXml.match(/<w:sz w:val="48"\/>/g).length, 20);
  assert.equal(repairedXml.match(/<w:szCs w:val="48"\/>/g).length, 20);
});

test("preserveDocxTableCellFontSizes leaves explicit new table-cell font sizes alone", async () => {
  const { preserveDocxTableCellFontSizes } = await loadDocxTableCellFontSizePreserverModule();
  const source = toArrayBuffer(await readFile(fixturePath));
  const explicitlyChanged = await updateDocumentXml(source, (documentXml) =>
    documentXml
      .replace('<w:sz w:val="48"/>', '<w:sz w:val="22"/>')
      .replace('<w:szCs w:val="48"/>', '<w:szCs w:val="22"/>'),
  );

  const preserved = await preserveDocxTableCellFontSizes(source, explicitlyChanged);
  const outputXml = await readDocumentXml(preserved.buffer);

  assert.equal(preserved.restoredRuns, 0);
  assert.equal(preserved.restoredTags, 0);
  assert.equal(preserved.status, "checked");
  assert.equal(preserved.matchedCellCount, 20);
  assert.match(outputXml, /<w:tc>[\s\S]*?<w:sz w:val="22"\/>[\s\S]*?<w:szCs w:val="22"\/>[\s\S]*?<\/w:tc>/);
  assert.equal(outputXml.match(/<w:sz w:val="48"\/>/g).length, 19);
});
