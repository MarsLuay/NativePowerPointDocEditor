import assert from "node:assert/strict";
import { Buffer } from "node:buffer";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildStoredZip, createDeck, createDeckEntries } from "../helpers/fixture-builder.mjs";

const fixtureRoot = path.dirname(fileURLToPath(import.meta.url));
const deckFixtureDirectory = path.join(fixtureRoot, "decks");
const docxFixtureDirectory = path.join(fixtureRoot, "docx");
const mode = process.argv[2] ?? "--check";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const deckFixtures = new Map([
  ["features.pptx", createDeck({ format: "pptx" })],
  ["features.ppsx", createDeck({ format: "ppsx" })],
  ["features.potx", createDeck({ format: "potx" })],
  ["simple-edit.pptx", createDeck({ format: "pptx", richFirstSlide: false })],
  ["macro-view-only.pptm", createDeck({ format: "pptm" })],
  ["macro-view-only.ppsm", createDeck({ format: "ppsm" })],
  ["macro-view-only.potm", createDeck({ format: "potm" })],
  ["large-deck.pptx", createDeck({ format: "pptx", slideCount: 160 })],
]);

const featureEntries = createDeckEntries({ format: "pptx" });
const featureDeck = deckFixtures.get("features.pptx");
deckFixtures.set("malformed-random.pptx", new TextEncoder().encode("This is intentionally not a ZIP archive.\n"));
deckFixtures.set("malformed-truncated.pptx", featureDeck.slice(0, Math.floor(featureDeck.byteLength / 2)));
deckFixtures.set(
  "malformed-unsafe-path.pptx",
  buildStoredZip([...featureEntries, { name: "../escape.xml", data: new TextEncoder().encode("<escape/>") }]),
);
deckFixtures.set(
  "malformed-duplicate-entry.pptx",
  buildStoredZip([
    ...featureEntries,
    {
      name: "ppt/presentation.xml",
      data: new TextEncoder().encode('<?xml version="1.0"?><duplicate/>'),
    },
  ]),
);

function xml(strings, ...values) {
  return strings.reduce((result, part, index) => result + part + (values[index] ?? ""), "");
}

function docxEntry(name, data) {
  return { name, data };
}

function largeCellRun(text) {
  return xml`<w:r>
  <w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>
    <w:sz w:val="48"/>
    <w:szCs w:val="48"/>
  </w:rPr>
  <w:t>${text}</w:t>
</w:r>`;
}

function tableCell(text, fill = "FFFFFF") {
  return xml`<w:tc>
  <w:tcPr>
    <w:tcW w:w="2340" w:type="dxa"/>
    <w:shd w:fill="${fill}"/>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    ${largeCellRun(text)}
  </w:p>
</w:tc>`;
}

function tableRow(cells, fill) {
  return `<w:tr>${cells.map((cell) => tableCell(cell, fill)).join("")}</w:tr>`;
}

function createTableCellFontSizeDocx() {
  const contentTypes = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  const rootRels = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdOfficeDocument" Type="${OFFICE_RELS_NS}/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rIdCoreProperties" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rIdAppProperties" Type="${OFFICE_RELS_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  const documentRels = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">
  <Relationship Id="rIdStyles" Type="${OFFICE_RELS_NS}/styles" Target="styles.xml"/>
</Relationships>`;
  const styles = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${WORD_NS}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>
</w:styles>`;
  const document = xml`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
        <w:t>DOCX table-cell font-size regression fixture</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Document defaults are intentionally 11 pt. Every table-cell run below is directly formatted at 24 pt, so any cell that falls back to the default should be obvious.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="9360" w:type="dxa"/>
        <w:tblLayout w:type="fixed"/>
        <w:tblCellMar>
          <w:top w:w="120" w:type="dxa"/>
          <w:start w:w="120" w:type="dxa"/>
          <w:bottom w:w="120" w:type="dxa"/>
          <w:end w:w="120" w:type="dxa"/>
        </w:tblCellMar>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="999999"/>
          <w:left w:val="single" w:sz="8" w:color="999999"/>
          <w:bottom w:val="single" w:sz="8" w:color="999999"/>
          <w:right w:val="single" w:sz="8" w:color="999999"/>
          <w:insideH w:val="single" w:sz="8" w:color="999999"/>
          <w:insideV w:val="single" w:sz="8" w:color="999999"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
        <w:gridCol w:w="2340"/>
      </w:tblGrid>
      ${tableRow(["A1", "A2", "A3", "A4"], "F2F4F7")}
      ${tableRow(["B1", "B2", "B3", "B4"])}
      ${tableRow(["Size", "Keep", "Wide", "Cell"], "F9FAFB")}
      ${tableRow(["C1", "C2", "C3", "C4"])}
      ${tableRow(["D1", "D2", "D3", "D4"], "F9FAFB")}
    </w:tbl>
    <w:p>
      <w:pPr><w:spacing w:before="160"/></w:pPr>
      <w:r><w:t>Reference default text: this paragraph should remain 11 pt.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  return buildStoredZip([
    docxEntry("[Content_Types].xml", contentTypes),
    docxEntry("_rels/.rels", rootRels),
    docxEntry("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>DOCX table-cell font-size regression fixture</dc:title></cp:coreProperties>'),
    docxEntry("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Native PowerPoint Doc Editor tests</Application></Properties>'),
    docxEntry("word/_rels/document.xml.rels", documentRels),
    docxEntry("word/styles.xml", styles),
    docxEntry("word/document.xml", document),
  ]);
}

const docxFixtures = new Map([
  ["table-cell-direct-24pt-font.docx", createTableCellFontSizeDocx()],
]);

async function writeFixtures() {
  await mkdir(deckFixtureDirectory, { recursive: true });
  for (const [name, bytes] of deckFixtures) {
    await writeFile(path.join(deckFixtureDirectory, name), bytes);
  }
  await mkdir(docxFixtureDirectory, { recursive: true });
  for (const [name, bytes] of docxFixtures) {
    await writeFile(path.join(docxFixtureDirectory, name), bytes);
  }
  console.log(`Wrote ${deckFixtures.size + docxFixtures.size} deterministic Native PowerPoint fixtures.`);
}

async function checkFixtures() {
  for (const [name, expected] of deckFixtures) {
    const actual = await readFile(path.join(deckFixtureDirectory, name));
    assert.deepEqual(
      actual,
      Buffer.from(expected),
      `${name} does not match the deterministic generator. Run npm run test:update-fixtures.`,
    );
  }
  for (const [name, expected] of docxFixtures) {
    const actual = await readFile(path.join(docxFixtureDirectory, name));
    assert.deepEqual(
      actual,
      Buffer.from(expected),
      `${name} does not match the deterministic generator. Run npm run test:update-fixtures.`,
    );
  }
  console.log(`Verified ${deckFixtures.size + docxFixtures.size} deterministic Native PowerPoint fixtures.`);
}

if (mode === "--write") {
  await writeFixtures();
} else if (mode === "--check") {
  await checkFixtures();
} else {
  throw new Error(`Unknown fixture generator mode: ${mode}`);
}
