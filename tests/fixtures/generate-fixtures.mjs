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
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const LOREM_IPSUM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";

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

function createLoremIpsumPptx() {
  const slideTree = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
  const relationships = (entries) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;

  return buildStoredZip([
    docxEntry(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`,
    ),
    docxEntry(
      "_rels/.rels",
      relationships(`<Relationship Id="rId1" Type="${OFFICE_RELS_NS}/officeDocument" Target="ppt/presentation.xml"/>`),
    ),
    docxEntry(
      "ppt/presentation.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`,
    ),
    docxEntry(
      "ppt/_rels/presentation.xml.rels",
      relationships(`
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${OFFICE_RELS_NS}/slide" Target="slides/slide1.xml"/>`),
    ),
    docxEntry(
      "ppt/slideMasters/slideMaster1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld><p:spTree>${slideTree}</p:spTree></p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`,
    ),
    docxEntry(
      "ppt/slideMasters/_rels/slideMaster1.xml.rels",
      relationships(`
  <Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="${OFFICE_RELS_NS}/theme" Target="../theme/theme1.xml"/>`),
    ),
    docxEntry(
      "ppt/slideLayouts/slideLayout1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}" type="blank" preserve="1">
  <p:cSld><p:spTree>${slideTree}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
    ),
    docxEntry(
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
      relationships(`<Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`),
    ),
    docxEntry(
      "ppt/theme/theme1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${DRAWING_NS}" name="">
  <a:themeElements>
    <a:clrScheme name="">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>
      <a:accent1><a:srgbClr val="000000"/></a:accent1><a:accent2><a:srgbClr val="000000"/></a:accent2>
      <a:accent3><a:srgbClr val="000000"/></a:accent3><a:accent4><a:srgbClr val="000000"/></a:accent4>
      <a:accent5><a:srgbClr val="000000"/></a:accent5><a:accent6><a:srgbClr val="000000"/></a:accent6>
      <a:hlink><a:srgbClr val="000000"/></a:hlink><a:folHlink><a:srgbClr val="000000"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name=""><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name=""><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`,
    ),
    docxEntry(
      "ppt/slides/slide1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_RELS_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld><p:spTree>
    ${slideTree}
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="10058400" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="la" sz="2800"/><a:t>${LOREM_IPSUM}</a:t></a:r><a:endParaRPr lang="la"/></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
    ),
    docxEntry(
      "ppt/slides/_rels/slide1.xml.rels",
      relationships(`<Relationship Id="rId1" Type="${OFFICE_RELS_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`),
    ),
  ]);
}

const docxFixtures = new Map([
  ["table-cell-direct-24pt-font.docx", createTableCellFontSizeDocx()],
]);
const publishFixtures = new Map([
  ["lorem-ipsum.pptx", createLoremIpsumPptx()],
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
  for (const [name, bytes] of publishFixtures) {
    await writeFile(path.join(fixtureRoot, name), bytes);
  }
  console.log(`Wrote ${deckFixtures.size + docxFixtures.size + publishFixtures.size} deterministic Native PowerPoint fixtures.`);
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
  for (const [name, expected] of publishFixtures) {
    const actual = await readFile(path.join(fixtureRoot, name));
    assert.deepEqual(
      actual,
      Buffer.from(expected),
      `${name} does not match the deterministic generator. Run npm run test:update-fixtures.`,
    );
  }
  console.log(`Verified ${deckFixtures.size + docxFixtures.size + publishFixtures.size} deterministic Native PowerPoint fixtures.`);
}

if (mode === "--write") {
  await writeFixtures();
} else if (mode === "--check") {
  await checkFixtures();
} else {
  throw new Error(`Unknown fixture generator mode: ${mode}`);
}
