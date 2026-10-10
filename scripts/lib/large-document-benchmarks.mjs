import { performance } from "node:perf_hooks";
import JSZip from "jszip";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createDeck } from "../../tests/helpers/fixture-builder.mjs";
import {
  loadPresentationEngineModule,
  loadDocxTextExtractorModule,
  loadDocxTableCellFontSizePreserverModule,
} from "../../tests/helpers/load-plugin-modules.mjs";

const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");

globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

export const DEFAULT_DOCX_TIERS = [10, 50, 100];
export const OPT_IN_DOCX_TIERS = [250];

export const DEFAULT_PPTX_TIERS = [10, 50, 100];
export const OPT_IN_PPTX_TIERS = [250, 500];

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const LOREM = "The quick brown fox jumps over the lazy dog. Continuous document performance verification measures latency across realistic OOXML page and slide counts.";
const FIXTURE_DATE = new Date("2000-01-01T00:00:00.000Z");

function round(val) {
  return Math.round(val * 10) / 10;
}

export async function generateDeterministicDocx(pageCount) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
    { date: FIXTURE_DATE },
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    { date: FIXTURE_DATE },
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    { date: FIXTURE_DATE },
  );
  zip.file(
    "word/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${WORD_NS}">
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`,
    { date: FIXTURE_DATE },
  );

  let body = "";
  for (let page = 1; page <= pageCount; page++) {
    body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Page ${page} - Benchmark Section Heading</w:t></w:r></w:p>`;
    for (let p = 1; p <= 4; p++) {
      body += `<w:p><w:r><w:t>Page ${page}, Paragraph ${p}: ${LOREM}</w:t></w:r></w:p>`;
    }
    body += `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tr>
        <w:tc><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Page ${page} Col A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Page ${page} Col B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Value ${page}.1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Value ${page}.2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    if (page < pageCount) {
      body += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
    }
  }

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${WORD_NS}"><w:body>${body}</w:body></w:document>`,
    { date: FIXTURE_DATE },
  );

  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

export function generateDeterministicPptx(slideCount) {
  const deckBuffer = createDeck({ format: "pptx", slideCount, richFirstSlide: true });
  return deckBuffer.buffer.slice(deckBuffer.byteOffset, deckBuffer.byteOffset + deckBuffer.byteLength);
}

export async function runDocxBenchmarkTier(pageCount) {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const { preserveDocxTableCellFontSizes } = await loadDocxTableCellFontSizePreserverModule();

  const memStart = process.memoryUsage?.().heapUsed ?? 0;
  const docxBuffer = await generateDeterministicDocx(pageCount);
  const documentSizeBytes = docxBuffer.byteLength;

  // 1. Open to first visible
  const openFirstStart = performance.now();
  const zip = await JSZip.loadAsync(docxBuffer.slice(0));
  const docXml = await zip.file("word/document.xml").async("string");
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXml, "application/xml");
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, "p"));
  const firstPageVisibleText = paragraphs.slice(0, 5).map((p) => p.textContent).join(" ");
  const openToFirstVisibleMs = round(performance.now() - openFirstStart);

  // 2. Open to usable
  const openUsableStart = performance.now();
  const stylesXml = await zip.file("word/styles.xml")?.async("string");
  if (stylesXml) parser.parseFromString(stylesXml, "application/xml");
  const allTables = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, "tbl"));
  const totalParagraphCount = paragraphs.length;
  const totalTableCount = allTables.length;
  const openToUsableMs = round(openToFirstVisibleMs + (performance.now() - openUsableStart));

  // 3. Cold navigation to middle page
  const coldStart = performance.now();
  const midPage = Math.floor(pageCount / 2);
  const pageBreaks = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, "br")).filter(
    (br) => br.getAttributeNS(WORD_NS, "type") === "page" || br.getAttribute("w:type") === "page"
  );
  const midTarget = pageBreaks[Math.min(midPage, pageBreaks.length - 1)];
  const midParent = midTarget ? midTarget.parentNode : null;
  const midText = midParent ? midParent.textContent : "";
  const coldNavigationMs = round(performance.now() - coldStart);

  // 4. Warm navigation back to page 1
  const warmStart = performance.now();
  const warmFirstPara = paragraphs[0]?.textContent ?? "";
  const warmNavigationMs = round(performance.now() - warmStart);

  // 5. Typing mutation
  const typingStart = performance.now();
  const newP = xmlDoc.createElementNS(WORD_NS, "w:p");
  const newR = xmlDoc.createElementNS(WORD_NS, "w:r");
  const newT = xmlDoc.createElementNS(WORD_NS, "w:t");
  newT.textContent = `Typed mutation at tier ${pageCount}`;
  newR.appendChild(newT);
  newP.appendChild(newR);
  xmlDoc.documentElement.firstChild.appendChild(newP);
  await preserveDocxTableCellFontSizes(docxBuffer, docxBuffer);
  const typingMutationMs = round(performance.now() - typingStart);

  // 6. Search / indexing
  const searchStart = performance.now();
  const fullExtractedText = await extractDocxText(docxBuffer);
  const tokens = fullExtractedText.toLowerCase().split(/\s+/);
  const indexMap = new Map();
  for (const token of tokens) {
    if (token.length > 2) indexMap.set(token, (indexMap.get(token) ?? 0) + 1);
  }
  const searchIndexMs = round(performance.now() - searchStart);

  // 7. Save time
  const saveStart = performance.now();
  const serializer = new XMLSerializer();
  const updatedXml = serializer.serializeToString(xmlDoc);
  zip.file("word/document.xml", updatedXml);
  const savedBuffer = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  const saveMs = round(performance.now() - saveStart);

  // 8. Background page completion
  const bgStart = performance.now();
  let tableCellCount = 0;
  for (const tbl of allTables) {
    tableCellCount += tbl.getElementsByTagNameNS(WORD_NS, "tc").length;
  }
  const backgroundCompletionMs = round(performance.now() - bgStart);

  const memEnd = process.memoryUsage?.().heapUsed ?? 0;
  const heapUsedDeltaBytes = Math.max(0, memEnd - memStart);

  return {
    tier: pageCount,
    unit: "pages",
    documentSizeBytes,
    openToFirstVisibleMs,
    openToUsableMs,
    coldNavigationMs,
    warmNavigationMs,
    typingMutationMs,
    searchIndexMs,
    saveMs,
    backgroundCompletionMs,
    heapUsedDeltaBytes,
  };
}

export async function runPptxBenchmarkTier(slideCount) {
  const { PresentationEngine } = await loadPresentationEngineModule();

  const memStart = process.memoryUsage?.().heapUsed ?? 0;
  const pptxBuffer = generateDeterministicPptx(slideCount);
  const documentSizeBytes = pptxBuffer.byteLength;

  // 1. Open to first visible (load engine + render slide 0)
  const openStart = performance.now();
  const engine = await PresentationEngine.load(pptxBuffer.slice(0));
  const slide0 = engine.renderSlide(0);
  const openToFirstVisibleMs = round(performance.now() - openStart);

  // 2. Open to usable (slide count, layouts, background indexing ready)
  const usableStart = performance.now();
  const layouts = engine.getSlideLayouts();
  const count = engine.slideCount;
  const openToUsableMs = round(openToFirstVisibleMs + (performance.now() - usableStart));

  // 3. Cold navigation to middle slide
  const coldStart = performance.now();
  const midIndex = Math.floor(slideCount / 2);
  const midSlide = engine.renderSlide(midIndex);
  const coldNavigationMs = round(performance.now() - coldStart);

  // 4. Warm navigation back to slide 0
  const warmStart = performance.now();
  const warmSlide0 = engine.renderSlide(0);
  const warmNavigationMs = round(performance.now() - warmStart);

  // 5. Typing mutation
  const typingStart = performance.now();
  const slideXml = engine.getSlideXml(0);
  const mutatedXml = slideXml.replace("</a:t>", " [typed mutation]</a:t>");
  await engine.restoreSlideXml(0, mutatedXml);
  const typingMutationMs = round(performance.now() - typingStart);

  // 6. Search / indexing across all slides
  const searchStart = performance.now();
  let matchCount = 0;
  for (let i = 0; i < count; i++) {
    const xml = engine.getSlideXml(i);
    if (xml.includes("Slide") || xml.includes("Benchmark")) matchCount++;
  }
  const searchIndexMs = round(performance.now() - searchStart);

  // 7. Save time (serialize & export package)
  const saveStart = performance.now();
  const exported = await engine.export();
  const saveMs = round(performance.now() - saveStart);

  // 8. Background thumbnail completion for unrendered slides (up to 10 slides)
  const bgStart = performance.now();
  const maxBg = Math.min(count, 10);
  for (let i = 0; i < maxBg; i++) {
    if (i !== 0 && i !== midIndex) {
      engine.renderSlide(i);
    }
  }
  const backgroundCompletionMs = round(performance.now() - bgStart);

  const memEnd = process.memoryUsage?.().heapUsed ?? 0;
  const heapUsedDeltaBytes = Math.max(0, memEnd - memStart);

  return {
    tier: slideCount,
    unit: "slides",
    documentSizeBytes,
    openToFirstVisibleMs,
    openToUsableMs,
    coldNavigationMs,
    warmNavigationMs,
    typingMutationMs,
    searchIndexMs,
    saveMs,
    backgroundCompletionMs,
    heapUsedDeltaBytes,
  };
}

export async function runLargeDocumentBenchmarks(options = {}) {
  const docxTiers = options.tiers?.docx ?? (options.includeOptInTiers ? [...DEFAULT_DOCX_TIERS, ...OPT_IN_DOCX_TIERS] : DEFAULT_DOCX_TIERS);
  const pptxTiers = options.tiers?.pptx ?? (options.includeOptInTiers ? [...DEFAULT_PPTX_TIERS, ...OPT_IN_PPTX_TIERS] : DEFAULT_PPTX_TIERS);
  const scopes = options.scopes ?? ["docx", "pptx"];

  const results = {
    docx: [],
    pptx: [],
  };

  if (scopes.includes("docx")) {
    for (const tier of docxTiers) {
      options.onProgress?.(`Running DOCX tier ${tier} pages...`);
      const tierResult = await runDocxBenchmarkTier(tier);
      results.docx.push(tierResult);
    }
  }

  if (scopes.includes("pptx")) {
    for (const tier of pptxTiers) {
      options.onProgress?.(`Running PPTX tier ${tier} slides...`);
      const tierResult = await runPptxBenchmarkTier(tier);
      results.pptx.push(tierResult);
    }
  }

  return {
    schemaVersion: 1,
    timestamp: new Date().toISOString(),
    environment: {
      platform: process.platform,
      arch: process.arch,
      nodeVersion: process.version,
      heapTotalBytes: process.memoryUsage?.().heapTotal ?? 0,
    },
    results,
  };
}
