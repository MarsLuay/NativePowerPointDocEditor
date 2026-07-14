import assert from "node:assert/strict";
import { test } from "node:test";
import JSZip from "jszip";
import { loadDocxTextExtractorModule } from "./helpers/load-plugin-modules.mjs";

// Builds a minimal, self-contained .docx (OOXML zip) in memory. The export
// pipeline in DocxView.createExportContent feeds extractDocxText's output into
// the txt/md/html/rtf derivations, so the plain-text contract verified here is
// what those non-docx exports ultimately serialize.
async function createDocxBuffer(parts) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  for (const [path, xml] of Object.entries(parts)) {
    zip.file(path, xml);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

function wrapBody(...inner) {
  return [
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
    ...inner,
    "</w:body></w:document>",
  ].join("");
}

test("extractDocxText returns clean plain text for the txt export path", async () => {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>",
      "<w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>",
    ),
  });

  const text = await extractDocxText(buffer);

  assert.equal(text, "First paragraph\nSecond paragraph");
  // The txt export is `${text}\n` (DocxView.createExportContent), so the
  // derived payload is deterministic given this extracted text.
  assert.equal(`${text}\n`, "First paragraph\nSecond paragraph\n");
});

test("extractDocxMarkdown separates paragraphs with a blank line (not txt-identical)", async () => {
  const { extractDocxText, extractDocxMarkdown } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>",
      "<w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>",
    ),
  });

  const text = await extractDocxText(buffer);
  const markdown = await extractDocxMarkdown(buffer);

  // Markdown must be real Markdown, not byte-identical to the plain-text export.
  assert.equal(markdown, "First paragraph\n\nSecond paragraph");
  assert.notEqual(`${markdown}\n`, `${text}\n`);
});

test("extractDocxMarkdown preserves an inline break as a Markdown hard break", async () => {
  const { extractDocxMarkdown } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p><w:r><w:t>First line</w:t><w:br/><w:t>Second line</w:t></w:r></w:p>",
    ),
  });

  assert.equal(await extractDocxMarkdown(buffer), "First line  \nSecond line");
});

test("extractDocxMarkdown maps headings, emphasis, and list items", async () => {
  const { extractDocxMarkdown } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title here</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Subsection</w:t></w:r></w:p>',
      "<w:p>",
      "<w:r><w:t>Plain </w:t></w:r>",
      '<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>',
      "<w:r><w:t> and </w:t></w:r>",
      '<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>',
      "<w:r><w:t> text</w:t></w:r>",
      "</w:p>",
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>',
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Nested item</w:t></w:r></w:p>',
    ),
  });

  const markdown = await extractDocxMarkdown(buffer);

  assert.equal(
    markdown,
    [
      "# Title here",
      "## Subsection",
      "Plain **bold** and *italic* text",
      "- First item",
      "  - Nested item",
    ].join("\n\n"),
  );
});

test("extractDocxMarkdown ignores bold toggles disabled via w:val", async () => {
  const { extractDocxMarkdown } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      '<w:p><w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>Not bold</w:t></w:r></w:p>',
    ),
  });

  assert.equal(await extractDocxMarkdown(buffer), "Not bold");
});

test("extractDocxText decodes entities and preserves tab/break structure", async () => {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p>",
      "<w:r><w:t>A &amp; B &lt;C&gt; &#65;&#x42;</w:t></w:r>",
      "<w:r><w:tab/></w:r>",
      "<w:r><w:t>tabbed</w:t></w:r>",
      "<w:r><w:br/></w:r>",
      "<w:r><w:t>broken line</w:t></w:r>",
      "</w:p>",
    ),
  });

  assert.equal(
    await extractDocxText(buffer),
    "A & B <C> AB\ttabbed\nbroken line",
  );
});

test("extractDocxText splits paragraphs so the HTML export can group them", async () => {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p><w:r><w:t>Heading</w:t></w:r></w:p>",
      "<w:p></w:p>",
      "<w:p><w:r><w:t>Body text</w:t></w:r></w:p>",
    ),
  });

  const text = await extractDocxText(buffer);

  // createPlainTextHtml splits on /\n{2,}/ into <p> blocks. The extractor must
  // therefore surface a blank-line boundary between the two real paragraphs.
  assert.equal(text, "Heading\n\nBody text");
  const htmlParagraphs = text.split(/\n{2,}/).map((p) => p.trim()).filter(Boolean);
  assert.deepEqual(htmlParagraphs, ["Heading", "Body text"]);
});

test("extractDocxText orders document body before headers/footnotes parts", async () => {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/headers/header1.xml":
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>',
    "word/footnotes.xml":
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footnote text</w:t></w:r></w:p></w:footnotes>',
    "word/document.xml": wrapBody("<w:p><w:r><w:t>Body text</w:t></w:r></w:p>"),
  });

  // sortTextParts pins document.xml first, then localeCompare orders the rest
  // (footnotes < headers), with parts joined by a blank line.
  assert.equal(
    await extractDocxText(buffer),
    "Body text\n\nFootnote text\n\nHeader text",
  );
});

test("extractDocxText normalizes runaway whitespace before serialization", async () => {
  const { extractDocxText } = await loadDocxTextExtractorModule();
  const buffer = await createDocxBuffer({
    "word/document.xml": wrapBody(
      "<w:p><w:r><w:t>Top</w:t></w:r></w:p>",
      "<w:p></w:p>",
      "<w:p></w:p>",
      "<w:p></w:p>",
      "<w:p><w:r><w:t>Bottom   </w:t></w:r></w:p>",
    ),
  });

  // Three-plus newlines collapse to a single blank line and trailing spaces are
  // trimmed, so every text-based export starts from a tidy string.
  assert.equal(await extractDocxText(buffer), "Top\n\nBottom");
});
