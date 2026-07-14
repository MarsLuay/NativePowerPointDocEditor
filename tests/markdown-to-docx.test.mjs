import assert from "node:assert/strict";
import { test } from "node:test";
import JSZip from "jszip";

import { loadMarkdownToDocxModule } from "./helpers/load-plugin-modules.mjs";

test("buildMarkdownDocxArrayBuffer creates an editable DOCX with Markdown structure", async () => {
  const { buildMarkdownDocxArrayBuffer } = await loadMarkdownToDocxModule();
  const buffer = await buildMarkdownDocxArrayBuffer([
    "---",
    "category: test",
    "---",
    "# Document title",
    "",
    "A **bold** and *italic* paragraph with `inline code` and [a link](https://example.com).",
    "",
    "- Bullet item",
    "1. Numbered item",
    "> Quoted text",
    "",
    "```ts",
    "const answer = 42;",
    "```",
  ].join("\n"));
  const zip = await JSZip.loadAsync(buffer);

  for (const path of [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/numbering.xml",
  ]) {
    assert.ok(zip.file(path), `Expected DOCX part ${path}`);
  }

  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /<w:pStyle w:val="Heading1"\/>/);
  assert.match(documentXml, /<w:b\/><w:bCs\/>/);
  assert.match(documentXml, /<w:i\/><w:iCs\/>/);
  assert.match(documentXml, /w:ascii="Courier New"/);
  assert.match(documentXml, /<w:numId w:val="1"\/>/);
  assert.match(documentXml, /<w:numId w:val="2"\/>/);
  assert.match(documentXml, /Quoted text/);
  assert.match(documentXml, /a link/);
  assert.match(documentXml, /https:\/\/example\.com/);
  assert.doesNotMatch(documentXml, /category: test/);
});

test("buildMarkdownDocxArrayBuffer escapes XML and removes invalid control characters", async () => {
  const { buildMarkdownDocxArrayBuffer } = await loadMarkdownToDocxModule();
  const zip = await JSZip.loadAsync(await buildMarkdownDocxArrayBuffer("A\t& B < C > D \"quoted\" 'value'\u0000\u0001\u0008\u000B\u000C\u000E\u001F"));
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.match(documentXml, /A\t&amp; B &lt; C &gt; D &quot;quoted&quot; &apos;value&apos;/);
  for (const character of ["\u0000", "\u0001", "\u0008", "\u000B", "\u000C", "\u000E", "\u001F"]) {
    assert.equal(documentXml.includes(character), false);
  }
});

test("resolveMarkdownDocxOutputPath creates a numbered sibling on collisions", async () => {
  const { buildMarkdownDocxCandidatePath, resolveMarkdownDocxOutputPath } = await loadMarkdownToDocxModule();
  const existing = new Set(["Notes/Plan.docx", "Notes/Plan 2.docx"]);

  assert.equal(buildMarkdownDocxCandidatePath("Notes/Plan.MD"), "Notes/Plan.docx");
  assert.equal(
    resolveMarkdownDocxOutputPath("Notes/Plan.md", path => existing.has(path)),
    "Notes/Plan 3.docx",
  );
  assert.throws(() => buildMarkdownDocxCandidatePath("Notes/Plan.txt"), /must end in \.md/);
});
