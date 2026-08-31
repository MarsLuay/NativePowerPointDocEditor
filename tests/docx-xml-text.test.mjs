import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxXmlTextModule } from "./helpers/load-plugin-modules.mjs";

test("extractDocxTextFromXml extracts text, tabs, and line breaks", async () => {
  const { extractDocxTextFromXml } = await loadDocxXmlTextModule();
  const xml = `
    <w:p>
      <w:r>
        <w:t>Hello &amp; welcome</w:t>
        <w:tab/>
        <w:t>Tab</w:t>
        <w:br/>
        <w:t>Line</w:t>
        <w:cr/>
        <w:t>Break</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">  Space  </w:t>
      </w:r>
    </w:p>
  `;
  const result = extractDocxTextFromXml(xml);
  assert.equal(result, "Hello & welcome\tTab\nLine\nBreak\nSpace");
});

test("extractDocxRunText extracts run text without paragraph normalizations", async () => {
  const { extractDocxRunText } = await loadDocxXmlTextModule();
  const xml = `
    <w:r>
      <w:t>Hello &amp; welcome</w:t>
      <w:tab/>
      <w:t>Tab</w:t>
      <w:br/>
      <w:t>Line</w:t>
      <w:cr/>
      <w:t>Break</w:t>
    </w:r>
  `;
  const result = extractDocxRunText(xml);
  assert.equal(result, "Hello & welcome\tTab\nLine\nBreak");
});
