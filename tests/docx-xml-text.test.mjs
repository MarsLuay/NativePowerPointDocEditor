import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxXmlTextModule } from "./helpers/load-plugin-modules.mjs";

test("extractDocxRunText extracts text and handles tabs, breaks, and entities", async () => {
    const { extractDocxRunText } = await loadDocxXmlTextModule();

    assert.equal(extractDocxRunText(""), "");
    assert.equal(extractDocxRunText("<w:t>Hello</w:t>"), "Hello");
    assert.equal(extractDocxRunText("<w:t xml:space=\"preserve\">  Hello  </w:t>"), "  Hello  ");
    assert.equal(extractDocxRunText("<w:t>Hello</w:t><w:tab/><w:t>World</w:t>"), "Hello\tWorld");
    assert.equal(extractDocxRunText("<w:t>Hello</w:t><w:br/><w:t>World</w:t>"), "Hello\nWorld");
    assert.equal(extractDocxRunText("<w:t>Hello</w:t><w:cr/><w:t>World</w:t>"), "Hello\nWorld");

    // Entities
    assert.equal(extractDocxRunText("<w:t>&lt;b&gt;bold&lt;/b&gt;</w:t>"), "<b>bold</b>");
    assert.equal(extractDocxRunText("<w:t>A &amp; B</w:t>"), "A & B");
    assert.equal(extractDocxRunText("<w:t>&#34;quotes&#34;</w:t>"), "\"quotes\"");
    assert.equal(extractDocxRunText("<w:t>&#x27;single&#x27;</w:t>"), "'single'");

    // Complex combination
    const complex = '<w:t>Hello &amp; welcome</w:t><w:tab/><w:t>Tab</w:t><w:br/><w:t>Line</w:t>';
    assert.equal(extractDocxRunText(complex), "Hello & welcome\tTab\nLine");

    // Invalid formatting tags that should be ignored
    assert.equal(extractDocxRunText("<w:r><w:t>Hello</w:t></w:r>"), "Hello");

    // Test for multiline text
    assert.equal(extractDocxRunText("<w:t>Line 1\nLine 2</w:t>"), "Line 1\nLine 2");
});
