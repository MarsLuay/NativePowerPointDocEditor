import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxXmlTextModule } from "./helpers/load-plugin-modules.mjs";

test("normalizeDocxExtractedText normalizes line endings and spaces correctly", async () => {
	const { normalizeDocxExtractedText } = await loadDocxXmlTextModule();

	// Returns normalized text, removes extra newlines, trims edge spaces
	assert.equal(normalizeDocxExtractedText("a\r\nb"), "a\n\nb");
	assert.equal(normalizeDocxExtractedText("hello\rworld"), "hello\nworld");
	assert.equal(normalizeDocxExtractedText("  \thello  \n \tworld \n \n "), "hello\nworld");
	assert.equal(normalizeDocxExtractedText("line1\n\n\n\nline2"), "line1\n\nline2");
});
