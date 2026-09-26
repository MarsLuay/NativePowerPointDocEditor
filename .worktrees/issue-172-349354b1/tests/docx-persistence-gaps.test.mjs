import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("DOCX host keeps find-replace marks and routes image insert through core", async () => {
	const adapter = await readFile(
		path.join(projectRoot, "src/docx/adapter/DocxEditorAdapter.ts"),
		"utf8",
	);

	assert.match(adapter, /insertImageFromFile/);
	assert.match(adapter, /collectMarksAtRange/);
	assert.match(adapter, /schema\.text\(replacement, marks\)/);
});

test("paragraph attr edits sync _originalFormatting for serialize", async () => {
	const paraExt = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/prosemirror/extensions/core/ParagraphExtension.ts",
		),
		"utf8",
	);
	const sync = await readFile(
		path.join(projectRoot, "docx-editor/packages/core/src/prosemirror/syncOriginalFormatting.ts"),
		"utf8",
	);
	const listExt = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/prosemirror/extensions/features/ListExtension.ts",
		),
		"utf8",
	);
	const bold = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/prosemirror/extensions/marks/BoldExtension.ts",
		),
		"utf8",
	);

	assert.match(sync, /mergeParagraphAttrsWithOriginalFormatting/);
	assert.match(sync, /originalFormattingAfterApplyStyle/);
	assert.match(paraExt, /mergeParagraphAttrsWithOriginalFormatting/);
	assert.match(paraExt, /originalFormattingAfterApplyStyle/);
	assert.match(listExt, /mergeParagraphAttrsWithOriginalFormatting/);
	assert.match(bold, /toggleMarkWithParagraphDefaults/);
});
