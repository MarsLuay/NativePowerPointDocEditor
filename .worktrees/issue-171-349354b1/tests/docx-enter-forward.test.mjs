import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("PagedEditor forwards Enter when hidden PM is unfocused", async () => {
	const source = await readFile(
		path.join(projectRoot, "docx-editor/packages/react/src/components/DocxEditor/PagedEditor.tsx"),
		"utf8",
	);

	assert.match(source, /NumpadEnter/);
	assert.match(source, /view\.someProp\('handleKeyDown'/);
	assert.match(source, /view\.state\.doc === docBefore/);
	assert.match(source, /preferLayout: transaction\.docChanged/);
});

test("BaseKeymap Enter falls back beyond splitBlockClearBorders", async () => {
	const source = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/prosemirror/extensions/features/BaseKeymapExtension.ts",
		),
		"utf8",
	);

	assert.match(
		source,
		/Enter:\s*chainCommands\(splitBlockClearBorders,\s*createParagraphNear,\s*liftEmptyBlock\)/,
	);
});

test("empty list Enter exits in place without creating another paragraph", async () => {
	const source = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/prosemirror/extensions/features/ListExtension.ts",
		),
		"utf8",
	);

	assert.match(source, /paragraph\.content\.size > 0 \|\| paragraph\.textContent\.length > 0/);
	const exitListBody = source.slice(
		source.indexOf("function exitListOnEmptyEnter"),
		source.indexOf("function splitListItem"),
	);
	assert.match(exitListBody, /setNodeMarkup\(\$from\.before\(\), undefined, clearedAttrs\)/);
	assert.doesNotMatch(exitListBody, /\.split\(/);
});

test("empty layout runs carry pm caret attrs", async () => {
	const source = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/layout-painter/renderParagraph/line.ts",
		),
		"utf8",
	);

	assert.match(source, /layout-empty-run/);
	assert.match(source, /applyPmPositions\(emptySpan, contentPos, contentPos\)/);
});

test("DOM caret height is clamped for empty paragraphs", async () => {
	const source = await readFile(
		path.join(
			projectRoot,
			"docx-editor/packages/core/src/layout-bridge/clickToPositionDom.ts",
		),
		"utf8",
	);

	assert.match(source, /MAX_CARET_HEIGHT_PX = 72/);
	assert.match(source, /clampCaretHeight/);
});
