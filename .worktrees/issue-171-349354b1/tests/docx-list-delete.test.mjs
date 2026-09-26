import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("DOCX visual surface forwards deletion keys to the hidden ProseMirror keymap", async () => {
	const source = await readFile(
		path.join(projectRoot, "docx-editor/packages/react/src/components/DocxEditor/PagedEditor.tsx"),
		"utf8",
	);

	assert.match(source, /e\.key === 'Backspace' \|\| e\.key === 'Delete'/);
	assert.match(source, /view\.someProp\('handleKeyDown', \(f: Function\) => f\(view, e\.nativeEvent\)\)/);
	assert.match(source, /deleteCharacterAfterFocusRecovery\(view, e\.key\)/);
	assert.match(source, /tr\.delete\(selection\.from - 1, selection\.from\)/);
	assert.match(source, /e\.defaultPrevented && !isDeletionKey/);
});

test("DOCX list keymap removes markers and list indentation with either deletion key", async () => {
	const source = await readFile(
		path.join(projectRoot, "docx-editor/packages/core/src/prosemirror/extensions/features/ListExtension.ts"),
		"utf8",
	);

	assert.match(source, /export function exitListAtCaretStart/);
	assert.match(source, /indentLeft: null/);
	assert.match(source, /indentFirstLine: null/);
	assert.match(source, /hangingIndent: null/);
	assert.match(source, /Backspace: exitListAtCaretStart\(\)/);
	assert.match(source, /Delete: exitListAtCaretStart\(\)/);
});
