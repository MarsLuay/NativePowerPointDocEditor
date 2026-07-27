import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('DOCX chrome reconciles after the vendor editor view becomes ready', async () => {
	const view = await readFile(path.join(projectRoot, 'src/DocxView.tsx'), 'utf8');

	assert.match(view, /private reconcileEditorChromeAfterViewReady\(\)/);
	assert.match(view, /this\.runEditorChromeSync\(\);/);
	assert.match(view, /DOCX chrome reconciled after editor view ready/);
	assert.match(
		view,
		/if \(phase === 'editor-view-ready'\) \{\s*this\.reconcileEditorChromeAfterViewReady\(\);\s*\}/,
	);
});

test('DOCX chrome fallback nodes use their parent owner document', async () => {
	const view = await readFile(path.join(projectRoot, 'src/DocxView.tsx'), 'utf8');

	assert.match(view, /private createDetachedEditorChromeElement/);
	assert.match(view, /createDetachedDocxEditorChromeElement\(parent, tagName\)/);
	assert.doesNotMatch(view, /activeDocument\.create(?:El|Div|Span)\(/);
	for (const menu of ['edit', 'search', 'settings', 'duplicate', 'export-as', 'find-hidden-text', 'insert-image']) {
		assert.match(view, new RegExp(`createDetachedEditorChromeElement\\([^)]*'${menu}'\\)`));
	}
});
