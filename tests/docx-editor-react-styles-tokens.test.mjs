import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const stylesPath = path.join(
	projectRoot,
	'docx-editor/packages/react/dist/styles.css',
);

test('react dist styles.css ships .docx-editor-root theme tokens for portaled menus', async () => {
	const css = await readFile(stylesPath, 'utf8');
	assert.match(css, /\.docx-editor-root\s*\{/);
	assert.match(css, /--doc-surface\s*:/);
	assert.match(css, /--popover\s*:/);
	assert.match(css, /\.docx-editor-root\s+\.bg-popover\b/);
});
