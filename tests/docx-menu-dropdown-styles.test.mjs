import assert from 'node:assert/strict';
import { existsSync, readdirSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function readReactMenuSurfaceSources() {
	const distDir = path.join(projectRoot, 'vendor/docx-editor-runtime/react/dist');
	assert.ok(existsSync(distDir), 'missing vendor/docx-editor-runtime/react/dist');
	const files = readdirSync(distDir).filter((name) => name.endsWith('.js') || name.endsWith('.mjs'));
	assert.ok(files.length > 0, 'vendor/docx-editor-runtime/react/dist has no JS bundles');
	const combined = (
		await Promise.all(files.map((name) => readFile(path.join(distDir, name), 'utf8')))
	).join('\n');
	return { menuDropdown: combined, titleBar: combined };
}

test('DOCX top-level menus share the semantic menu surface contract', async () => {
	const [css, surface] = await Promise.all([
		readFile(path.join(projectRoot, 'styles.css'), 'utf8'),
		readReactMenuSurfaceSources(),
	]);
	const { menuDropdown, titleBar } = surface;

	assert.match(css, /\[data-native-powerpoint-doc-editor-menu-dropdown\]/);
	assert.match(css, /--doc-menu-bg:\s*var\(--npde-menu-bg\)/);
	assert.match(css, /--doc-menu-item-padding:\s*6px 12px/);
	assert.match(css, /--doc-menu-item-font-size:\s*13px/);
	assert.match(css, /\.native-powerpoint-doc-editor-edit-menu/);
	assert.match(menuDropdown, /var\(--doc-menu-bg, var\(--doc-surface\)\)/);
	assert.match(menuDropdown, /var\(--doc-menu-item-padding, 6px 12px\)/);
	assert.match(menuDropdown, /var\(--doc-menu-item-hover-bg, var\(--doc-bg-hover\)\)/);
	assert.match(titleBar, /var\(--doc-menu-item-padding, 6px 12px\)/);
});
