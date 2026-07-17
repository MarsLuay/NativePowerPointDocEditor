import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('DOCX top-level menus share the semantic menu surface contract', async () => {
	const [css, menuDropdown, titleBar] = await Promise.all([
		readFile(path.join(projectRoot, 'styles.css'), 'utf8'),
		readFile(path.join(projectRoot, 'docx-editor/packages/react/src/components/ui/MenuDropdown.tsx'), 'utf8'),
		readFile(path.join(projectRoot, 'docx-editor/packages/react/src/components/TitleBar.tsx'), 'utf8'),
	]);

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
