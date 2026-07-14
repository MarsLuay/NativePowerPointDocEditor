import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const srcRoot = path.resolve('src');
const menuControlsPath = path.join(srcRoot, 'menuControls.ts');

function readSource(relativePath) {
	return fs.readFileSync(path.join(srcRoot, relativePath), 'utf8');
}

function lineForIndex(text, index) {
	return text.slice(0, index).split('\n').length;
}

function collectSourceFiles(root) {
	const files = [];
	const skipDirNames = new Set(['vendor', 'node_modules', 'build', 'dist']);

	function visit(current) {
		for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
			const fullPath = path.join(current, entry.name);
			if (entry.isDirectory()) {
				if (!skipDirNames.has(entry.name)) {
					visit(fullPath);
				}
				continue;
			}
			if (entry.isFile() && ['.ts', '.tsx'].includes(path.extname(entry.name))) {
				files.push(fullPath);
			}
		}
	}

	visit(root);
	return files;
}

const menuControls = fs.readFileSync(menuControlsPath, 'utf8');
const helperExports = [
	'createPopoverShell',
	'positionPopoverBelow',
	'bindPopoverDismissHandlers',
	'configureMenuItemButton',
	'createMenuItem',
	'createMenuSection',
	'createMenuRow',
	'createCheckboxRow',
	'createSelectRow',
	'createActionRow',
];

for (const helper of helperExports) {
	assert.match(
		menuControls,
		new RegExp(`export function ${helper}\\b`),
		`menuControls.ts must continue exporting ${helper}.`,
	);
}

const requiredConsumers = [
	['DocxView.tsx', ['createPopoverShell', 'createMenuItem']],
	['powerpoint/ui/NativePowerPointView.ts', ['createPopoverShell', 'createMenuItem']],
	['powerpoint/findReplaceController.ts', ['createPopoverShell', 'createMenuItem']],
	['powerpoint/insertController.ts', ['createPopoverShell', 'createMenuItem']],
	['powerpoint/menuBarController.ts', ['createPopoverShell', 'createMenuItem', 'createMenuSection']],
	['powerpoint/textToolbarController.ts', ['createPopoverShell', 'createMenuItem']],
];

for (const [relativePath, helpers] of requiredConsumers) {
	const source = readSource(relativePath);
	assert.match(
		source,
		/menuControls/,
		`${relativePath} should use shared menuControls helpers for popovers/menus.`,
	);
	for (const helper of helpers) {
		assert.match(
			source,
			new RegExp(`\\b${helper}\\b`),
			`${relativePath} should keep using ${helper}.`,
		);
	}
}

const docxView = readSource('DocxView.tsx');
assert.match(
	docxView,
	/openPluginSettings\(\)/,
	'DocxView.tsx settings menu item should link to the plugin settings tab.',
);
assert.match(
	docxView,
	/openTabById\('native-powerpoint-doc-editor'\)/,
	'DocxView.tsx settings menu item should open the Native PowerPoint Doc Editor settings tab.',
);
assert.doesNotMatch(
	docxView,
	/renderEditorSettingsMenu/,
	'DocxView.tsx should not render a separate native DOCX settings menu.',
);

const failures = [];
const rawMenuItemRoleRe = /setAttribute\(\s*['"]role['"]\s*,\s*['"]menuitem['"]\s*\)/g;

for (const filePath of collectSourceFiles(srcRoot)) {
	if (filePath === menuControlsPath) {
		continue;
	}
	const text = fs.readFileSync(filePath, 'utf8');
	for (const match of text.matchAll(rawMenuItemRoleRe)) {
		const index = match.index ?? 0;
		const nearbySource = text.slice(Math.max(0, index - 1200), index + 300);
		const isAllowedDocxMenubarPatch = path.basename(filePath) === 'DocxView.tsx'
			&& /native-powerpoint-doc-editor(?:NoToolbarTooltip|-search-menu-button|-edit-menu-button|-settings-menu-button)/.test(nearbySource);
		if (!isAllowedDocxMenubarPatch) {
			failures.push(
				`${path.relative(process.cwd(), filePath)}:${lineForIndex(text, index)} should use configureMenuItemButton/createMenuItem instead of raw role="menuitem" wiring.`,
			);
		}
	}
}

assert.deepEqual(
	failures,
	[],
	`Shared UI pattern guard failed:\n${failures.join('\n')}`,
);

console.log('Shared UI pattern check passed: popovers and menu rows stay on shared helpers.');
