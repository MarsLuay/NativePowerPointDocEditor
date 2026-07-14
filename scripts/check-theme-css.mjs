import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const stylePath = path.resolve('styles.css');
const sourceRoot = path.resolve('src');
const renderedPdfExportPath = path.join(sourceRoot, 'renderedPdfExport.ts');

function lineForIndex(text, index) {
	return text.slice(0, index).split('\n').length;
}

function selectorForDeclaration(text, index) {
	const blockStart = text.lastIndexOf('{', index);
	const previousBlockEnd = text.lastIndexOf('}', index);
	if (blockStart === -1 || previousBlockEnd > blockStart) {
		return '';
	}
	const selectorStart = text.lastIndexOf('}', blockStart - 1) + 1;
	return text.slice(selectorStart, blockStart).trim().replace(/\s+/g, ' ');
}

function isAllowedImportantSelector(selector) {
	return selector.includes('.workspace-leaf-content')
		&& selector.includes('.native-powerpoint-doc-editor-host')
		&& selector.includes('[data-native-powerpoint-doc-editor-root]');
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

const css = fs.readFileSync(stylePath, 'utf8');
const failures = [];

const requiredDocxScrollbarFragments = [
	'--npde-docx-toolbar-shell-bg: var(--npde-chrome-bg);',
	'--npde-docx-formatting-bar-bg: var(--npde-toolbar-bg);',
	'background: var(--npde-docx-toolbar-shell-bg) !important;',
	'background: var(--npde-docx-formatting-bar-bg) !important;',
	"[data-native-powerpoint-doc-editor-formatting-bar]::-webkit-scrollbar {",
	"[data-native-powerpoint-doc-editor-formatting-bar]::-webkit-scrollbar-track",
	"[data-native-powerpoint-doc-editor-formatting-bar]::-webkit-scrollbar-thumb",
	"[data-native-powerpoint-doc-editor-formatting-bar]::-webkit-scrollbar-corner",
	'[data-native-powerpoint-doc-editor-scroll-container] {',
	'--doc-scrollbar-track: var(--npde-docx-document-scrollbar-track);',
];

const requiredDocxDropdownFragments = [
	"[data-native-powerpoint-doc-editor-formatting-bar].native-powerpoint-doc-editor-formatting-dropdown-open {",
	'overflow: visible;',
	"left: calc(-1 * var(--native-powerpoint-doc-editor-formatting-scroll-left, 0px));",
	':not(.docx-color-picker-dropdown button)',
	"[style*='position: fixed'][role='dialog'][aria-modal='true']",
];

const requiredDocxDarkDocumentFragments = [
	'--doc-bg: var(--npde-editor-bg);',
	'--doc-text: var(--npde-editor-text);',
	'--doc-caret: var(--npde-document-text);',
	'[data-native-powerpoint-doc-editor-page] {',
	'background: var(--npde-document-bg) !important;',
	'color: var(--npde-document-text) !important;',
	'color-scheme: light !important;',
	'filter: none !important;',
];

const requiredSettingsButtonFragments = [
	'.native-powerpoint-doc-editor-editor-settings-action:focus-visible',
	'.native-powerpoint-doc-editor-editor-settings-menu',
	'.native-powerpoint-doc-editor-editor-settings-action',
	'.native-powerpoint-doc-editor-editor-settings-row > button',
	'.native-powerpoint-doc-editor-editor-settings-row.mod-action > button',
	'--npde-settings-action-border:',
	'border: 1px solid var(--npde-editor-border-strong);',
	'box-shadow: none !important;',
	'outline: 2px solid var(--npde-toolbar-focus-ring);',
];

for (const fragment of requiredDocxScrollbarFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must keep DOCX scrollbar styling on the actual scrolling element: ${fragment}`,
		);
	}
}

for (const fragment of requiredDocxDropdownFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must keep DOCX toolbar dropdowns unclipped and preserve component-owned color swatches: ${fragment}`,
		);
	}
}

if (/:is\(button,\s*input,\s*select,\s*textarea,[^)]*\)\s*\{\s*background:\s*transparent\s*!important;/s.test(css)) {
	failures.push(
		`${path.relative(process.cwd(), stylePath)} must not blank every DOCX toolbar button; color-picker swatches own their inline backgrounds.`,
	);
}

if (/\[style\*='position: fixed'\]\[role='dialog'\]\s*\{/.test(css)) {
	failures.push(
		`${path.relative(process.cwd(), stylePath)} must not stretch non-modal fixed dialogs such as DOCX color pickers to the viewport.`,
	);
}

for (const fragment of requiredDocxDarkDocumentFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must keep DOCX dark mode chrome separate from the white document page: ${fragment}`,
		);
	}
}

for (const fragment of requiredSettingsButtonFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must keep settings action buttons visibly outlined in resolved themes: ${fragment}`,
		);
	}
}

const requiredDocxEigenpalTooltipHideFragments = [
	"data-native-powerpoint-doc-editor-toolbar-tooltips='custom'",
	"[data-native-powerpoint-doc-editor-eigenpal-tooltip='true']",
	'display: none !important;',
	'opacity: 0 !important;',
	'visibility: hidden !important;',
];

const requiredDocxToolbarTooltipFragments = [
	'--npde-docx-toolbar-tooltip-bg: #0f172a;',
	'--npde-docx-toolbar-tooltip-text: #f8fafc;',
	'background: var(--npde-docx-toolbar-tooltip-bg);',
	'color: var(--npde-docx-toolbar-tooltip-text);',
];

for (const fragment of requiredDocxEigenpalTooltipHideFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must hide Eigenpal toolbar tooltips while the plugin renders its own labels: ${fragment}`,
		);
	}
}

for (const fragment of requiredDocxToolbarTooltipFragments) {
	if (!css.includes(fragment)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)} must keep DOCX toolbar tooltips on the fixed dark palette: ${fragment}`,
		);
	}
}

if (css.includes("[data-native-powerpoint-doc-editor-toolbar]::-webkit-scrollbar")) {
	failures.push(
		`${path.relative(process.cwd(), stylePath)} must not attach DOCX scrollbar styling to the non-scrolling editor-toolbar parent.`,
	);
}

if (/\[data-native-powerpoint-doc-editor-toolbar\]\s*>\s*div\s*\{/.test(css)) {
	failures.push(
		`${path.relative(process.cwd(), stylePath)} must keep the DOCX toolbar shell and its direct child surfaces independently themed.`,
	);
}

for (const match of css.matchAll(/!important/g)) {
	const index = match.index ?? 0;
	const selector = selectorForDeclaration(css, index);
	if (!isAllowedImportantSelector(selector)) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)}:${lineForIndex(css, index)} uses !important outside the DOCX editor override allowlist.`,
		);
	}
}

const colorLiteralRe = /#[0-9a-fA-F]{3,8}\b|rgba?\(|hsla?\(/;
css.split('\n').forEach((line, index) => {
	if (colorLiteralRe.test(line) && !line.includes('--npde-')) {
		failures.push(
			`${path.relative(process.cwd(), stylePath)}:${index + 1} should route hardcoded colors through --npde-* tokens.`,
		);
	}
});

for (const filePath of collectSourceFiles(sourceRoot)) {
	const text = fs.readFileSync(filePath, 'utf8');
	for (const match of text.matchAll(/!important/g)) {
		const index = match.index ?? 0;
		if (filePath !== renderedPdfExportPath) {
			failures.push(
				`${path.relative(process.cwd(), filePath)}:${lineForIndex(text, index)} uses !important outside styles.css/PDF export isolation.`,
			);
			continue;
		}
		const line = text.split('\n')[lineForIndex(text, index) - 1] ?? '';
		if (!line.includes('native-powerpoint-doc-editor-pdf-export')) {
			failures.push(
				`${path.relative(process.cwd(), filePath)}:${lineForIndex(text, index)} PDF export !important must stay scoped to native-powerpoint-doc-editor-pdf-export selectors.`,
			);
		}
	}
}

assert.deepEqual(
	failures,
	[],
	`Theme CSS guard failed:\n${failures.join('\n')}`,
);

console.log('Theme CSS check passed: tokens and !important allowlists are respected.');
