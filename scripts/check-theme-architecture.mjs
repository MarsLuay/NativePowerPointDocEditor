import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const srcRoot = path.resolve('src');
const allowedThemeResolutionFiles = new Set([
	path.join(srcRoot, 'main.ts'),
	path.join(srcRoot, 'settings.ts'),
]);

const sourceExtensions = new Set(['.ts', '.tsx']);
const skipDirNames = new Set(['vendor', 'node_modules', 'build', 'dist']);

function collectSourceFiles(root) {
	const files = [];

	function visit(current) {
		for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
			const fullPath = path.join(current, entry.name);
			if (entry.isDirectory()) {
				if (!skipDirNames.has(entry.name)) {
					visit(fullPath);
				}
				continue;
			}
			if (entry.isFile() && sourceExtensions.has(path.extname(entry.name))) {
				files.push(fullPath);
			}
		}
	}

	visit(root);
	return files;
}

function lineForIndex(text, index) {
	return text.slice(0, index).split('\n').length;
}

function readSource(relativePath) {
	return fs.readFileSync(path.join(srcRoot, relativePath), 'utf8');
}

const failures = [];
for (const filePath of collectSourceFiles(srcRoot)) {
	const text = fs.readFileSync(filePath, 'utf8');
	const relativePath = path.relative(process.cwd(), filePath);
	const canResolveTheme = allowedThemeResolutionFiles.has(filePath);

	if (!canResolveTheme) {
		for (const match of text.matchAll(/\bresolveEditorThemePreference\b/g)) {
			failures.push(
				`${relativePath}:${lineForIndex(text, match.index ?? 0)} should consume resolvedEditorTheme instead of resolving theme locally.`,
			);
		}
		for (const match of text.matchAll(/classList(?:\?\.)?\.contains\(\s*['"]theme-dark['"]/g)) {
			failures.push(
				`${relativePath}:${lineForIndex(text, match.index ?? 0)} should not inspect Obsidian's body theme directly.`,
			);
		}
	}
}

const mainSource = readSource('main.ts');
const settingsSource = readSource('settings.ts');
const docxViewSource = readSource('DocxView.tsx');
const docxReactViewSource = readSource('DocxReactView.tsx');
const pptxViewSource = readSource('NativePowerPointView.ts');
const docxSupportSource = readSource('docxSupport.ts');

assert.match(
	mainSource,
	/getResolvedEditorTheme\(\)/,
	'main.ts must expose getResolvedEditorTheme() as the central resolved theme accessor.',
);
assert.match(
	mainSource,
	/resolveCurrentEditorTheme\(\)/,
	'main.ts must keep resolved theme computation centralized.',
);
assert.match(
	settingsSource,
	/resolvedEditorTheme:\s*EditorThemeResolution/,
	'settings.ts must carry resolvedEditorTheme in the PowerPoint settings shape.',
);
assert.match(
	docxViewSource,
	/getResolvedEditorTheme/,
	'DocxView must receive resolved editor theme instead of recomputing it.',
);
assert.match(
	docxSupportSource,
	/getResolvedEditorTheme/,
	'docxSupport must pass the plugin resolved theme accessor into DocxView.',
);
assert.match(
	docxViewSource,
	/resolvedEditorTheme:\s*this\.getResolvedEditorTheme\(\)/,
	'DocxView must pass the centralized resolved theme into DocxReactView.',
);
assert.match(
	docxReactViewSource,
	/colorMode=\{resolvedEditorTheme\}/,
	'DocxReactView must pass the centralized resolved theme into the DOCX editor.',
);
assert.doesNotMatch(
	docxReactViewSource,
	/colorMode=["']light["']/,
	'DocxReactView must not force the DOCX editor library into light mode.',
);
assert.match(
	pptxViewSource,
	/resolvedEditorTheme/,
	'NativePowerPointView must consume resolvedEditorTheme from settings.',
);

assert.deepEqual(
	failures,
	[],
	`Theme architecture guard failed:\n${failures.join('\n')}`,
);

console.log('Theme architecture check passed: views consume the centralized resolved theme.');
