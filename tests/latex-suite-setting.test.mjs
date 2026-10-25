import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
	findClosedMathSpans,
	findMathSpansToCommit,
	isMathReadyToCommit,
	unwrapLatexSuiteInner,
} from '../src/latexSuite/mathSpans.ts';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('LaTeX Suite setting defaults off and stays out of the DOCX menu', async () => {
	const settings = await readFile(path.join(projectRoot, 'src/settings.ts'), 'utf8');
	assert.match(settings, /enableLatexSuite: false/);
	assert.match(settings, /raw\.enableLatexSuite === true/);
	assert.match(settings, /isLatexSuiteInstalled\(this\.app\)/);
	assert.match(settings, /control: \{ type: 'toggle' as const, key: 'enableLatexSuite' \}/);
	assert.match(settings, /this\.plugin\.applyLatexSuite\(\)/);

	const catalog = await readFile(path.join(projectRoot, 'src/i18n/settingsCatalog.ts'), 'utf8');
	const docxMenu = catalog.slice(catalog.indexOf('export function getDocxEditorSettingsMenuSections'));
	assert.doesNotMatch(docxMenu, /enableLatexSuite/);

	const english = JSON.parse(await readFile(path.join(projectRoot, 'locales/en/settings.json'), 'utf8'));
	assert.equal(english.latexSuite.enableLatexSuite.name, 'LaTeX Suite in Office files');
	assert.match(english.latexSuite.enableLatexSuite.description, /Off by default/);
	const polish = JSON.parse(await readFile(path.join(projectRoot, 'locales/pl/settings.json'), 'utf8'));
	assert.equal(typeof polish.latexSuite.enableLatexSuite.name, 'string');
});

test('LaTeX Suite detection uses the installed plugin manifest', async () => {
	const source = await readFile(path.join(projectRoot, 'src/latexSuite/detect.ts'), 'utf8');
	assert.match(source, /export const LATEX_SUITE_PLUGIN_ID = 'obsidian-latex-suite'/);
	assert.match(source, /manifests\?\.\s*\[\s*LATEX_SUITE_PLUGIN_ID\s*\]/);
});

test('closed math spans skip the caret and empty braces', () => {
	assert.deepEqual(findClosedMathSpans('$x^2$', 5).map((span) => span.inner), ['x^2']);
	assert.equal(findClosedMathSpans('$x^2$', 2).length, 0);
	assert.equal(findClosedMathSpans('${}${}', 0).length, 0);
	assert.equal(findClosedMathSpans('$$a+b$$', 7)[0]?.display, 'block');
	assert.equal(findClosedMathSpans('\\(a\\)', 5)[0]?.inner, 'a');
});

test('mk-wrapped cdot is ready to commit on Tab even with caret inside', () => {
	const source = '${}\\cdot{}$';
	const caretAfterCdot = source.indexOf('\\cdot') + '\\cdot'.length;
	assert.equal(unwrapLatexSuiteInner('{}\\cdot{}'), '\\cdot');
	assert.equal(findClosedMathSpans(source, caretAfterCdot).length, 0);
	const spans = findMathSpansToCommit(source, caretAfterCdot, { commitReadyAtCaret: true });
	assert.equal(spans.length, 1);
	assert.equal(unwrapLatexSuiteInner(spans[0].inner), '\\cdot');
	assert.equal(isMathReadyToCommit(source, spans[0], caretAfterCdot), true);
	assert.equal(findMathSpansToCommit('${}{}$', 3, { commitReadyAtCaret: true }).length, 0);
	assert.equal(findMathSpansToCommit('${} \\cdot{}$', '${} \\cdot{}$'.length).map((span) => unwrapLatexSuiteInner(span.inner))[0], '\\cdot');

	const fraction = '${}\\frac{a}{}{}$';
	const caretAfterNumerator = fraction.indexOf('a') + 1;
	assert.equal(findMathSpansToCommit(fraction, caretAfterNumerator, { commitReadyAtCaret: true }).length, 0);
});

test('plugin wires a LaTeX Suite host through settings', async () => {
	const main = await readFile(path.join(projectRoot, 'src/main.ts'), 'utf8');
	assert.match(main, /this\.setupLatexSuite\(\)/);
	assert.match(main, /this\.latexSuiteHost\?\.destroy\(\)/);
	assert.match(main, /applyLatexSuite\(\): void/);
	const host = await readFile(path.join(projectRoot, 'src/latexSuite/host.ts'), 'utf8');
	assert.match(host, /commitReadyAtCaret: finish/);
	assert.match(host, /commitNonEmptyAtCaret: true/);
	const view = await readFile(path.join(projectRoot, 'src/DocxReactView.tsx'), 'utf8');
	assert.match(view, /handleDocxKeydown\(view, event\)/);
	assert.match(view, /handleDocxBlur\(view\)/);
	assert.match(view, /handleDocxCaretSettled\(view\)/);
	const pptx = await readFile(path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
	assert.match(pptx, /handleTextareaKeydown\(event, editor/);
	assert.match(pptx, /handleTextareaBlur\(editor/);
	assert.match(pptx, /handleTextareaCaretSettled\(editor/);
});
