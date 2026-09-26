import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const { DOMParser } = require('@xmldom/xmldom');

globalThis.DOMParser ??= DOMParser;

let modulePromise;

async function loadDrawingMlTextModule() {
	modulePromise ??= (async () => {
		const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-drawingml-text-'));
		const outfile = path.join(outputDirectory, 'drawingml-text.cjs');
		await build({
		stdin: {
			contents: [
				"export { ensureDefaultShrinkAutofit } from './src/powerpoint/drawingmlText.ts';",
				"export { isDrawingTextBoxShape } from './src/powerpoint/drawingmlText.ts';",
				"export { deleteDrawingTextRanges } from './src/powerpoint/drawingmlText.ts';",
				"export { getDrawingParagraphs, getDrawingRunText, getDrawingRuns } from './src/powerpoint/drawingmlText.ts';",
				"export { setDrawingText } from './src/powerpoint/drawingmlText.ts';",
				"export { setDrawingParagraphText } from './src/powerpoint/drawingmlText.ts';",
				"export { splitDrawingParagraphAtOffset } from './src/powerpoint/drawingmlText.ts';",
				"export { getDrawingParagraphFontSummary } from './src/powerpoint/drawingmlText.ts';",
				"export { EMPTY_PARAGRAPH_RENDER_ANCHOR } from './src/powerpoint/drawingmlText.ts';",
				"export { replaceDrawingParagraphs } from './src/powerpoint/drawingmlText.ts';",
				"export { hasEmptyDrawingParagraphBefore, removeEmptyDrawingParagraphBefore, mergeDrawingParagraphWithPrevious } from './src/powerpoint/drawingmlText.ts';",
				"export { getDescendants, parseXml } from './src/powerpoint/ooxmlXml.ts';",
			].join('\n'),
				resolveDir: projectRoot,
				sourcefile: 'drawingml-text-test-entry.ts',
			},
			bundle: true,
			format: 'cjs',
			logLevel: 'silent',
			outfile,
			platform: 'node',
			target: 'node22',
		});
		return require(outfile);
	})();
	return modulePromise;
}

test('edited text bodies default to shrink-to-fit only when no auto-fit mode is explicit', async () => {
	const { ensureDefaultShrinkAutofit, getDescendants, parseXml } = await loadDrawingMlTextModule();
	const defaultShape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ASAP text</a:t></a:r></a:p></p:txBody>
		</p:sp>
	`, 'default-autofit.xml').documentElement;
	assert.equal(ensureDefaultShrinkAutofit(defaultShape, defaultShape.ownerDocument), true);
	assert.equal(getDescendants(defaultShape, 'normAutofit').length, 1);
	assert.equal(ensureDefaultShrinkAutofit(defaultShape, defaultShape.ownerDocument), false);

	const fixedShape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>Keep size</a:t></a:r></a:p></p:txBody>
		</p:sp>
	`, 'explicit-no-autofit.xml').documentElement;
	assert.equal(ensureDefaultShrinkAutofit(fixedShape, fixedShape.ownerDocument), false);
	assert.equal(getDescendants(fixedShape, 'noAutofit').length, 1);
	assert.equal(getDescendants(fixedShape, 'normAutofit').length, 0);
});

test('text boxes get spAutoFit instead of shrink-to-fit, and legacy normAutofit is healed', async () => {
	const { ensureDefaultShrinkAutofit, getDescendants, isDrawingTextBoxShape, parseXml } = await loadDrawingMlTextModule();

	const bareTextBox = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:nvSpPr><p:cNvPr id="1" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
			<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
		</p:sp>
	`, 'textbox-bare.xml').documentElement;
	assert.equal(isDrawingTextBoxShape(bareTextBox), true);
	assert.equal(ensureDefaultShrinkAutofit(bareTextBox, bareTextBox.ownerDocument), true);
	assert.equal(getDescendants(bareTextBox, 'spAutoFit').length, 1);
	assert.equal(getDescendants(bareTextBox, 'normAutofit').length, 0);
	assert.equal(ensureDefaultShrinkAutofit(bareTextBox, bareTextBox.ownerDocument), false);

	const shrunkLegacy = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
			<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody>
		</p:sp>
	`, 'textbox-legacy-shrink.xml').documentElement;
	assert.equal(isDrawingTextBoxShape(shrunkLegacy), true);
	assert.equal(ensureDefaultShrinkAutofit(shrunkLegacy, shrunkLegacy.ownerDocument), true);
	assert.equal(getDescendants(shrunkLegacy, 'spAutoFit').length, 1);
	assert.equal(getDescendants(shrunkLegacy, 'normAutofit').length, 0);
});

test('setDrawingText populates an empty template placeholder paragraph', async () => {
	const { getDescendants, parseXml, setDrawingText } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
		</p:sp>
	`, 'empty-template-placeholder.xml').documentElement;

	setDrawingText(shape, 'Marwan Luay');
	const textElements = getDescendants(shape, 't');
	assert.equal(textElements.length, 1);
	assert.equal(textElements[0]?.textContent, 'Marwan Luay');
});

test('setDrawingParagraphText populates an empty paragraph while retaining its paragraph structure', async () => {
	const { getDescendants, parseXml, setDrawingParagraphText } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/>
				<a:p><a:r><a:rPr lang="en-US" sz="2600" b="1"/><a:t>Styled neighbor</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>
				<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:endParaRPr/></a:p>
			</p:txBody>
		</p:sp>
	`, 'empty-paragraph.xml').documentElement;

	setDrawingParagraphText(shape, 1, 'Hardware review is still required.');
	const textElements = getDescendants(shape, 't');
	const paragraph = getDescendants(shape, 'p')[1];
	assert.equal(textElements.length, 2);
	assert.equal(textElements[1]?.textContent, 'Hardware review is still required.');
	assert.equal(getDescendants(paragraph, 'pPr').length, 1);
	assert.equal(getDescendants(paragraph, 'r').length, 1);
	assert.equal(getDescendants(paragraph, 'rPr')[0]?.getAttribute('sz'), '2600');
});

test('setDrawingParagraphText keeps existing run fonts when inline text changes', async () => {
	const { getDescendants, getDrawingParagraphs, getDrawingRunText, parseXml, setDrawingParagraphText } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/><a:p>
				<a:r><a:rPr sz="2400"><a:latin typeface="Aptos"/></a:rPr><a:t>ASAP </a:t></a:r>
				<a:r><a:rPr sz="3400" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>filtration</a:t></a:r>
				<a:endParaRPr/>
			</a:p></p:txBody>
		</p:sp>
	`, 'preserve-inline-run-fonts.xml').documentElement;

	setDrawingParagraphText(shape, 0, 'ASAP solid-phase filtration');
	const runs = getDescendants(getDrawingParagraphs(shape)[0], 'r');
	assert.deepEqual(runs.map(getDrawingRunText), ['ASAP ', 'solid-phase filtration']);
	assert.equal(getDescendants(runs[0], 'rPr')[0]?.getAttribute('sz'), '2400');
	assert.equal(getDescendants(runs[0], 'latin')[0]?.getAttribute('typeface'), 'Aptos');
	assert.equal(getDescendants(runs[1], 'rPr')[0]?.getAttribute('sz'), '3400');
	assert.equal(getDescendants(runs[1], 'rPr')[0]?.getAttribute('b'), '1');
	assert.equal(getDescendants(runs[1], 'latin')[0]?.getAttribute('typeface'), 'Arial');
});

test('removeEmptyDrawingParagraphBefore removes only a structurally empty predecessor', async () => {
	const {
		getDescendants,
		parseXml,
		hasEmptyDrawingParagraphBefore,
		removeEmptyDrawingParagraphBefore,
	} = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/>
				<a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2600"/><a:t></a:t></a:r><a:endParaRPr/></a:p>
				<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="3200" b="1"/><a:t>Keep this paragraph</a:t></a:r><a:endParaRPr/></a:p>
			</p:txBody>
		</p:sp>
	`, 'remove-empty-paragraph.xml').documentElement;

	assert.equal(hasEmptyDrawingParagraphBefore(shape, 1), true);
	assert.equal(removeEmptyDrawingParagraphBefore(shape, 1), true);
	const paragraphs = getDescendants(shape, 'p');
	assert.equal(paragraphs.length, 1);
	assert.equal(getDescendants(paragraphs[0], 't')[0]?.textContent, 'Keep this paragraph');
	assert.equal(getDescendants(paragraphs[0], 'pPr')[0]?.getAttribute('algn'), 'ctr');
	assert.equal(getDescendants(paragraphs[0], 'rPr')[0]?.getAttribute('sz'), '3200');

	assert.equal(hasEmptyDrawingParagraphBefore(shape, 0), false);
	assert.equal(removeEmptyDrawingParagraphBefore(shape, 0), false);
});

test('removeEmptyDrawingParagraphBefore keeps a soft break and non-empty text', async () => {
	const { parseXml, hasEmptyDrawingParagraphBefore } = await loadDrawingMlTextModule();
	for (const content of [
		'<a:br/>',
		'<a:r><a:t>Not empty</a:t></a:r>',
	]) {
		const shape = parseXml(`
			<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${content}<a:endParaRPr/></a:p><a:p><a:r><a:t>Current</a:t></a:r></a:p></p:txBody>
			</p:sp>
		`, 'retain-content-paragraph.xml').documentElement;
		assert.equal(hasEmptyDrawingParagraphBefore(shape, 1), false);
	}
});

test('mergeDrawingParagraphWithPrevious joins non-empty paragraphs without flattening their runs', async () => {
	const { getDescendants, getDrawingParagraphs, getDrawingRunText, mergeDrawingParagraphWithPrevious, parseXml } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/>
				<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2600"/><a:t>Heading</a:t></a:r><a:endParaRPr/></a:p>
				<a:p><a:pPr algn="r"/><a:r><a:rPr b="1"/><a:t> body</a:t></a:r><a:br/><a:r><a:rPr i="1"/><a:t>tail</a:t></a:r><a:endParaRPr/></a:p>
			</p:txBody>
		</p:sp>
	`, 'merge-preceding-paragraph.xml').documentElement;

	assert.deepEqual(mergeDrawingParagraphWithPrevious(shape, 1), { merged: true, caretOffset: 7 });
	const paragraphs = getDrawingParagraphs(shape);
	assert.equal(paragraphs.length, 1);
	assert.equal(getDescendants(paragraphs[0], 'pPr')[0]?.getAttribute('algn'), 'ctr');
	assert.deepEqual(getDescendants(paragraphs[0], 'r').map(getDrawingRunText), ['Heading', ' body', 'tail']);
	assert.equal(getDescendants(paragraphs[0], 'br').length, 1);
	assert.equal(getDescendants(paragraphs[0], 'rPr')[1]?.getAttribute('b'), '1');
	assert.equal(getDescendants(paragraphs[0], 'rPr')[2]?.getAttribute('i'), '1');
	assert.deepEqual(mergeDrawingParagraphWithPrevious(shape, 0), { merged: false, caretOffset: 0 });
});

test('deleteDrawingTextRanges joins a cross-paragraph selection without flattening run formatting', async () => {
	const { deleteDrawingTextRanges, getDescendants, getDrawingParagraphs, getDrawingRunText, parseXml } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/>
				<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2600"/><a:t>Alpha </a:t></a:r><a:r><a:rPr b="1"/><a:t>beta</a:t></a:r><a:endParaRPr/></a:p>
				<a:p><a:r><a:rPr i="1"/><a:t>Gamma </a:t></a:r><a:r><a:rPr sz="1800"/><a:t>delta</a:t></a:r><a:endParaRPr/></a:p>
			</p:txBody>
		</p:sp>
	`, 'delete-cross-paragraph-range.xml').documentElement;

	const result = deleteDrawingTextRanges(shape, [
		{ paragraphIndex: 0, start: 6, end: 10 },
		{ paragraphIndex: 1, start: 0, end: 6 },
	]);

	assert.deepEqual(result, {
		changed: true,
		paragraphIndex: 0,
		caretOffset: 6,
		deletedRangeCount: 2,
		removedParagraphCount: 1,
		mergedParagraphs: true,
	});
	const paragraphs = getDrawingParagraphs(shape);
	assert.equal(paragraphs.length, 1);
	const runs = getDescendants(paragraphs[0], 'r').filter((run) => getDrawingRunText(run) !== '');
	assert.equal(getDrawingRunText(runs[0]), 'Alpha ');
	assert.equal(getDrawingRunText(runs[1]), 'delta');
	assert.equal(getDescendants(runs[1], 'rPr')[0]?.getAttribute('sz'), '1800');
	assert.equal(getDescendants(paragraphs[0], 'pPr')[0]?.getAttribute('algn'), 'ctr');
});

test('replaceDrawingParagraphs does not inherit heading bold onto later body paragraphs', async () => {
	const { getDescendants, parseXml, replaceDrawingParagraphs } = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody><a:bodyPr/><a:lstStyle/>
				<a:p><a:r><a:rPr lang="en-US" sz="3200" b="1"/><a:t>Heading</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>
			</p:txBody>
		</p:sp>
	`, 'bold-heading-template.xml').documentElement;

	replaceDrawingParagraphs(shape, [
		{ text: 'ASAP and the Fluent', listStyle: 'none', bold: true },
		{ text: 'Body copy about autonomous experiments.', listStyle: 'none' },
		{ text: 'A bullet about Fluent scripts.', listStyle: 'bullet' },
	]);

	const paragraphs = getDescendants(shape, 'p');
	assert.equal(paragraphs.length, 3);
	assert.equal(getDescendants(paragraphs[0], 'rPr')[0]?.getAttribute('b'), '1');
	assert.equal(getDescendants(paragraphs[1], 'rPr')[0]?.getAttribute('b'), '0');
	assert.equal(getDescendants(paragraphs[2], 'rPr')[0]?.getAttribute('b'), '0');
	assert.equal(getDescendants(paragraphs[0], 't')[0]?.textContent, 'ASAP and the Fluent');
	assert.equal(getDescendants(paragraphs[1], 't')[0]?.textContent, 'Body copy about autonomous experiments.');
});

test('Enter at end of ZWSP-only paragraph keeps styled anchor on source', async () => {
	const {
		EMPTY_PARAGRAPH_RENDER_ANCHOR,
		getDescendants,
		getDrawingParagraphs,
		getDrawingRunText,
		getDrawingRuns,
		parseXml,
		splitDrawingParagraphAtOffset,
	} = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:r><a:rPr lang="en-US" sz="2100" typeface="Calibri"/><a:t>${EMPTY_PARAGRAPH_RENDER_ANCHOR}</a:t></a:r>
					<a:endParaRPr lang="en-US" sz="2100"/>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'zwsp-only-paragraph.xml').documentElement;

	const beforeRuns = getDrawingRuns(getDrawingParagraphs(shape)[0]);
	const ooxmlLen = beforeRuns.reduce((sum, run) => sum + getDrawingRunText(run).length, 0);
	assert.equal(ooxmlLen, 1);

	const inserted = splitDrawingParagraphAtOffset(shape, 0, ooxmlLen);
	assert.equal(inserted, 1);
	const paragraphs = getDrawingParagraphs(shape);
	assert.equal(paragraphs.length, 2);
	assert.equal(getDrawingRunText(getDrawingRuns(paragraphs[0])[0]), EMPTY_PARAGRAPH_RENDER_ANCHOR);
	assert.equal(getDescendants(getDrawingRuns(paragraphs[0])[0], 'rPr')[0]?.getAttribute('sz'), '2100');
	assert.equal(getDrawingRunText(getDrawingRuns(paragraphs[1])[0]), EMPTY_PARAGRAPH_RENDER_ANCHOR);
});

test('Enter at EOF seeds empty suffix from caret run, not first run', async () => {
	const {
		EMPTY_PARAGRAPH_RENDER_ANCHOR,
		getDescendants,
		getDrawingParagraphs,
		getDrawingRunText,
		getDrawingRuns,
		parseXml,
		splitDrawingParagraphAtOffset,
	} = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:r>
						<a:rPr lang="en-US" sz="1800"><a:latin typeface="Arial"/></a:rPr>
						<a:t>Hello </a:t>
					</a:r>
					<a:r>
						<a:rPr lang="en-US" sz="2800"><a:latin typeface="Georgia"/></a:rPr>
						<a:t>World</a:t>
					</a:r>
					<a:endParaRPr lang="en-US" sz="1800"/>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'multi-run-eof-split.xml').documentElement;

	const source = getDrawingParagraphs(shape)[0];
	const ooxmlLen = getDrawingRuns(source).reduce((sum, run) => sum + getDrawingRunText(run).length, 0);
	assert.equal(ooxmlLen, 11);

	assert.equal(splitDrawingParagraphAtOffset(shape, 0, ooxmlLen), 1);
	const paragraphs = getDrawingParagraphs(shape);
	assert.equal(paragraphs.length, 2);
	assert.equal(getDrawingRunText(getDrawingRuns(paragraphs[0])[0]), 'Hello ');
	assert.equal(getDrawingRunText(getDrawingRuns(paragraphs[0])[1]), 'World');
	const suffixRun = getDrawingRuns(paragraphs[1])[0];
	assert.equal(getDrawingRunText(suffixRun), EMPTY_PARAGRAPH_RENDER_ANCHOR);
	assert.equal(getDescendants(suffixRun, 'rPr')[0]?.getAttribute('sz'), '2800');
	assert.equal(getDescendants(suffixRun, 'latin')[0]?.getAttribute('typeface'), 'Georgia');
	const suffixEnd = getDescendants(paragraphs[1], 'endParaRPr')[0];
	assert.equal(suffixEnd?.getAttribute('sz'), '2800');
	assert.equal(getDescendants(suffixEnd, 'latin')[0]?.getAttribute('typeface'), 'Georgia');
});

test('font summary includes ea and endParaRPr typefaces', async () => {
	const {
		getDrawingParagraphFontSummary,
		getDrawingParagraphs,
		parseXml,
	} = await loadDrawingMlTextModule();
	const shape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:r>
						<a:rPr lang="en-US" sz="2100"><a:ea typeface="Yu Gothic"/></a:rPr>
						<a:t>文</a:t>
					</a:r>
					<a:endParaRPr lang="en-US" sz="2100"><a:latin typeface="Calibri"/></a:endParaRPr>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'font-summary-ea-end.xml').documentElement;

	const summary = getDrawingParagraphFontSummary(getDrawingParagraphs(shape)[0]);
	assert.deepEqual(summary?.fontSizesPt, [21]);
	assert.ok(summary?.fontFamilies.includes('Yu Gothic'));
	assert.ok(summary?.fontFamilies.includes('Calibri'));
});

test('font summary reads rPr typeface attribute and pPr/defRPr', async () => {
	const {
		getDrawingParagraphFontSummary,
		getDrawingParagraphs,
		parseXml,
	} = await loadDrawingMlTextModule();
	const attrShape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:r><a:rPr lang="en-US" sz="2100" typeface="Calibri"/><a:t>Hi</a:t></a:r>
					<a:endParaRPr lang="en-US" sz="2100"/>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'font-summary-rpr-attr.xml').documentElement;
	assert.ok(getDrawingParagraphFontSummary(getDrawingParagraphs(attrShape)[0])?.fontFamilies.includes('Calibri'));

	const defShape = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:pPr><a:defRPr sz="1800"><a:latin typeface="Georgia"/></a:defRPr></a:pPr>
					<a:r><a:rPr lang="en-US" sz="2100"/><a:t>Hi</a:t></a:r>
					<a:endParaRPr lang="en-US" sz="2100"/>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'font-summary-defRPr.xml').documentElement;
	const defSummary = getDrawingParagraphFontSummary(getDrawingParagraphs(defShape)[0]);
	assert.ok(defSummary?.fontFamilies.includes('Georgia'));
	assert.deepEqual(defSummary?.fontSizesPt, [18, 21]);

	const szOnly = parseXml(`
		<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
			xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<p:txBody>
				<a:bodyPr/><a:lstStyle/>
				<a:p>
					<a:r><a:rPr lang="en-US" sz="2100"/><a:t>Hi</a:t></a:r>
					<a:endParaRPr lang="en-US" sz="2100"/>
				</a:p>
			</p:txBody>
		</p:sp>
	`, 'font-summary-sz-only.xml').documentElement;
	const szSummary = getDrawingParagraphFontSummary(getDrawingParagraphs(szOnly)[0]);
	assert.deepEqual(szSummary?.fontFamilies, []);
	assert.deepEqual(szSummary?.fontSizesPt, [21]);
});
