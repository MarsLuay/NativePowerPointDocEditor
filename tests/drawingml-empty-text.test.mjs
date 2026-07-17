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
				"export { deleteDrawingTextRanges } from './src/powerpoint/drawingmlText.ts';",
				"export { getDrawingParagraphs, getDrawingRunText } from './src/powerpoint/drawingmlText.ts';",
				"export { setDrawingText } from './src/powerpoint/drawingmlText.ts';",
				"export { setDrawingParagraphText } from './src/powerpoint/drawingmlText.ts';",
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
