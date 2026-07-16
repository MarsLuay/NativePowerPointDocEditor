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
				"export { setDrawingText } from './src/powerpoint/drawingmlText.ts';",
				"export { setDrawingParagraphText } from './src/powerpoint/drawingmlText.ts';",
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
