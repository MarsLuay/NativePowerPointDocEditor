import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { docxEditorAliases } from './helpers/docx-esbuild-aliases.mjs';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

async function loadBlankPackagesModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-blank-office-'));
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(path.join(obsidianModuleDirectory, 'index.js'), 'module.exports = {};');

	const outfile = path.join(outputDirectory, 'blankOfficePackages.cjs');
	await build({
		alias: docxEditorAliases,
		entryPoints: [path.join(projectRoot, 'src/vault/blankOfficePackages.ts')],
		bundle: true,
		platform: 'node',
		format: 'cjs',
		outfile,
		external: ['obsidian'],
		logLevel: 'silent',
	});
	return require(outfile);
}

test('blank DOCX package has required parts and styles', async () => {
	const { buildBlankDocxArrayBuffer } = await loadBlankPackagesModule();
	const buffer = await buildBlankDocxArrayBuffer();
	const zip = await JSZip.loadAsync(buffer);
	assert.ok(zip.file('[Content_Types].xml'));
	assert.ok(zip.file('_rels/.rels'));
	assert.ok(zip.file('word/document.xml'));
	assert.ok(zip.file('word/styles.xml'));
	assert.ok(zip.file('word/_rels/document.xml.rels'));
	const documentXml = await zip.file('word/document.xml').async('string');
	assert.match(documentXml, /<w:body>/);
	const stylesXml = await zip.file('word/styles.xml').async('string');
	assert.match(stylesXml, /w:ascii="Arial"/);
	assert.match(stylesXml, /w:hAnsi="Arial"/);
	assert.doesNotMatch(stylesXml, /w:ascii="Calibri"/);
});

test('blank PPTX package has one blank slide plus master/layout/theme', async () => {
	const { buildBlankPptxArrayBuffer } = await loadBlankPackagesModule();
	const buffer = await buildBlankPptxArrayBuffer();
	const zip = await JSZip.loadAsync(buffer);
	assert.ok(zip.file('[Content_Types].xml'));
	assert.ok(zip.file('ppt/presentation.xml'));
	assert.ok(zip.file('ppt/slides/slide1.xml'));
	assert.ok(zip.file('ppt/slideMasters/slideMaster1.xml'));
	assert.ok(zip.file('ppt/slideLayouts/slideLayout1.xml'));
	assert.ok(zip.file('ppt/theme/theme1.xml'));
	const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
	assert.match(slideXml, /<p:spTree>/);
	assert.doesNotMatch(slideXml, /<p:sp>/);
});
