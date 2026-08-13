import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

async function createDocxBuffer(documentXml) {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>');
	zip.file(
		'_rels/.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
	zip.file('word/document.xml', documentXml);
	return zip.generateAsync({ type: 'arraybuffer' });
}

function wrapDocument(body) {
	return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
}

let cachedService;
let cachedDescribe;
let cachedObsidian;

async function loadService() {
	if (cachedService) return cachedService;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-layout-test-'));
	const obsidianDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianDirectory, { recursive: true });
	await writeFile(path.join(obsidianDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
	);
	cachedObsidian = require(path.join(obsidianDirectory, 'index.js'));
	const outfile = path.join(outputDirectory, 'docx-service.cjs');
	await build({
		absWorkingDir: outputDirectory,
		entryPoints: [path.join(projectRoot, 'src/ai/docxDocumentService.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		external: ['obsidian'],
	});
	cachedService = require(outfile);
	return cachedService;
}

async function loadDescribe() {
	if (cachedDescribe) return cachedDescribe;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-layout-describe-'));
	const outfile = path.join(outputDirectory, 'docx-describe.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/docxDescribe.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedDescribe = require(outfile);
	return cachedDescribe;
}

function createMockVault(initialFiles) {
	const { TFile } = cachedObsidian;
	const store = new Map(initialFiles);
	return {
		store,
		getAbstractFileByPath(filePath) {
			return store.has(filePath) ? new TFile(filePath, filePath.split('.').pop() ?? '') : null;
		},
		async readBinary(file) {
			const bytes = store.get(file.path);
			return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		},
		async modifyBinary(file, buffer) {
			store.set(file.path, Buffer.from(buffer));
		},
	};
}

test('DOCX AI layout operations expose and patch compact resume layout', async () => {
	const { DocxDocumentService } = await loadService();
	const { describeDocxFromBuffer } = await loadDescribe();
	const docPath = 'Life/Financials/resume-fixture.docx';
	const initialBuffer = await createDocxBuffer(wrapDocument([
		'<w:p><w:pPr><w:spacing w:after="40"/><w:rPr><w:rFonts w:ascii="Aptos" w:eastAsia="Arial"/><w:sz w:val="34"/><w:szCs w:val="34"/><w:lang w:val="en-US"/></w:rPr></w:pPr></w:p>',
		'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>',
	].join('')));
	const vault = createMockVault([[docPath, Buffer.from(initialBuffer)]]);
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const result = await service.apply(docPath, [
		{
			op: 'docx.setParagraphDefaultRunStyle',
			blockId: 'body/p[0]',
			style: { fontSizePt: 12 },
		},
		{
			op: 'docx.setParagraphLayout',
			blockId: 'body/p[0]',
			layout: {
				spacing: { before: 0, after: 0, line: 252, lineRule: 'auto' },
				indent: { left: 360, hanging: 360 },
				alignment: 'left',
			},
		},
		{
			op: 'docx.setSectionLayout',
			sectionIndex: 0,
			layout: { margins: { top: 432, right: 720, bottom: 432, left: 720 } },
		},
	]);
	assert.equal(result.ok, true, JSON.stringify(result.errors));
	assert.equal(result.preview?.[0]?.before?.fontSizePt, 17);
	await service.save(docPath);

	const saved = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength), docPath);
	assert.deepEqual(snapshot.blocks[0]?.layout?.spacing, { before: 0, after: 0, line: 252, lineRule: 'auto' });
	assert.deepEqual(snapshot.blocks[0]?.layout?.indent, { left: 360, hanging: 360 });
	assert.deepEqual(snapshot.blocks[0]?.defaultRunStyle, { fontFamily: 'Aptos', fontSizePt: 12 });
	assert.deepEqual(snapshot.sections[0]?.margins, { top: 432, right: 720, bottom: 432, left: 720 });

	const output = await JSZip.loadAsync(saved);
	const documentXml = await output.file('word/document.xml')?.async('string') ?? '';
	assert.match(documentXml, /<w:spacing w:after="0" w:before="0" w:line="252" w:lineRule="auto"\/>/);
	assert.match(documentXml, /<w:rFonts w:ascii="Aptos" w:eastAsia="Arial"\/>/);
	assert.match(documentXml, /<w:sz w:val="24"\/><w:szCs w:val="24"\/><w:lang w:val="en-US"\/>/);
});

test('DOCX AI layout operations preserve unrelated attributes and self-closing properties', async () => {
	const { DocxDocumentService } = await loadService();
	const docPath = 'Life/Financials/resume-layout-preservation.docx';
	const initialBuffer = await createDocxBuffer(wrapDocument([
		'<w:p><w:pPr/><w:r><w:t>Self-closing properties</w:t></w:r></w:p>',
		'<w:p><w:pPr><w:spacing w:after="40" w:beforeAutospacing="1"/><w:ind w:left="720" w:startChars="2"/></w:pPr><w:r><w:t>Existing properties</w:t></w:r></w:p>',
		'<w:sectPr w:rsidR="00AA" w:rsidSect="00BB"><w:pgSz w:w="12240" w:h="15840" w:code="keep"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:gutterAtTop="1"/></w:sectPr>',
	].join('')));
	const vault = createMockVault([[docPath, Buffer.from(initialBuffer)]]);
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const result = await service.apply(docPath, [
		{
			op: 'docx.setParagraphLayout',
			blockId: 'body/p[0]',
			layout: { spacing: { after: 0 } },
		},
		{
			op: 'docx.setParagraphLayout',
			blockId: 'body/p[1]',
			layout: { spacing: { before: 0 }, indent: { left: 360 } },
		},
		{
			op: 'docx.setSectionLayout',
			sectionIndex: 0,
			layout: {
				pageSize: { width: 12240, height: 15840, orient: 'portrait' },
				margins: { top: 432 },
			},
		},
	]);
	assert.equal(result.ok, true, JSON.stringify(result.errors));
	await service.save(docPath);

	const saved = vault.store.get(docPath);
	const output = await JSZip.loadAsync(saved);
	const documentXml = await output.file('word/document.xml')?.async('string') ?? '';
	assert.doesNotMatch(documentXml, /<w:pPr\s*\/>/);
	assert.match(documentXml, /<w:pPr><w:spacing w:after="0"\/><\/w:pPr><w:r>/);
	assert.match(documentXml, /<w:spacing\b(?=[^>]*w:after="40")(?=[^>]*w:beforeAutospacing="1")(?=[^>]*w:before="0")[^>]*\/>/);
	assert.match(documentXml, /<w:ind\b(?=[^>]*w:left="360")(?=[^>]*w:startChars="2")[^>]*\/>/);
	assert.match(documentXml, /<w:sectPr\b(?=[^>]*w:rsidR="00AA")(?=[^>]*w:rsidSect="00BB")[^>]*>/);
	assert.match(documentXml, /<w:pgSz\b(?=[^>]*w:w="12240")(?=[^>]*w:h="15840")(?=[^>]*w:orient="portrait")(?=[^>]*w:code="keep")[^>]*\/>/);
	assert.match(documentXml, /<w:pgMar\b(?=[^>]*w:top="432")(?=[^>]*w:right="720")(?=[^>]*w:gutterAtTop="1")[^>]*\/>/);
});
