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

async function createDocxBuffer(parts) {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>');
	zip.file(
		'_rels/.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
	for (const [partPath, xml] of Object.entries(parts)) {
		zip.file(partPath, xml);
	}
	return zip.generateAsync({ type: 'arraybuffer' });
}

function wrapBody(...inner) {
	return [
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
		...inner,
		'</w:body></w:document>',
	].join('');
}

let cachedDocxServiceModule;
let cachedObsidianStub;

function createMockVault(initialFiles) {
	const { TFile } = cachedObsidianStub;
	const store = new Map(initialFiles);
	return {
		store,
		getAbstractFileByPath(filePath) {
			if (!store.has(filePath)) return null;
			return new TFile(filePath, filePath.split('.').pop() ?? '');
		},
		async readBinary(file) {
			const bytes = store.get(file.path);
			if (!bytes) throw new Error(`Missing file: ${file.path}`);
			return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		},
		async modifyBinary(file, buffer) {
			store.set(file.path, Buffer.from(buffer));
		},
	};
}

async function loadDocxServiceModule() {
	if (cachedDocxServiceModule) return cachedDocxServiceModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-apply-test-'));
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
	);
	cachedObsidianStub = require(path.join(obsidianModuleDirectory, 'index.js'));
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
	cachedDocxServiceModule = require(outfile);
	return cachedDocxServiceModule;
}

let cachedDescribeModule;

async function loadDocxDescribeModule() {
	if (cachedDescribeModule) return cachedDescribeModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-apply-describe-'));
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
	cachedDescribeModule = require(outfile);
	return cachedDescribeModule;
}

test('DocxDocumentService applies setRunText and replaceText headlessly', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const docPath = 'notes/agent.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>Second line</w:t></w:r></w:p>',
		),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{
			op: 'docx.setRunText',
			blockId: 'body/p[0]',
			runId: 'body/p[0]/r[0]',
			text: 'Hello agent',
		},
		{
			op: 'docx.replaceText',
			query: 'line',
			replacement: 'paragraph',
		},
	]);
	assert.equal(applyResult.ok, true);
	assert.ok(applyResult.changed?.includes('body/p[0]/r[0]'));

	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blocks[0]?.text, 'Hello agent');
	assert.equal(snapshot.blocks[1]?.text, 'Second paragraph');
});

test('DocxDocumentService dryRun does not persist edits', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const docPath = 'notes/dry-run.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Keep me</w:t></w:r></w:p>'),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const dryRun = await service.apply(
		docPath,
		[{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'Changed' }],
		{ dryRun: true },
	);
	assert.equal(dryRun.ok, true);

	const unchanged = await describeDocxFromBuffer(initialBuffer, docPath);
	assert.equal(unchanged.blocks[0]?.text, 'Keep me');
});

test('DocxDocumentService inserts table after anchor block', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const docPath = 'notes/table.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Before</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>After</w:t></w:r></w:p>',
		),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.insertTable', afterBlockId: 'body/p[0]', rows: 2, cols: 2 },
	]);
	assert.equal(applyResult.ok, true);
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blockCount, 3);
	assert.equal(snapshot.blocks[0]?.text, 'Before');
	assert.equal(snapshot.blocks[1]?.kind, 'table');
	assert.equal(snapshot.blocks[1]?.rows, 2);
	assert.equal(snapshot.blocks[1]?.cols, 2);
	assert.equal(snapshot.blocks[2]?.text, 'After');
});

test('DocxDocumentService syncs open DOCX view after agent apply and save', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/open-view.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Before agent</w:t></w:r></w:p>'),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	let viewBuffer = initialBuffer.slice(0);
	const reloads = [];
	const mockView = {
		getLoadedDocumentPath: () => docPath,
		canAgentEdit: () => true,
		exportBufferForAgent: async () => viewBuffer.slice(0),
		reloadFromAgentBuffer: async (buffer) => {
			reloads.push(buffer.byteLength);
			viewBuffer = buffer.slice(0);
		},
		saveCurrentDocument: async () => true,
	};

	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => mockView,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'After agent' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	assert.equal(reloads.length, 1);

	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));
	assert.equal(reloads.length, 2);

	const savedZip = await JSZip.loadAsync(vault.store.get(docPath).buffer);
	const savedXml = await savedZip.file('word/document.xml').async('string');
	assert.match(savedXml, /After agent/);
});

test('DocxDocumentService undo restores headless agent edits', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/undo.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Before undo</w:t></w:r></w:p>'),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'After undo' },
	]);
	assert.equal(applyResult.ok, true);
	assert.equal(applyResult.canUndo, true);

	const undoResult = await service.undo(docPath);
	assert.equal(undoResult.ok, true);

	const lease = await service['sessions'].acquire(docPath);
	const xml = lease.patch.getDocumentXml();
	assert.match(xml, /Before undo/);
	assert.doesNotMatch(xml, /After undo/);
});

test('describeDocxFromBuffer exposes image blocks and replaceImage uses drawing paragraph ids', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const MINIMAL_PNG = Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
		'base64',
	);
	const imageParagraph = [
		'<w:p>',
		'<w:r><w:drawing>',
		'<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">',
		'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
		'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
		'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">',
		'<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/></pic:blipFill>',
		'</pic:pic></a:graphicData></a:graphic></wp:inline>',
		'</w:drawing></w:r></w:p>',
	].join('');

	const docPath = 'notes/image.docx';
	const imagePath = 'assets/replace.png';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Before image</w:t></w:r></w:p>',
			imageParagraph,
		),
		'word/_rels/document.xml.rels': [
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
			'<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
			'</Relationships>',
		].join(''),
		'word/media/image1.png': MINIMAL_PNG,
	});

	const vault = createMockVault(new Map([
		[docPath, Buffer.from(initialBuffer)],
		[imagePath, MINIMAL_PNG],
	]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const snapshot = await describeDocxFromBuffer(initialBuffer, docPath);
	assert.equal(snapshot.scope.sources.includes('word/document.xml'), true);
	assert.equal(snapshot.scope.writeExcluded.includes('comments'), true);
	assert.equal(snapshot.blocks[0]?.kind, 'paragraph');
	assert.equal(snapshot.blocks[1]?.kind, 'image');
	assert.equal(snapshot.blocks[1]?.id, 'body/p[1]');
	assert.equal(snapshot.blocks[1]?.relationshipId, 'rId5');
	assert.equal(snapshot.blocks[1]?.mediaPath, 'word/media/image1.png');
	assert.equal(snapshot.blocks[1]?.runs, undefined);

	const replacePng = Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAACb/HwjAAAADUlEQVR42mP8z5+3HgAFfwJ/5m6p5QAAAABJRU5ErkJggg==',
		'base64',
	);
	vault.store.set(imagePath, replacePng);

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.replaceImage', blockId: 'body/p[1]', vaultImagePath: imagePath },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedZip = await JSZip.loadAsync(vault.store.get(docPath).buffer);
	const savedImage = await savedZip.file('word/media/image1.png').async('nodebuffer');
	assert.equal(savedImage.equals(replacePng), true);
});

test('apply can edit header and footer runs via stable ids', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();

	const docPath = 'notes/parts.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Body</w:t></w:r></w:p>'),
		'word/headers/header1.xml': [
			'<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
			'<w:p><w:r><w:t>Old header</w:t></w:r></w:p>',
			'</w:hdr>',
		].join(''),
		'word/footers/footer1.xml': [
			'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
			'<w:p><w:r><w:t>Old footer</w:t></w:r></w:p>',
			'</w:ftr>',
		].join(''),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.setRunText', blockId: 'header/1/p[0]', runId: 'header/1/p[0]/r[0]', text: 'New header' },
		{ op: 'docx.setRunText', blockId: 'footer/1/p[0]', runId: 'footer/1/p[0]/r[0]', text: 'New footer' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedZip = await JSZip.loadAsync(vault.store.get(docPath).buffer);
	const headerXml = await savedZip.file('word/headers/header1.xml').async('string');
	const footerXml = await savedZip.file('word/footers/footer1.xml').async('string');
	assert.match(headerXml, /New header/);
	assert.match(footerXml, /New footer/);
	assert.doesNotMatch(headerXml, /Old header/);
	assert.doesNotMatch(footerXml, /Old footer/);
});

test('DocxDocumentService applies insertText, deleteRange, hyperlink, and paragraph break ops', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const docPath = 'notes/phase2.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>'),
	});

	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [
		{ op: 'docx.insertText', blockId: 'body/p[0]', offset: 5, text: ' brave' },
		{
			op: 'docx.insertHyperlink',
			range: {
				start: { blockId: 'body/p[0]', offset: 6 },
				end: { blockId: 'body/p[0]', offset: 11 },
			},
			url: 'https://example.com',
			tooltip: 'Example',
		},
		{
			op: 'docx.removeHyperlink',
			range: {
				start: { blockId: 'body/p[0]', offset: 6 },
				end: { blockId: 'body/p[0]', offset: 11 },
			},
		},
		{ op: 'docx.insertParagraphBreak', blockId: 'body/p[0]', offset: 6 },
		{
			op: 'docx.deleteRange',
			range: {
				start: { blockId: 'body/p[1]', offset: 0 },
				end: { blockId: 'body/p[1]', offset: 5 },
			},
		},
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const savedZip = await JSZip.loadAsync(savedBytes.buffer);
	const documentXml = await savedZip.file('word/document.xml').async('string');
	const relsXml = await savedZip.file('word/_rels/document.xml.rels').async('string');
	assert.match(documentXml, /Hello/);
	assert.match(documentXml, /world/);
	assert.match(relsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"/);
	assert.match(relsXml, /Target="https:\/\/example\.com"/);
	assert.doesNotMatch(documentXml, /<w:hyperlink\b[^>]*>[\s\S]*?<w:t>brave<\/w:t>/);

	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blockCount, 2);
	assert.equal(snapshot.blocks[0]?.text, 'Hello ');
	assert.equal(snapshot.blocks[1]?.text, ' world');
});
