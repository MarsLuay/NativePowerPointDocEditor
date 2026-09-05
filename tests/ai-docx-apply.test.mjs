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

test('DocxDocumentService replaces all table cell content while preserving cell properties', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const docPath = 'notes/multi-paragraph-cell.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Before</w:t></w:r></w:p>',
			'<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Template title</w:t></w:r></w:p><w:p><w:r><w:t>Template instructions</w:t></w:r></w:p><w:p><w:r><w:drawing/></w:r></w:p></w:tc></w:tr></w:tbl>',
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
		{ op: 'docx.setCellText', cellId: 'body/tbl[0]/tr[0]/tc[0]', text: 'Updated & verified' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blocks[1]?.kind, 'table');
	assert.equal(snapshot.blocks[1]?.cells?.[0]?.text, 'Updated & verified');

	const output = await JSZip.loadAsync(savedBytes.buffer);
	const documentXml = await output.file('word/document.xml')?.async('string');
	assert.match(documentXml ?? '', /<w:tcPr><w:tcW w:w="2400" w:type="dxa"\/><\/w:tcPr>/);
	assert.match(documentXml ?? '', /<w:t>Updated &amp; verified<\/w:t>/);
	assert.doesNotMatch(documentXml ?? '', /Template title|Template instructions|<w:drawing\b/);
	assert.equal((documentXml?.match(/<w:p\b/g) ?? []).length, 3);
});

test('DocxDocumentService deletes only the requested table', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const docPath = 'notes/delete-table.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Before</w:t></w:r></w:p>',
			'<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Remove me</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
			'<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>References</w:t></w:r></w:p>',
		),
	});
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(docPath, [{ op: 'docx.deleteTable', tableId: 'body/tbl[0]' }]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blockCount, 2);
	assert.equal(snapshot.blocks[0]?.text, 'Before');
	assert.equal(snapshot.blocks[1]?.text, 'References');

	const output = await JSZip.loadAsync(savedBytes.buffer);
	const documentXml = await output.file('word/document.xml')?.async('string');
	assert.doesNotMatch(documentXml ?? '', /Remove me|<w:tbl\b/);
	assert.match(documentXml ?? '', /<w:pPr><w:keepNext\/><\/w:pPr><w:r><w:rPr><w:b\/><\/w:rPr><w:t>References<\/w:t>/);
});

test('DocxDocumentService applies every setRunText operation in a single batch', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();
	const docPath = 'notes/multi-run-batch.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>First</w:t></w:r><w:r><w:t> middle</w:t></w:r><w:r><w:t> last</w:t></w:r></w:p>',
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
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'Updated' },
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[1]', text: '' },
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[2]', text: '' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blocks[0]?.text, 'Updated');

	const output = await JSZip.loadAsync(savedBytes.buffer);
	const documentXml = await output.file('word/document.xml')?.async('string');
	assert.match(documentXml ?? '', /<w:t>Updated<\/w:t>/);
	assert.doesNotMatch(documentXml ?? '', /middle|last/);
	assert.equal((documentXml?.match(/<w:r\b/g) ?? []).length, 3);
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

test('DocxDocumentService makes empty-run writes explicit and describes unsaved headless edits', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/empty-run.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t></w:t></w:r></w:p>',
			'<w:p><w:r><w:t>Following paragraph</w:t></w:r></w:p>',
		),
	});
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const falseSuccess = await service.apply(docPath, [{
		op: 'docx.setRunText',
		blockId: 'body/p[0]',
		runId: 'body/p[0]/r[0]',
		text: 'Inserted',
	}]);
	assert.equal(falseSuccess.ok, false);
	assert.equal(falseSuccess.errors[0]?.code, 'EMPTY_RUN_USE_INSERT_TEXT');

	const inserted = await service.apply(docPath, [{
		op: 'docx.insertText',
		blockId: 'body/p[0]',
		offset: 0,
		text: 'Inserted',
	}]);
	assert.equal(inserted.ok, true, JSON.stringify(inserted.errors));

	const described = await service.describe(docPath);
	assert.equal(described.ok, true, JSON.stringify(described.errors));
	assert.equal(described.snapshot.blocks[0]?.text, 'Inserted');
	assert.equal(described.snapshot.blocks[1]?.text, 'Following paragraph');
});

test('DocxDocumentService inserts stable list paragraphs and reports created ids', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/list-insert.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p w14:paraId="AAAAAAA1" w14:textId="11111111"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Existing bullet</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>Following paragraph</w:t></w:r></w:p>',
		),
	});
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const breakResult = await service.apply(docPath, [{
		op: 'docx.insertParagraphBreak',
		blockId: 'body/p[0]',
		offset: 0,
	}]);
	assert.equal(breakResult.ok, true, JSON.stringify(breakResult.errors));
	assert.deepEqual(breakResult.created, ['body/p[1]']);
	assert.equal(breakResult.preview?.[0]?.after?.inheritedListProperties, true);

	const afterBreak = await service.describe(docPath);
	assert.equal(afterBreak.ok, true, JSON.stringify(afterBreak.errors));
	assert.equal(afterBreak.snapshot.blocks[0]?.text, '');
	assert.equal(afterBreak.snapshot.blocks[1]?.text, 'Existing bullet');
	assert.equal(afterBreak.snapshot.blockCount, 3);

	const insertResult = await service.apply(docPath, [{
		op: 'docx.insertParagraphsAfter',
		afterBlockId: 'body/p[1]',
		paragraphs: [{ text: 'Added bullet', listStyle: 'bullet', bold: true }],
	}]);
	assert.equal(insertResult.ok, true, JSON.stringify(insertResult.errors));
	assert.deepEqual(insertResult.created, ['body/p[2]']);
	assert.equal(insertResult.preview?.[0]?.after?.inheritedListProperties, true);

	const described = await service.describe(docPath);
	assert.equal(described.ok, true, JSON.stringify(described.errors));
	assert.equal(described.snapshot.blocks[2]?.text, 'Added bullet');
	assert.equal(described.snapshot.blocks[2]?.runs?.[0]?.bold, true);

	await service.save(docPath);
	const savedBytes = vault.store.get(docPath);
	const savedXml = await (await JSZip.loadAsync(savedBytes.buffer)).file('word/document.xml').async('string');
	assert.equal((savedXml.match(/<w:numPr\b/g) ?? []).length, 3);
	assert.equal((savedXml.match(/w14:paraId="AAAAAAA1"/g) ?? []).length, 1);
	assert.equal((savedXml.match(/w14:paraId="[0-9A-F]+"/g) ?? []).length, 3);
});

test('DocxDocumentService deletes a blank paragraph without merging the next heading', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { describeDocxFromBuffer } = await loadDocxDescribeModule();

	const docPath = 'notes/delete-blank.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Before</w:t></w:r></w:p>',
			'<w:p/>',
			'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>TECHNICAL SKILLS</w:t></w:r></w:p>',
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
		{ op: 'docx.deleteBlock', blockId: 'body/p[1]' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedBytes = vault.store.get(docPath);
	const snapshot = await describeDocxFromBuffer(
		savedBytes.buffer.slice(savedBytes.byteOffset, savedBytes.byteOffset + savedBytes.byteLength),
		docPath,
	);
	assert.equal(snapshot.blockCount, 2);
	assert.equal(snapshot.blocks[0]?.text, 'Before');
	assert.equal(snapshot.blocks[1]?.text, 'TECHNICAL SKILLS');

	const savedZip = await JSZip.loadAsync(savedBytes.buffer);
	const documentXml = await savedZip.file('word/document.xml').async('string');
	assert.match(documentXml, /<w:p><w:r><w:rPr><w:b\/><\/w:rPr><w:t>TECHNICAL SKILLS<\/w:t><\/w:r><\/w:p>/);
});

test('DocxDocumentService sets a paragraph bottom border without replacing heading content', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();

	const docPath = 'notes/heading-rule.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:pPr><w:spacing w:line="252" w:lineRule="auto"/></w:pPr><w:r><w:t>TECHNICAL SKILLS</w:t></w:r></w:p>',
			'<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
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
			op: 'docx.setParagraphBottomBorder',
			blockId: 'body/p[0]',
			border: { style: 'single', size: 6, space: 1, color: '000000' },
		},
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedZip = await JSZip.loadAsync(vault.store.get(docPath).buffer);
	const documentXml = await savedZip.file('word/document.xml').async('string');
	assert.match(documentXml, /<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"\/><\/w:pBdr>/);
	assert.match(documentXml, /<w:t>TECHNICAL SKILLS<\/w:t>/);
	assert.match(documentXml, /<w:spacing w:line="252" w:lineRule="auto"\/>/);
});

test('DocxDocumentService replaceText does not damage unrelated DOCX structural attributes', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();

	const docPath = 'notes/replace-preserve.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t xml:space="preserve">w:t preserve</w:t></w:r></w:p>',
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
			op: 'docx.replaceText',
			query: 'preserve',
			replacement: 'kept',
		},
		{
			op: 'docx.replaceText',
			query: 'w:t',
			replacement: 'v:t',
		}
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	await service.save(docPath);

	const savedZip = await JSZip.loadAsync(vault.store.get(docPath).buffer);
	const documentXml = await savedZip.file('word/document.xml').async('string');
	assert.match(documentXml, /<w:t xml:space="preserve">v:t kept<\/w:t>/);
});
