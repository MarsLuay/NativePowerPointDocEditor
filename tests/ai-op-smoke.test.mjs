import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { loadPowerPointPackageModule, loadPresentationEngineModule } from './helpers/load-plugin-modules.mjs';
import { readDeck, toArrayBuffer } from './helpers/renderer.mjs';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

const TRANSFORM = { x: 914_400, y: 685_800, cx: 2_743_200, cy: 2_057_400, rot: 0 };
const MINIMAL_PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
	'base64',
);

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

let cachedObsidianStub;
let cachedDocxServiceModule;
let cachedDocxAgentReloadGuardModule;
let cachedPptxAiModules;

async function setupObsidianStub(outputDirectory) {
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		[
			'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } }',
			'const normalizePath = (value) => String(value).replace(/\\\\/g, "/").replace(/\\/{2,}/g, "/");',
			'module.exports = { TFile, normalizePath };',
		].join('\n'),
	);
	return require(path.join(obsidianModuleDirectory, 'index.js'));
}

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
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-smoke-docx-'));
	cachedObsidianStub = await setupObsidianStub(outputDirectory);
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

async function loadDocxAgentReloadGuardModule() {
	if (cachedDocxAgentReloadGuardModule) return cachedDocxAgentReloadGuardModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-reload-guard-'));
	const outfile = path.join(outputDirectory, 'docx-agent-reload-guard.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/docx/DocxAgentReloadGuard.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedDocxAgentReloadGuardModule = require(outfile);
	return cachedDocxAgentReloadGuardModule;
}

async function loadPptxAiModules() {
	if (cachedPptxAiModules) return cachedPptxAiModules;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-smoke-pptx-'));
	cachedObsidianStub = await setupObsidianStub(outputDirectory);
	const executorOut = path.join(outputDirectory, 'pptx-op-executor.cjs');
	const wasmPath = path.join(projectRoot, 'node_modules/pptx-svg/dist/main.wasm');
	const wasmPlugin = {
		name: 'inline-pptx-svg-wasm',
		setup(buildContext) {
			buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async ({ path: modulePath }) => {
				const { readFile: read } = await import('node:fs/promises');
				let contents = await read(modulePath, 'utf8');
				contents = contents.replace(
					"const DEFAULT_WASM_URL = new URL('./main.wasm', import.meta.url).href;",
					'const DEFAULT_WASM_URL = undefined;',
				);
				return { contents, loader: 'js' };
			});
		},
	};
	await build({
		absWorkingDir: outputDirectory,
		entryPoints: [path.join(projectRoot, 'src/ai/pptxOpExecutor.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile: executorOut,
		platform: 'node',
		target: 'node22',
		external: ['obsidian'],
		loader: { '.wasm': 'binary' },
		plugins: [wasmPlugin],
	});
	cachedPptxAiModules = {
		...require(executorOut),
		wasm: await readFile(wasmPath),
	};
	return cachedPptxAiModules;
}

function discoverPptxTargets(engine) {
	const slideIndex = 0;
	engine.renderSlide(slideIndex);
	let textShape = 0;
	let imageShape = 0;
	let chartShape = 0;
	let fillShape = 0;
	const shapeIndices = [];
	for (let shapeIndex = 0; shapeIndex < 32; shapeIndex++) {
		try {
			const svg = engine.renderShape(slideIndex, shapeIndex);
			if (!svg) continue;
			shapeIndices.push(shapeIndex);
		if (svg.includes('<text')) textShape = shapeIndex;
		if (engine.isImageShape(slideIndex, shapeIndex)) imageShape = shapeIndex;
		if (svg.includes('data-ooxml-shape-type="chart"')) chartShape = shapeIndex;
		if (engine.canSetShapeFillColor(slideIndex, shapeIndex)) fillShape = shapeIndex;
		} catch {
			break;
		}
	}
	assert.ok(shapeIndices.length >= 2, 'features.pptx must expose at least two shapes on slide 0');
	return {
		textShape,
		imageShape,
		chartShape,
		fillShape,
		shapeIndices,
		transform: TRANSFORM,
	};
}

test('DOCX agent ops smoke all implemented operations', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/smoke.docx';
	const imagePath = 'assets/smoke.png';
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
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Alpha old</w:t></w:r></w:p>',
			'<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
			imageParagraph,
		),
		'docProps/core.xml': [
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">',
			'<dc:creator>Template Author</dc:creator><cp:lastModifiedBy>Template Editor</cp:lastModifiedBy><cp:revision>7</cp:revision>',
			'</cp:coreProperties>',
		].join(''),
		'word/_rels/document.xml.rels': [
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
			'<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
			'</Relationships>',
		].join(''),
		'word/media/image1.png': MINIMAL_PNG,
	});

	const vault = createMockVault(
		new Map([
			[docPath, Buffer.from(initialBuffer)],
			[imagePath, MINIMAL_PNG],
		]),
	);
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const ops = [
		{ op: 'docx.removeComments' },
		{ op: 'docx.setCoreProperties', creator: 'Marwan & Luay', lastModifiedBy: 'Marwan Luay' },
		{ op: 'docx.setRunText', blockId: 'body/p[0]', runId: 'body/p[0]/r[0]', text: 'Alpha new' },
		{ op: 'docx.setRunStyle', runId: 'body/p[0]/r[0]', style: { bold: true } },
		{ op: 'docx.setParagraphStyle', blockId: 'body/p[0]', style: { name: 'Heading2' } },
		{ op: 'docx.insertText', blockId: 'body/p[0]', offset: 5, text: ' extra' },
		{ op: 'docx.deleteRange', range: { start: { blockId: 'body/p[0]', offset: 0 }, end: { blockId: 'body/p[0]', offset: 1 } } },
		{ op: 'docx.insertHyperlink', range: { start: { blockId: 'body/p[0]', offset: 1 }, end: { blockId: 'body/p[0]', offset: 4 } }, url: 'https://example.com' },
		{ op: 'docx.removeHyperlink', range: { start: { blockId: 'body/p[0]', offset: 1 }, end: { blockId: 'body/p[0]', offset: 4 } } },
		{ op: 'docx.setCellText', cellId: 'body/tbl[0]/tr[0]/tc[0]', text: 'Updated cell' },
		{ op: 'docx.setCellStyle', cellId: 'body/tbl[0]/tr[0]/tc[0]', style: { name: 'Normal' } },
		{ op: 'docx.replaceImage', blockId: 'body/p[1]', vaultImagePath: imagePath },
		{ op: 'docx.insertTable', afterBlockId: 'body/p[0]', rows: 1, cols: 2 },
		{ op: 'docx.insertImage', afterBlockId: 'body/p[0]', vaultImagePath: imagePath },
		{ op: 'docx.replaceText', query: 'old', replacement: 'fresh' },
		{ op: 'docx.insertParagraphBreak', blockId: 'body/p[0]', offset: 3 },
	];

	const applyResult = await service.apply(docPath, ops);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));
	assert.ok(vault.store.get(docPath)?.byteLength > 0);
	const output = await JSZip.loadAsync(vault.store.get(docPath));
	const coreProperties = await output.file('docProps/core.xml')?.async('string');
	assert.match(coreProperties ?? '', /<dc:creator>Marwan &amp; Luay<\/dc:creator>/);
	assert.match(coreProperties ?? '', /<cp:lastModifiedBy>Marwan Luay<\/cp:lastModifiedBy>/);
	assert.match(coreProperties ?? '', /<cp:revision>7<\/cp:revision>/);
});

test('DOCX replaceBodyParagraphs rewrites body while preserving sectPr', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'thank-you-body.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody(
			'<w:p><w:r><w:t>Old</w:t></w:r></w:p>',
			'<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
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
			op: 'docx.replaceBodyParagraphs',
			paragraphs: ['Marwan Luay', '', 'Thank you for the scholarship.'],
		},
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));

	const output = await JSZip.loadAsync(vault.store.get(docPath));
	const documentXml = await output.file('word/document.xml')?.async('string') ?? '';
	assert.match(documentXml, /Marwan Luay/);
	assert.match(documentXml, /Thank you for the scholarship\./);
	assert.match(documentXml, /<w:sectPr>/);
	assert.doesNotMatch(documentXml, /<w:tbl\b/);
	assert.doesNotMatch(documentXml, />Old</);
	const bodyInner = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(documentXml)?.[1] ?? '';
	const topLevelParagraphs = bodyInner.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
	assert.equal(topLevelParagraphs.length, 3, documentXml);
});

test('DOCX metadata dry runs leave the cached session unchanged', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const docPath = 'notes/dry-run-metadata.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Unchanged</w:t></w:r></w:p>'),
		'docProps/core.xml': [
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">',
			'<dc:creator>Template Author</dc:creator><cp:lastModifiedBy>Template Editor</cp:lastModifiedBy>',
			'</cp:coreProperties>',
		].join(''),
	});
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const service = new DocxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenDocxView: () => null,
		findOpenPptxView: () => null,
	});

	const applyResult = await service.apply(
		docPath,
		[{ op: 'docx.setCoreProperties', creator: 'Marwan Luay', lastModifiedBy: 'Marwan Luay' }],
		{ dryRun: true },
	);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));
	const output = await JSZip.loadAsync(vault.store.get(docPath));
	const coreProperties = await output.file('docProps/core.xml')?.async('string');
	assert.match(coreProperties ?? '', /<dc:creator>Template Author<\/dc:creator>/);
	assert.match(coreProperties ?? '', /<cp:lastModifiedBy>Template Editor<\/cp:lastModifiedBy>/);
});

test('DOCX agent reload guard retains the latest package until its matching editor session is ready', async (t) => {
	const { DocxAgentReloadGuard } = await loadDocxAgentReloadGuardModule();
	const hadWindow = Object.prototype.hasOwnProperty.call(globalThis, 'window');
	const originalWindow = globalThis.window;
	t.after(() => {
		if (hadWindow) {
			globalThis.window = originalWindow;
		} else {
			delete globalThis.window;
		}
	});
	if (!globalThis.window) {
		globalThis.window = {
			setTimeout: globalThis.setTimeout.bind(globalThis),
			clearTimeout: globalThis.clearTimeout.bind(globalThis),
		};
	}
	const guard = new DocxAgentReloadGuard();
	const first = new Uint8Array([1, 2, 3]).buffer;
	const second = new Uint8Array([4, 5, 6]).buffer;
	const firstIdentity = { documentSession: 8, filePath: 'notes/first.docx' };
	const secondIdentity = { documentSession: 9, filePath: 'notes/first.docx' };

	guard.begin(firstIdentity, first);
	assert.deepEqual([...new Uint8Array(guard.getPendingBuffer(firstIdentity))], [1, 2, 3]);
	assert.equal(
		guard.complete({ ...firstIdentity, documentSession: 7 }),
		false,
		'a stale editor-ready callback must not release the package',
	);
	assert.deepEqual([...new Uint8Array(guard.getPendingBuffer(firstIdentity))], [1, 2, 3]);

	guard.begin(secondIdentity, second);
	assert.equal(guard.complete(firstIdentity), false, 'a superseded session must not release the newer package');
	assert.deepEqual([...new Uint8Array(guard.getPendingBuffer(secondIdentity))], [4, 5, 6]);
	const ready = guard.waitForReady(secondIdentity, 100);
	assert.equal(guard.complete(secondIdentity), true);
	await ready;
	assert.equal(guard.getPendingBuffer(secondIdentity), null);
	const latest = guard.getLatestBufferAfter(0, secondIdentity);
	assert.ok(latest);
	assert.deepEqual([...new Uint8Array(latest.buffer)], [4, 5, 6]);

	const thirdIdentity = { documentSession: 10, filePath: 'notes/first.docx' };
	guard.begin(thirdIdentity, second);
	await assert.rejects(guard.waitForReady(thirdIdentity, 1), /did not become ready within 1ms/);
	assert.equal(guard.getPendingBuffer(thirdIdentity), null, 'a timed-out guard must not mask future saves');

	guard.clear();
	assert.equal(guard.getPendingBuffer(secondIdentity), null);
	assert.equal(guard.getLatestBufferAfter(0, secondIdentity), null);
});

test('DOCX metadata survives immediate view-mode agent save while React still has the old editor buffer', async () => {
	const { DocxDocumentService } = await loadDocxServiceModule();
	const { DocxAgentReloadGuard } = await loadDocxAgentReloadGuardModule();
	const docPath = 'notes/open-view-metadata.docx';
	const initialBuffer = await createDocxBuffer({
		'word/document.xml': wrapBody('<w:p><w:r><w:t>Visible text stays unchanged</w:t></w:r></w:p>'),
		'docProps/core.xml': [
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">',
			'<dc:creator>Cozzi, Matt</dc:creator><cp:lastModifiedBy>Cozzi, Matt</cp:lastModifiedBy>',
			'</cp:coreProperties>',
		].join(''),
	});
	const vault = createMockVault(new Map([[docPath, Buffer.from(initialBuffer)]]));
	const reloadGuard = new DocxAgentReloadGuard();
	let documentSession = 0;
	let reloadIdentity = null;
	const staleEditorBuffer = initialBuffer.slice(0);
	const mockView = {
		getLoadedDocumentPath: () => docPath,
		canAgentEdit: () => true,
		exportBufferForAgent: async () => reloadIdentity
			? reloadGuard.getPendingBuffer(reloadIdentity) ?? staleEditorBuffer.slice(0)
			: staleEditorBuffer.slice(0),
		reloadFromAgentBuffer: async (buffer) => {
			documentSession += 1;
			reloadIdentity = { documentSession, filePath: docPath };
			reloadGuard.begin(reloadIdentity, buffer);
			// This deliberately leaves staleEditorBuffer unchanged, matching the
			// interval before React remounts the editor for the new document key.
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
		{ op: 'docx.setCoreProperties', creator: 'Marwan Luay', lastModifiedBy: 'Marwan Luay' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));

	const saveResult = await service.save(docPath);
	assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));
	const output = await JSZip.loadAsync(vault.store.get(docPath));
	const coreProperties = await output.file('docProps/core.xml')?.async('string');
	const documentXml = await output.file('word/document.xml')?.async('string');
	assert.match(coreProperties ?? '', /<dc:creator>Marwan Luay<\/dc:creator>/);
	assert.match(coreProperties ?? '', /<cp:lastModifiedBy>Marwan Luay<\/cp:lastModifiedBy>/);
	assert.match(documentXml ?? '', /Visible text stays unchanged/);
});

test('PPTX agent ops smoke dispatches every operation', async (t) => {
	const { executePptxOp } = await loadPptxAiModules();
	const { PresentationEngine } = await loadPresentationEngineModule();
	const featuresBytes = toArrayBuffer(await readDeck('features.pptx'));
	const fixtureImage = toArrayBuffer(await readDeck('features.pptx'));
	const zip = await JSZip.loadAsync(fixtureImage);
	const imageEntry = zip.file('ppt/media/image1.png');
	assert.ok(imageEntry, 'features.pptx must ship ppt/media/image1.png');
	const imageBytes = new Uint8Array(await imageEntry.async('arraybuffer'));

	const vault = createMockVault(new Map([['assets/smoke.png', Buffer.from(imageBytes)]]));
	const probeEngine = await PresentationEngine.load(featuresBytes);
	const targets = discoverPptxTargets(probeEngine);

	const builders = {
		'pptx.updateShapeText': () => ({ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: targets.textShape, text: 'Smoke title' }),
		'pptx.replaceShapeParagraphs': () => ({
			op: 'pptx.replaceShapeParagraphs',
			slideIndex: 0,
			shapeIndex: targets.textShape,
			paragraphs: [
				{ text: 'Smoke heading', listStyle: 'none' },
				{ text: 'Native smoke bullet', listStyle: 'bullet' },
			],
		}),
		'pptx.updateParagraphText': () => ({ op: 'pptx.updateParagraphText', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, text: 'Paragraph' }),
		'pptx.updateTextRun': () => ({ op: 'pptx.updateTextRun', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, runIndex: 0, text: 'Run' }),
		'pptx.replaceText': () => ({ op: 'pptx.replaceText', query: 'Smoke', replacement: 'Test' }),
		'pptx.setRunStyle': () => ({ op: 'pptx.setRunStyle', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, runIndex: 0, style: { bold: true } }),
		'pptx.setParagraphAlignment': () => ({ op: 'pptx.setParagraphAlignment', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, align: 'ctr' }),
		'pptx.applyListStyle': () => ({ op: 'pptx.applyListStyle', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, style: 'bullet' }),
		'pptx.updateTransform': () => ({ op: 'pptx.updateTransform', slideIndex: 0, shapeIndex: targets.textShape, transform: targets.transform }),
		'pptx.setShapeFillColor': () => ({ op: 'pptx.setShapeFillColor', slideIndex: 0, shapeIndex: targets.fillShape, hex: '#1B75BB' }),
		'pptx.reorderShapes': () => ({ op: 'pptx.reorderShapes', slideIndex: 0, shapeIndex: targets.textShape, mode: 'forward' }),
		'pptx.groupShapes': () => ({
			op: 'pptx.groupShapes',
			slideIndex: 0,
			shapeIndices: [targets.shapeIndices[0], targets.shapeIndices[1]],
		}),
		'pptx.ungroupShapes': async (eng) => {
			const grouped = await eng.groupShapes(0, [targets.shapeIndices[0], targets.shapeIndices[1]]);
			return { op: 'pptx.ungroupShapes', slideIndex: 0, shapeIndex: grouped };
		},
		'pptx.flipShape': () => ({ op: 'pptx.flipShape', slideIndex: 0, shapeIndex: targets.textShape, axis: 'horizontal' }),
		'pptx.deleteShape': () => ({ op: 'pptx.deleteShape', slideIndex: 0, shapeIndex: targets.shapeIndices[0] }),
		'pptx.addImage': () => ({ op: 'pptx.addImage', slideIndex: 0, vaultImagePath: 'assets/smoke.png', transform: TRANSFORM }),
		'pptx.addShape': () => ({ op: 'pptx.addShape', slideIndex: 0, geometry: 'rect', transform: TRANSFORM }),
		'pptx.addTextBox': () => ({ op: 'pptx.addTextBox', slideIndex: 0, transform: TRANSFORM }),
		'pptx.addTable': () => ({ op: 'pptx.addTable', slideIndex: 0, rows: 2, cols: 2, transform: TRANSFORM }),
		'pptx.addChart': () => ({ op: 'pptx.addChart', slideIndex: 0, transform: TRANSFORM }),
		'pptx.addSlide': () => ({ op: 'pptx.addSlide', afterIndex: 0, layout: 'blank' }),
		'pptx.deleteSlide': async (eng) => {
			await eng.addSlide(0);
			return { op: 'pptx.deleteSlide', slideIndex: 1 };
		},
		'pptx.moveSlide': async (eng) => {
			await eng.addSlide(0);
			return { op: 'pptx.moveSlide', slideIndex: 1, direction: -1 };
		},
		'pptx.duplicateSlide': () => ({ op: 'pptx.duplicateSlide', slideIndex: 0 }),
		'pptx.reorderSlides': async (eng) => {
			await eng.addSlide(0);
			return { op: 'pptx.reorderSlides', order: [1, 0] };
		},
		'pptx.setSlideBackground': () => ({ op: 'pptx.setSlideBackground', slideIndex: 0, colorHex: 'F0F0F0' }),
		'pptx.setImageCrop': () => ({ op: 'pptx.setImageCrop', slideIndex: 0, shapeIndex: targets.imageShape, crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 } }),
		'pptx.resetImage': async (eng) => {
			await eng.setImageCrop(0, targets.imageShape, { left: 0.1, top: 0, right: 0, bottom: 0 });
			return { op: 'pptx.resetImage', slideIndex: 0, shapeIndex: targets.imageShape };
		},
		'pptx.replaceImage': () => ({ op: 'pptx.replaceImage', slideIndex: 0, shapeIndex: targets.imageShape, vaultImagePath: 'assets/smoke.png' }),
		'pptx.replaceImageFromShape': () => ({ op: 'pptx.replaceImageFromShape', slideIndex: 0, shapeIndex: targets.imageShape, sourceSlideIndex: 0, sourceShapeIndex: targets.imageShape }),
		'pptx.updateChartData': () => ({
			op: 'pptx.updateChartData',
			slideIndex: 0,
			shapeIndex: targets.chartShape,
			data: { categories: ['A', 'B'], series: [{ name: 'Series 1', values: [3, 4] }] },
		}),
	};

	for (const [opId] of Object.entries(builders)) {
		await t.test(opId, async () => {
			const freshEngine = await PresentationEngine.load(featuresBytes);
			const op = await Promise.resolve(builders[opId](freshEngine));
			const dryRun = opId === 'pptx.updateChartData';
			const result = await executePptxOp(
				{
					engine: freshEngine,
					vault,
					filePath: 'features.pptx',
					dryRun,
				},
				op,
			);
			assert.ok(result.changedIds.length > 0, `${opId} should report changed ids`);
			if (!dryRun) {
				const exported = await freshEngine.export();
				await PresentationEngine.validateRoundTrip(exported, freshEngine.slideCount);
			}
		});
	}
});

test('PPTX image reset reports no-op state and clears authored crop', async () => {
	const { PresentationEngine } = await loadPresentationEngineModule();
	const source = toArrayBuffer(await readDeck('features.pptx'));
	const engine = await PresentationEngine.load(source);
	const { imageShape } = discoverPptxTargets(engine);

	assert.deepEqual(engine.getImageResetState(0, imageShape), { hasCrop: false, effectNames: [] });
	const noOp = await engine.resetImage(0, imageShape);
	assert.deepEqual(noOp, { changed: false, cropRemoved: false, hasCrop: false, effectNames: [] });

	await engine.setImageCrop(0, imageShape, { left: 12.5, top: 0, right: 0, bottom: 0 });
	assert.deepEqual(engine.getImageResetState(0, imageShape), { hasCrop: true, effectNames: [] });
	const reset = await engine.resetImage(0, imageShape);
	assert.deepEqual(reset, { changed: true, cropRemoved: true, hasCrop: true, effectNames: [] });
	assert.deepEqual(engine.getImageResetState(0, imageShape), { hasCrop: false, effectNames: [] });

	const exported = await engine.export();
	await PresentationEngine.validateRoundTrip(exported, engine.slideCount);
	const reloaded = await PresentationEngine.load(exported);
	assert.deepEqual(reloaded.getImageResetState(0, imageShape), { hasCrop: false, effectNames: [] });
});

test('PPTX agent list paragraphs are native and updateShapeText rejects list-like newlines', async () => {
	const { executePptxOp } = await loadPptxAiModules();
	const { PresentationEngine } = await loadPresentationEngineModule();
	const featuresBytes = toArrayBuffer(await readDeck('features.pptx'));
	const engine = await PresentationEngine.load(featuresBytes);
	const targets = discoverPptxTargets(engine);
	const vault = createMockVault(new Map());
	const context = {
		engine,
		vault,
		filePath: 'features.pptx',
		dryRun: false,
	};

	const result = await executePptxOp(context, {
		op: 'pptx.replaceShapeParagraphs',
		slideIndex: 0,
		shapeIndex: targets.textShape,
		paragraphs: [
			{ text: 'Current support', listStyle: 'none' },
			{ text: 'Native bullet one', listStyle: 'bullet' },
			{ text: 'Native bullet two', listStyle: 'bullet' },
		],
	});
	assert.ok(result.changedIds.length > 0);
	assert.equal(engine.getParagraphRunText(0, targets.textShape, 1), 'Native bullet one');
	assert.equal(engine.getParagraphListStyle(0, targets.textShape, 1), 'bullet');

	const exported = await engine.export();
	const zip = await JSZip.loadAsync(exported);
	const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('string');
	assert.ok(slideXml);
	assert.equal((slideXml.match(/<a:buChar\b[^>]*\bchar="•"/g) || []).length, 2);
	assert.doesNotMatch(slideXml, /<a:t>[^<]*•/);

	await assert.rejects(
		executePptxOp(context, {
			op: 'pptx.updateShapeText',
			slideIndex: 0,
			shapeIndex: targets.textShape,
			text: 'This would be one paragraph\nwith a fake list item',
		}),
		(error) => error?.code === 'SCHEMA_INVALID' && /replaceShapeParagraphs/.test(error.message),
	);
});

test('PPTX image deletion persists only with its explicit validation allowance', async () => {
	const { PresentationEngine } = await loadPresentationEngineModule();
	const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
	const source = toArrayBuffer(await readDeck('features.pptx'));
	const engine = await PresentationEngine.load(source);
	const { imageShape } = discoverPptxTargets(engine);

	await engine.deleteShape(0, imageShape);
	const exported = await engine.export();
	const unallowed = await validatePowerPointExportContents(source, exported);
	assert.equal(unallowed.ok, false);
	assert.ok(unallowed.errors.some((error) => error.includes('image') || error.includes('media')));

	const allowed = await validatePowerPointExportContents(source, exported, {
		allowedMarkerRemovals: engine.getProtectedSlideMarkerRemovalAllowance(),
		allowedPartRemovals: engine.getPrunedPackageParts(),
	});
	assert.equal(allowed.ok, true, JSON.stringify(allowed.errors));
	const reloaded = await PresentationEngine.load(exported);
	assert.notEqual(reloaded.isImageShape(0, imageShape), true);
});

test('PPTX chart/table/group deletion records matching protected-marker allowances', async () => {
	const { PresentationEngine } = await loadPresentationEngineModule();
	const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
	const source = toArrayBuffer(await readDeck('features.pptx'));
	const cases = [
		{ label: 'chart', marker: 'chart', pick: (svg) => svg.includes('data-ooxml-shape-type="chart"') },
		{ label: 'table', marker: 'table', pick: (svg) => svg.includes('data-ooxml-shape-type="table"') },
		{ label: 'group', marker: 'groupedShape', pick: (svg) => svg.includes('data-ooxml-shape-type="group"') },
	];

	for (const testCase of cases) {
		const engine = await PresentationEngine.load(source);
		engine.renderSlide(0);
		let shapeIndex = -1;
		for (let index = 0; index < 32; index += 1) {
			try {
				const svg = engine.renderShape(0, index);
				if (svg && testCase.pick(svg)) {
					shapeIndex = index;
					break;
				}
			} catch {
				break;
			}
		}
		assert.ok(shapeIndex >= 0, `${testCase.label}: expected a deletable shape on features.pptx`);

		await engine.deleteShape(0, shapeIndex);
		const exported = await engine.export();
		const unallowed = await validatePowerPointExportContents(source, exported);
		assert.equal(unallowed.ok, false, `${testCase.label}: expected validation to require an allowance`);
		assert.ok(
			unallowed.errors.some((error) => error.includes(testCase.marker)),
			`${testCase.label}: expected ${testCase.marker} drop error, got ${JSON.stringify(unallowed.errors)}`,
		);

		const allowance = engine.getProtectedSlideMarkerRemovalAllowance();
		assert.ok(
			(allowance[testCase.marker] ?? 0) > 0,
			`${testCase.label}: expected ${testCase.marker} allowance, got ${JSON.stringify(allowance)}`,
		);
		const allowed = await validatePowerPointExportContents(source, exported, {
			allowedMarkerRemovals: allowance,
		});
		assert.equal(allowed.ok, true, `${testCase.label}: ${JSON.stringify(allowed.errors)}`);
	}
});

async function loadCreateOfficeDocumentModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-create-doc-'));
	await setupObsidianStub(outputDirectory);
	const outfile = path.join(outputDirectory, 'create-office-document.cjs');
	await build({
		absWorkingDir: outputDirectory,
		entryPoints: [path.join(projectRoot, 'src/ai/createOfficeDocument.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		external: ['obsidian'],
	});
	return require(outfile);
}

function createMutableVault() {
	const { TFile } = cachedObsidianStub;
	const store = new Map();
	const folders = new Set();
	return {
		store,
		folders,
		getAbstractFileByPath(filePath) {
			if (folders.has(filePath)) return { path: filePath };
			if (!store.has(filePath)) return null;
			return new TFile(filePath, filePath.split('.').pop() ?? '');
		},
		async createFolder(folderPath) {
			folders.add(folderPath);
		},
		async createBinary(filePath, buffer) {
			store.set(filePath, Buffer.from(buffer));
			return new TFile(filePath, filePath.split('.').pop() ?? '');
		},
		async modifyBinary(file, buffer) {
			store.set(file.path, Buffer.from(buffer));
		},
		async readBinary(file) {
			const bytes = store.get(file.path);
			if (!bytes) throw new Error(`Missing file: ${file.path}`);
			return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
		},
	};
}

test('createOfficeDocument writes a blank DOCX with optional paragraphs', async () => {
	cachedObsidianStub = await setupObsidianStub(await mkdtemp(path.join(tmpdir(), 'npde-ai-create-stub-')));
	const { createOfficeDocument } = await loadCreateOfficeDocumentModule();
	const vault = createMutableVault();
	const pathName = 'School/Applications and Fafsa/Letter.docx';
	const created = await createOfficeDocument(vault, {
		path: pathName,
		kind: 'docx',
		paragraphs: ['Line one', '', 'Line three'],
	});
	assert.equal(created.ok, true, JSON.stringify(created.errors));
	assert.equal(created.path, pathName);
	assert.ok(vault.store.has(pathName));
	assert.ok(vault.folders.has('School'));
	assert.ok(vault.folders.has('School/Applications and Fafsa'));

	const zip = await JSZip.loadAsync(vault.store.get(pathName));
	const documentXml = await zip.file('word/document.xml')?.async('string');
	assert.match(documentXml ?? '', /Line one/);
	assert.match(documentXml ?? '', /Line three/);
	assert.match(documentXml ?? '', /<w:sectPr\b/);

	const blocked = await createOfficeDocument(vault, {
		path: pathName,
		kind: 'docx',
		paragraphs: ['Nope'],
	});
	assert.equal(blocked.ok, false);
	assert.equal(blocked.errors[0]?.code, 'FILE_EXISTS');
});
