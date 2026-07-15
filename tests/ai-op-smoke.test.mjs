import assert from 'node:assert/strict';
import { test } from 'node:test';
import JSZip from 'jszip';
import { build } from 'esbuild';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { loadPresentationEngineModule } from './helpers/load-plugin-modules.mjs';
import { getDocxRuntimeAliases } from './helpers/docx-runtime-aliases.mjs';
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
let cachedPptxAiModules;

async function setupObsidianStub(outputDirectory) {
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
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
		alias: await getDocxRuntimeAliases(projectRoot),
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
		alias: await getDocxRuntimeAliases(projectRoot),
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
	const shapeIndices = [];
	for (let shapeIndex = 0; shapeIndex < 32; shapeIndex++) {
		try {
			const svg = engine.renderShape(slideIndex, shapeIndex);
			if (!svg) continue;
			shapeIndices.push(shapeIndex);
			if (svg.includes('<text')) textShape = shapeIndex;
			if (engine.isImageShape(slideIndex, shapeIndex)) imageShape = shapeIndex;
			if (svg.includes('data-ooxml-shape-type="chart"')) chartShape = shapeIndex;
		} catch {
			break;
		}
	}
	assert.ok(shapeIndices.length >= 2, 'features.pptx must expose at least two shapes on slide 0');
	return {
		textShape,
		imageShape,
		chartShape,
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
		'pptx.updateParagraphText': () => ({ op: 'pptx.updateParagraphText', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, text: 'Paragraph' }),
		'pptx.updateTextRun': () => ({ op: 'pptx.updateTextRun', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, runIndex: 0, text: 'Run' }),
		'pptx.replaceText': () => ({ op: 'pptx.replaceText', query: 'Smoke', replacement: 'Test' }),
		'pptx.setRunStyle': () => ({ op: 'pptx.setRunStyle', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, runIndex: 0, style: { bold: true } }),
		'pptx.setParagraphAlignment': () => ({ op: 'pptx.setParagraphAlignment', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, align: 'ctr' }),
		'pptx.applyListStyle': () => ({ op: 'pptx.applyListStyle', slideIndex: 0, shapeIndex: targets.textShape, paragraphIndex: 0, style: 'bullet' }),
		'pptx.updateTransform': () => ({ op: 'pptx.updateTransform', slideIndex: 0, shapeIndex: targets.textShape, transform: targets.transform }),
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
		'pptx.resetImage': () => ({ op: 'pptx.resetImage', slideIndex: 0, shapeIndex: targets.imageShape }),
		'pptx.replaceImage': () => ({ op: 'pptx.replaceImage', slideIndex: 0, shapeIndex: targets.imageShape, vaultImagePath: 'assets/smoke.png' }),
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
