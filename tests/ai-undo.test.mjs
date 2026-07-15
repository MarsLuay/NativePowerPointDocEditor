import assert from 'node:assert/strict';
import { test } from 'node:test';
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

let cachedObsidianStub;
let cachedPptxServiceModule;

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

async function loadPptxServiceModule() {
	if (cachedPptxServiceModule) return cachedPptxServiceModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-pptx-undo-test-'));
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
	);
	cachedObsidianStub = require(path.join(obsidianModuleDirectory, 'index.js'));
	const outfile = path.join(outputDirectory, 'pptx-service.cjs');
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
		entryPoints: [path.join(projectRoot, 'src/ai/pptxDocumentService.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		external: ['obsidian'],
		loader: { '.wasm': 'binary' },
		plugins: [wasmPlugin],
	});
	cachedPptxServiceModule = {
		...require(outfile),
		wasm: await readFile(wasmPath),
	};
	return cachedPptxServiceModule;
}

async function discoverTextShape(engine) {
	for (let shapeIndex = 0; shapeIndex < 32; shapeIndex++) {
		try {
			const svg = engine.renderShape(0, shapeIndex);
			if (svg?.includes('<text')) {
				return shapeIndex;
			}
		} catch {
			break;
		}
	}
	throw new Error('No text shape found on slide 0');
}

test('PptxDocumentService undo restores headless agent edits', async () => {
	const { PptxDocumentService } = await loadPptxServiceModule();
	const { PresentationEngine } = await loadPresentationEngineModule();

	const deckPath = 'deck/features.pptx';
	const featuresBytes = toArrayBuffer(await readDeck('features.pptx'));
	const probe = await PresentationEngine.load(featuresBytes);
	const textShape = await discoverTextShape(probe);
	const originalText = probe.getParagraphRunText(0, textShape, 0) ?? '';

	const vault = createMockVault(new Map([[deckPath, Buffer.from(featuresBytes)]]));
	const service = new PptxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenPptxView: () => null,
		findOpenDocxView: () => null,
	});

	const applyResult = await service.apply(deckPath, [
		{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: textShape, text: 'After headless undo' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));
	assert.equal(applyResult.canUndo, true);

	const leaseAfterApply = await service.sessions.acquire(deckPath);
	assert.equal(leaseAfterApply.engine.getParagraphRunText(0, textShape, 0), 'After headless undo');

	const undoResult = await service.undo(deckPath);
	assert.equal(undoResult.ok, true, JSON.stringify(undoResult.errors));

	const leaseAfterUndo = await service.sessions.acquire(deckPath);
	assert.equal(leaseAfterUndo.engine.getParagraphRunText(0, textShape, 0), originalText);
});

test('PptxDocumentService redo restores headless agent undo', async () => {
	const { PptxDocumentService } = await loadPptxServiceModule();
	const { PresentationEngine } = await loadPresentationEngineModule();

	const deckPath = 'deck/features.pptx';
	const featuresBytes = toArrayBuffer(await readDeck('features.pptx'));
	const probe = await PresentationEngine.load(featuresBytes);
	const textShape = await discoverTextShape(probe);
	const originalText = probe.getParagraphRunText(0, textShape, 0) ?? '';

	const vault = createMockVault(new Map([[deckPath, Buffer.from(featuresBytes)]]));
	const service = new PptxDocumentService({
		vault,
		normalizePath: (value) => value,
		findOpenPptxView: () => null,
		findOpenDocxView: () => null,
	});

	const applyResult = await service.apply(deckPath, [
		{ op: 'pptx.updateShapeText', slideIndex: 0, shapeIndex: textShape, text: 'Redo me' },
	]);
	assert.equal(applyResult.ok, true, JSON.stringify(applyResult.errors));

	const undoResult = await service.undo(deckPath);
	assert.equal(undoResult.ok, true, JSON.stringify(undoResult.errors));

	const leaseAfterUndo = await service.sessions.acquire(deckPath);
	assert.equal(leaseAfterUndo.engine.getParagraphRunText(0, textShape, 0), originalText);

	const redoResult = await service.redo(deckPath);
	assert.equal(redoResult.ok, true, JSON.stringify(redoResult.errors));

	const leaseAfterRedo = await service.sessions.acquire(deckPath);
	assert.equal(leaseAfterRedo.engine.getParagraphRunText(0, textShape, 0), 'Redo me');
});
