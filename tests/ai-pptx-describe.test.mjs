import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { loadPresentationEngineModule } from './helpers/load-plugin-modules.mjs';
import { readDeck, toArrayBuffer } from './helpers/renderer.mjs';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedPptxDescribeModule;

async function loadPptxDescribeModule() {
	if (cachedPptxDescribeModule) return cachedPptxDescribeModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-pptx-describe-test-'));
	const obsidianModuleDirectory = path.join(outputDirectory, 'node_modules', 'obsidian');
	await mkdir(obsidianModuleDirectory, { recursive: true });
	await writeFile(path.join(obsidianModuleDirectory, 'package.json'), '{"name":"obsidian","main":"index.js"}');
	await writeFile(
		path.join(obsidianModuleDirectory, 'index.js'),
		'class TFile { constructor(path, extension) { this.path = path; this.extension = extension; } } module.exports = { TFile };',
	);
	const outfile = path.join(outputDirectory, 'pptx-describe.cjs');
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
		entryPoints: [path.join(projectRoot, 'src/ai/pptxDescribe.ts')],
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
	cachedPptxDescribeModule = require(outfile);
	return cachedPptxDescribeModule;
}

function findTextShape(snapshot) {
	for (const slide of snapshot.slides) {
		for (const shape of slide.shapes) {
			if (shape.editable && shape.paragraphs?.length) {
				return { slide, shape };
			}
		}
	}
	return null;
}

test('describePptxFromEngine includes runs, style, chart data, crop, and slide background', async () => {
	const { describePptxFromEngine } = await loadPptxDescribeModule();
	const { PresentationEngine } = await loadPresentationEngineModule();
	const featuresBytes = toArrayBuffer(await readDeck('features.pptx'));
	const engine = await PresentationEngine.load(featuresBytes);
	const snapshot = describePptxFromEngine(engine, 'deck/features.pptx');

	assert.equal(snapshot.format, 'pptx');
	assert.ok(snapshot.slides.length > 0);

	const textTarget = findTextShape(snapshot);
	assert.ok(textTarget, 'expected an editable textbox with paragraphs');
	const paragraph = textTarget.shape.paragraphs[0];
	assert.ok('listStyle' in paragraph, 'expected native list state on textbox paragraphs');
	assert.ok(paragraph.runs?.length, 'expected per-run IDs on textbox paragraphs');
	assert.match(paragraph.runs[0].id, /\/r:\d+$/);
	assert.ok(textTarget.shape.style, 'expected shape fill/stroke style snapshot');

	for (const slide of snapshot.slides) {
		for (const shape of slide.shapes) {
			if (shape.index < 0) {
				assert.equal(shape.editable, false, `layout shape ${shape.id} must be non-editable`);
				assert.equal(shape.style, null);
			}
		}
	}

	const chartShape = snapshot.slides
		.flatMap((slide) => slide.shapes)
		.find((shape) => shape.kind === 'chart' && shape.editable);
	if (chartShape) {
		assert.ok(chartShape.chartData, 'expected chart data on editable chart shapes');
		assert.ok(Array.isArray(chartShape.chartData.series));
	}

	const imageShape = snapshot.slides
		.flatMap((slide) => slide.shapes)
		.find((shape) => shape.kind === 'image' && shape.editable);
	if (imageShape) {
		assert.ok(imageShape.crop !== undefined, 'expected crop metadata on editable images');
	}

	const slideWithBackground = snapshot.slides.find((slide) => slide.background);
	if (slideWithBackground?.background) {
		assert.ok(
			slideWithBackground.background.colorHex
			|| slideWithBackground.background.imageHref
			|| slideWithBackground.background.crop,
		);
	}
});
