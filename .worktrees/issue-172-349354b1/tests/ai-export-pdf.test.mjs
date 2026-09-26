import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

const stubPlugin = {
	name: 'stub-heavy-deps',
	setup(buildContext) {
		buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: 'obsidian', namespace: 'stub' }));
		buildContext.onResolve({ filter: /^\.\.\/(PresentationEngine|PowerPointExport|SvgSecurity|logger)$/ }, (args) => ({
			path: args.path,
			namespace: 'stub',
		}));
		buildContext.onResolve({ filter: /^\.\.\/powerpoint\/svgUtils$/ }, (args) => ({
			path: args.path,
			namespace: 'stub',
		}));
		buildContext.onLoad({ filter: /.*/, namespace: 'stub' }, (args) => {
			if (args.path === 'obsidian') {
				return {
					contents: `
export class TFile {
  constructor(init = {}) {
    Object.assign(this, init);
  }
}
export const normalizePath = (value) => value.replace(/\\\\/g, '/').replace(/\\/{2,}/g, '/');
globalThis.__NPDE_TEST_TFile = TFile;
`,
					loader: 'js',
				};
			}
			return {
				contents: `
export const exportSlidesToPdf = async () => new ArrayBuffer(0);
export const SLIDE_PDF_EXPORT_SCALE = 2;
export const SLIDE_PDF_EXPORT_DPI = 150;
export const EMU_PER_INCH = 914400;
export const slideSizeEmuToPdfPoints = (cx, cy) => ({
  width: Math.round((cx / 914400) * 72 * 100) / 100,
  height: Math.round((cy / 914400) * 72 * 100) / 100,
});
export const createSvgElementFromString = () => null;
export const sanitizeSvg = (svg) => ({ svg, issues: [] });
export const normalizeSvgForDisplay = () => {};
export const debugLog = () => {};
export class PresentationEngine {}
`,
				loader: 'js',
			};
		});
	},
};

async function loadExportModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-export-pdf-'));
	const outfile = path.join(outputDirectory, 'pptx-export-pdf.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/pptxExportPdf.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		plugins: [stubPlugin],
	});
	return require(outfile);
}

test('resolveExportPdfOutputTarget defaults beside source and keep-both on conflict', async () => {
	const mod = await loadExportModule();
	const files = new Map();
	const vault = {
		getAbstractFileByPath(candidate) {
			return files.get(candidate) ?? null;
		},
	};
	const sourceFile = {
		basename: '24x36 Poster Template',
		parent: { path: 'Job/PNNL/Deliverables/Poster Info' },
	};

	const fresh = mod.resolveExportPdfOutputTarget(vault, sourceFile, {});
	assert.equal(fresh.path, 'Job/PNNL/Deliverables/Poster Info/24x36 Poster Template.pdf');
	assert.equal(fresh.replace, false);

	files.set(fresh.path, new globalThis.__NPDE_TEST_TFile({ path: fresh.path, extension: 'pdf' }));
	const kept = mod.resolveExportPdfOutputTarget(vault, sourceFile, { conflict: 'keep-both' });
	assert.equal(kept.path, 'Job/PNNL/Deliverables/Poster Info/24x36 Poster Template 2.pdf');

	const replaced = mod.resolveExportPdfOutputTarget(vault, sourceFile, { conflict: 'replace' });
	assert.equal(replaced.path, fresh.path);
	assert.equal(replaced.replace, true);
});

test('resolveExportPdfOutputTarget rejects non-pdf outputPath', async () => {
	const mod = await loadExportModule();
	const vault = { getAbstractFileByPath: () => null };
	const sourceFile = { basename: 'deck', parent: { path: 'Job' } };
	assert.throws(
		() => mod.resolveExportPdfOutputTarget(vault, sourceFile, { outputPath: 'Job/deck.pptx' }),
		(error) => error?.code === 'SCHEMA_INVALID',
	);
});
