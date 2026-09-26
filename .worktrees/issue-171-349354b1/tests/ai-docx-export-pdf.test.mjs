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
	name: 'stub-docx-export-deps',
	setup(buildContext) {
		buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: 'obsidian', namespace: 'stub' }));
		buildContext.onResolve({ filter: /^\.\.\/export\/artifactPaths$/ }, (args) => ({ path: args.path, namespace: 'stub' }));
		buildContext.onResolve({ filter: /^\.\.\/logger$/ }, (args) => ({ path: args.path, namespace: 'stub' }));
		buildContext.onLoad({ filter: /.*/, namespace: 'stub' }, (args) => {
			if (args.path === 'obsidian') {
				return {
					contents: `
export class TFile { constructor(init = {}) { Object.assign(this, init); } }
export const normalizePath = (value) => value.replace(/\\\\/g, '/').replace(/\\/{2,}/g, '/');
globalThis.__NPDE_TEST_TFile = TFile;
`,
					loader: 'js',
				};
			}
			if (args.path.endsWith('artifactPaths')) {
				return {
					contents: `
export const getAvailableNumberedPath = (requested, exists) => { for (let i = 2; ; i += 1) { const candidate = requested.replace(/\\.pdf$/i, ' ' + i + '.pdf'); if (!exists(candidate)) return candidate; } };
export const getVaultFolderPrefix = (path) => path ? path + '/' : '';
export const sanitizeArtifactBaseName = (name) => name;
export const writeVaultBinaryArtifact = async () => ({ path: 'output.pdf' });
`,
					loader: 'js',
				};
			}
			return { contents: 'export const debugLog = () => {};', loader: 'js' };
		});
	},
};

async function loadExportModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-docx-export-pdf-'));
	const outfile = path.join(outputDirectory, 'docx-export-pdf.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/docxExportPdf.ts')],
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

test('DOCX PDF export target defaults beside source and honors conflict policy', async () => {
	const mod = await loadExportModule();
	const files = new Map();
	const vault = { getAbstractFileByPath: (candidate) => files.get(candidate) ?? null };
	const sourceFile = new globalThis.__NPDE_TEST_TFile({
		basename: 'Marwan Luay Resume',
		parent: { path: 'Life/Financials' },
	});

	const fresh = mod.resolveDocxExportPdfOutputTarget(vault, sourceFile, {});
	assert.equal(fresh.path, 'Life/Financials/Marwan Luay Resume.pdf');
	files.set(fresh.path, new globalThis.__NPDE_TEST_TFile({ path: fresh.path, extension: 'pdf' }));
	assert.equal(mod.resolveDocxExportPdfOutputTarget(vault, sourceFile, {}).path, 'Life/Financials/Marwan Luay Resume 2.pdf');
	assert.equal(mod.resolveDocxExportPdfOutputTarget(vault, sourceFile, { conflict: 'replace' }).replace, true);
});
