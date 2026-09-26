import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

async function loadArtifactPaths() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-artifact-paths-'));
	const outfile = path.join(outputDirectory, 'artifact-paths.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/export/artifactPaths.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		plugins: [{
			name: 'stub-obsidian',
			setup(buildContext) {
				buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: 'obsidian', namespace: 'stub' }));
				buildContext.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({
					contents: `
export class TFile {
  constructor(path) { this.path = path; }
}
export const normalizePath = (value) => value;
globalThis.__NPDE_ARTIFACT_TFILE = TFile;
`,
					loader: 'js',
				}));
			},
		}],
	});
	return require(outfile);
}

test('writeVaultBinaryArtifact reports the requested path when the vault returns null', async () => {
	const mod = await loadArtifactPaths();
	const outputPath = 'exports/poster.pdf';
	let createCalls = 0;
	const vault = {
		async createBinary(path) {
			createCalls += 1;
			assert.equal(path, outputPath);
			return null;
		},
	};

	const result = await mod.writeVaultBinaryArtifact(vault, {
		path: outputPath,
		existingFile: null,
		replace: false,
	}, new ArrayBuffer(0));

	assert.equal(createCalls, 1);
	assert.deepEqual(result, { path: outputPath });
});
