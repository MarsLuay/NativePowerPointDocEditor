import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

test('slideSizeEmuToPdfPoints maps 24x36 poster EMUs to 1728x2592 pts', async () => {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-pdf-size-'));
	const outfile = path.join(outputDirectory, 'powerpoint-export.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/PowerPointExport.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		plugins: [
			{
				name: 'stub-pdf-deps',
				setup(buildContext) {
					buildContext.onResolve({ filter: /^(jszip|\.\/renderedPdfExport)$/ }, (args) => ({
						path: args.path,
						namespace: 'stub',
					}));
					buildContext.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({
						contents: `
export default {};
export const createRenderedImagePdf = () => new ArrayBuffer(0);
export const dataUrlToBytes = () => new Uint8Array();
`,
						loader: 'js',
					}));
				},
			},
		],
	});
	const mod = require(outfile);
	const page = mod.slideSizeEmuToPdfPoints(21945600, 32918400);
	assert.equal(page.width, 1728);
	assert.equal(page.height, 2592);
});
