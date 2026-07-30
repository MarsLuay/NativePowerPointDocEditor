import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { gzipSync } from 'node:zlib';
import path from 'node:path';
import esbuild from 'esbuild';
import { patchPptxRendererSource } from './patch-pptx-renderer.mjs';
import { PPTX_RUNTIME_ARTIFACT_SPECS } from './pptx-runtime-artifact-spec.mjs';

const PAYLOADS_RELATIVE = 'src/powerpoint/generated/runtimeArtifactPayloads.json';

const inlinePptxSvgWasmPlugin = {
	name: 'inline-pptx-svg-wasm',
	setup(build) {
		build.onLoad({ filter: /pptx-renderer\.js$/ }, async (args) => {
			const normalizedPath = args.path.replace(/\\/g, '/');
			if (!normalizedPath.endsWith('/node_modules/pptx-svg/dist/pptx-renderer.js')) {
				return undefined;
			}
			const source = await readFile(args.path, 'utf8');
			return { contents: patchPptxRendererSource(source), loader: 'js' };
		});
	},
};

/**
 * Build sibling optional .mjs runtimes and rewrite the gzip payload JSON
 * that main.js embeds for community-plugin installs.
 *
 * @param {{ projectRoot: string, minify?: boolean, logLevel?: import('esbuild').LogLevel }} options
 */
export async function emitPptxRuntimeArtifacts(options) {
	const { projectRoot, minify = false, logLevel = 'info' } = options;

	await Promise.all(PPTX_RUNTIME_ARTIFACT_SPECS.map(async ({ source, artifact, bundle }) => {
		await esbuild.build({
			entryPoints: [path.join(projectRoot, source)],
			outfile: path.join(projectRoot, artifact),
			bundle,
			format: 'esm',
			target: 'es2020',
			loader: { '.wasm': 'binary' },
			minify,
			logLevel,
			plugins: bundle ? [inlinePptxSvgWasmPlugin] : [],
		});
	}));

	const payloads = [];
	for (const { artifact } of PPTX_RUNTIME_ARTIFACT_SPECS) {
		const bytes = await readFile(path.join(projectRoot, artifact));
		const sha256 = createHash('sha256').update(bytes).digest('hex');
		const gzipBase64 = gzipSync(bytes, { level: 9 }).toString('base64');
		payloads.push({ artifact, sha256, gzipBase64 });
	}

	const outPath = path.join(projectRoot, PAYLOADS_RELATIVE);
	await mkdir(path.dirname(outPath), { recursive: true });
	await writeFile(outPath, `${JSON.stringify(payloads)}\n`, 'utf8');
	return { payloadsPath: outPath, payloads };
}

export { PAYLOADS_RELATIVE };
