#!/usr/bin/env node
/**
 * Headless PPTX/POTX → PDF via NPDE PresentationEngine + SVG raster (no PowerPoint).
 *
 * Usage:
 *   node scripts/export-pptx-to-pdf.mjs INPUT.potx|pptx|ppsx OUTPUT.pdf
 */
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { patchPptxRendererSource } from './lib/patch-pptx-renderer.mjs';
import {
	canUseElectronMainHarness,
	projectRoot,
	resolveElectronBinary,
} from './lib/text-offset-harness.mjs';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const electronMainPath = path.join(__dirname, 'lib/electron-export-pdf-main.cjs');
const cacheDir = path.join(projectRoot, 'scripts/.cache');

const inlinePptxSvgWasmPlugin = {
	name: 'inline-pptx-svg-wasm',
	setup(buildContext) {
		buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async ({ path: modulePath }) => {
			const contents = patchPptxRendererSource(await readFile(modulePath, 'utf8'));
			return { contents, loader: 'js' };
		});
	},
};

const obsidianShim = `
export const Platform = { isDesktop: true, isMacOS: false, isMobile: false, isMobileApp: false };
export const normalizePath = (value) => String(value).replace(/\\\\/g, '/').replace(/\\/{2,}/g, '/');
export class TFile {}
export class Notice { constructor() {} }
export class Component {
  register() {}
  load() {}
  unload() {}
}
`;

async function bundleBrowserIife(entryRelative, outfileName, globalName) {
	await mkdir(cacheDir, { recursive: true });
	const outfile = path.join(cacheDir, outfileName);
	await build({
		entryPoints: [path.join(projectRoot, entryRelative)],
		bundle: true,
		format: 'iife',
		globalName,
		platform: 'browser',
		target: 'es2020',
		outfile,
		logLevel: 'silent',
		loader: { '.wasm': 'binary' },
		plugins: [
			inlinePptxSvgWasmPlugin,
			{
				name: 'stub-obsidian',
				setup(buildContext) {
					buildContext.onResolve({ filter: /^obsidian$/ }, () => ({
						path: 'obsidian',
						namespace: 'stub-obsidian',
					}));
					buildContext.onLoad({ filter: /.*/, namespace: 'stub-obsidian' }, () => ({
						contents: obsidianShim,
						loader: 'js',
					}));
				},
			},
		],
		define: { 'process.env.NODE_ENV': '"production"' },
	});
	return outfile;
}

function usage() {
	console.error('Usage: node scripts/export-pptx-to-pdf.mjs INPUT.potx|pptx|ppsx OUTPUT.pdf');
	process.exit(2);
}

function runExportElectron(electronBinary, htmlFile, inputPath, outputPath, timeoutMs) {
	return new Promise((resolve, reject) => {
		const child = spawn(electronBinary, [electronMainPath], {
			stdio: ['ignore', 'pipe', 'pipe'],
			env: {
				...process.env,
				HARNESS_HTML: htmlFile,
				EXPORT_INPUT_PATH: inputPath,
				EXPORT_OUTPUT_PATH: outputPath,
				ELECTRON_DISABLE_SECURITY_WARNINGS: 'true',
			},
		});
		let stdout = '';
		let stderr = '';
		const timer = setTimeout(() => {
			child.kill('SIGKILL');
			reject(new Error(`Electron timed out after ${timeoutMs} ms: ${stderr || stdout}`));
		}, timeoutMs);

		child.stdout.on('data', (chunk) => {
			stdout += chunk;
		});
		child.stderr.on('data', (chunk) => {
			stderr += chunk;
		});
		child.on('error', (error) => {
			clearTimeout(timer);
			reject(error);
		});
		child.on('close', () => {
			clearTimeout(timer);
			const error = stdout.match(/^HARNESS_ERROR:(.*)$/m);
			if (error) {
				reject(new Error(`Electron harness error: ${error[1]}`));
				return;
			}
			const metrics = stdout.match(/^HARNESS_METRICS:(.*)$/m);
			if (!metrics) {
				reject(new Error(`Electron emitted no metrics. stderr: ${stderr.slice(-400)}`));
				return;
			}
			resolve(metrics[1]);
		});
	});
}

async function main() {
	const inputArg = process.argv[2];
	const outputArg = process.argv[3];
	if (!inputArg || !outputArg) usage();
	const inputPath = path.resolve(inputArg);
	const outputPath = path.resolve(outputArg);
	if (!existsSync(inputPath)) {
		console.error(`Input not found: ${inputPath}`);
		process.exit(1);
	}
	if (!outputPath.toLowerCase().endsWith('.pdf')) {
		console.error('Output path must end with .pdf');
		process.exit(1);
	}

	const electronBinary = resolveElectronBinary();
	if (!electronBinary || !existsSync(electronBinary)) {
		console.error(
			'Electron binary missing. Install NPDE deps (`npm install`) or use Obsidian ai.exportPdf.',
		);
		process.exit(1);
	}
	const electronOk = await canUseElectronMainHarness(electronBinary, 8_000).catch(() => false);
	if (!electronOk) {
		console.warn('Electron -e probe failed/timed out; continuing with binary path anyway.');
	}

	const [engineBundle, exportBundle] = await Promise.all([
		bundleBrowserIife('src/PresentationEngine.ts', 'presentation-engine.iife.js', 'PresentationEngineNS'),
		bundleBrowserIife('src/ai/pptxExportPdf.ts', 'pptx-export-pdf.iife.js', 'NpdeExportPdf'),
	]);

	const htmlPath = path.join(cacheDir, 'export-pptx-to-pdf.html');
	const html = `<!doctype html>
<html><head><meta charset="utf-8" /></head>
<body>
<script src="${pathToFileURL(engineBundle).href}"></script>
<script src="${pathToFileURL(exportBundle).href}"></script>
<script>
window.__npdeExportPdf = async function(inputBase64) {
  try {
    const binary = Uint8Array.from(atob(inputBase64), (c) => c.charCodeAt(0));
    const buffer = binary.buffer.slice(binary.byteOffset, binary.byteOffset + binary.byteLength);
    const Engine = PresentationEngineNS.PresentationEngine || PresentationEngineNS.default || PresentationEngineNS;
    const engine = await Engine.load(buffer);
    const { bytes, slideCount } = await NpdeExportPdf.exportPresentationToPdfBytes(engine, {}, document);
    const u8 = new Uint8Array(bytes);
    let binaryOut = '';
    const chunk = 0x8000;
    for (let i = 0; i < u8.length; i += chunk) {
      binaryOut += String.fromCharCode.apply(null, u8.subarray(i, i + chunk));
    }
    return JSON.stringify({
      ok: true,
      slideCount,
      bytes: u8.byteLength,
      pdfBase64: btoa(binaryOut),
      engine: 'npde',
    });
  } catch (error) {
    return JSON.stringify({
      ok: false,
      error: String((error && error.stack) || error),
    });
  }
};
</script>
</body></html>`;
	await writeFile(htmlPath, html, 'utf8');

	const inputStat = await readFile(inputPath);
	const timeoutMs = Math.max(300_000, Math.round(inputStat.byteLength / 10));
	const metricsRaw = await runExportElectron(
		electronBinary,
		htmlPath,
		inputPath,
		outputPath,
		timeoutMs,
	);
	const metrics = JSON.parse(metricsRaw);
	console.log(JSON.stringify(metrics));
}

main().catch((error) => {
	console.error(error);
	process.exit(1);
});
