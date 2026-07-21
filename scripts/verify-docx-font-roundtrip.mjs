import assert from 'node:assert/strict';
import { execFile as execFileCb } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import {
	createDocxEditorAliases,
	resolveDocxEditorPackagesRoot,
} from './lib/docx-editor-aliases.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const docxEditorAliases = await createDocxEditorAliases(resolveDocxEditorPackagesRoot(projectRoot), projectRoot);
const outputDir = path.join(projectRoot, 'results', 'docx-font-roundtrip');
const fixturePath = path.join(projectRoot, 'tests', 'fixtures', 'docx', 'table-cell-direct-24pt-font.docx');
const chromeBinary = process.env.CHROME_PATH
	|| '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

async function bundleHarness() {
	
	const harnessEntry = path.join(projectRoot, 'scripts', 'harness', 'docx-font-roundtrip-entry.tsx');
	if (!existsSync(harnessEntry)) {
		console.error(`Harness entry not present on release branches: ${harnessEntry}`);
		process.exit(1);
	}

	await build({
		alias: docxEditorAliases,
		entryPoints: [path.join(projectRoot, 'scripts', 'harness', 'docx-font-roundtrip-entry.tsx')],
		bundle: true,
		format: 'iife',
		logLevel: 'silent',
		outfile: path.join(outputDir, 'harness.js'),
		platform: 'browser',
		target: 'es2020',
		jsx: 'automatic',
		loader: { '.css': 'text' },
	});
}

function runChrome(url) {
	return new Promise((resolve, reject) => {
		execFileCb(
			chromeBinary,
			[
				'--headless=new',
				'--disable-gpu',
				'--no-sandbox',
				'--window-size=1440,1000',
				'--enable-logging=stderr',
				'--virtual-time-budget=20000',
				url,
			],
			{ maxBuffer: 20 * 1024 * 1024, timeout: 60000 },
			(error, stdout, stderr) => {
				const output = `${stdout}\n${stderr}`;
				if (error && !output.includes('FONT_ROUNDTRIP_RESULT:')) {
					reject(new Error(stderr || error.message));
					return;
				}
				resolve(output);
			},
		);
	});
}

async function main() {
	if (!existsSync(chromeBinary)) {
		throw new Error(`Chrome not found at ${chromeBinary}`);
	}

	await mkdir(outputDir, { recursive: true });
	await bundleHarness();
	const docxBase64 = (await readFile(fixturePath)).toString('base64');
	const htmlPath = path.join(outputDir, 'index.html');
	await writeFile(htmlPath, `<!doctype html>
<html lang="en">
<head><meta charset="utf-8"><title>DOCX font round trip</title></head>
<body>
<div id="root"></div>
<script>window.__DOCX_BASE64__ = ${JSON.stringify(docxBase64)};</script>
<script src="./harness.js"></script>
</body>
</html>`);

	const output = await runChrome(pathToFileURL(htmlPath).href);
	const matches = output.match(/FONT_ROUNDTRIP_RESULT:(\{.*\})/g);
	assert.ok(matches?.length, `missing font round-trip result\n${output.slice(-1200)}`);
	const metrics = JSON.parse(matches[matches.length - 1].replace('FONT_ROUNDTRIP_RESULT:', ''));
	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(metrics, null, 2)}\n`);

	assert.equal(metrics.editApplied, true, 'the harness edit should be applied');
	assert.equal(metrics.rawCellCount, 20, 'the editor should retain all table cells');
	assert.equal(metrics.raw24PtSizeCount, 20, 'the editor should serialize all table runs at 24 pt');
	assert.equal(metrics.raw24PtComplexSizeCount, 20, 'the editor should serialize all complex-script sizes at 24 pt');
	assert.equal(metrics.rawTable11PtSizeCount, 0, 'the editor should not write the 11 pt default into table runs');
	assert.equal(metrics.rawTable11PtComplexSizeCount, 0, 'the editor should not write the complex-script default into table runs');
	assert.equal(metrics.repaired24PtSizeCount, 20, 'all table runs should retain 24 pt size after repair');
	assert.equal(metrics.repaired24PtComplexSizeCount, 20, 'all complex-script sizes should retain 24 pt after repair');
	for (const [text, size] of Object.entries(metrics.renderedCellFontSizes)) {
		assert.notEqual(size, null, `${text}: rendered text element should be found`);
		assert.ok(Number.parseFloat(size) >= 31, `${text}: expected rendered 24 pt text, got ${size}`);
	}

	console.log('DOCX font round-trip verification passed.');
	console.log(JSON.stringify(metrics, null, 2));
}

await main();
