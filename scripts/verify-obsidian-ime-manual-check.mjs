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

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const docxEditorAliases = await createDocxEditorAliases(resolveDocxEditorPackagesRoot(projectRoot));
const outputDir = path.join(projectRoot, 'results', 'ime-live-verify');
const demoDocxPath = path.join(projectRoot, 'tests', 'fixtures', 'lorem-ipsum.docx');

const chromeBinary =
	process.env.CHROME_PATH
	|| '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

const scenarios = [
	{
		name: 'baseline',
		label: '100% zoom, outline closed',
		expectZoom: false,
		expectMarginLeft: false,
	},
	{
		name: 'zoom125Outline',
		label: '125% zoom, outline open',
		expectZoom: true,
		expectMarginLeft: false,
	},
];

async function bundleHarness() {
	const outfile = path.join(outputDir, 'harness.js');
	
	const harnessEntry = path.join(projectRoot, 'scripts/harness/docx-ime-live-verify-entry.tsx');
	if (!existsSync(harnessEntry)) {
		console.log(`Skipping: harness entry not present on release branches: ${harnessEntry}`);
		process.exit(0);
	}

	await build({
		alias: docxEditorAliases,
		entryPoints: [path.join(projectRoot, 'scripts/harness/docx-ime-live-verify-entry.tsx')],
		bundle: true,
		format: 'iife',
		logLevel: 'silent',
		outfile,
		platform: 'browser',
		target: 'es2020',
		jsx: 'automatic',
		loader: { '.css': 'text' },
	});
	return outfile;
}

function createHarnessHtml(scenarioName, docxBase64) {
	return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>DOCX live verify — ${scenarioName}</title>
</head>
<body>
  <div id="root"></div>
  <script>
    window.__DOCX_BASE64__ = ${JSON.stringify(docxBase64)};
    window.__HARNESS_SCENARIO__ = ${JSON.stringify(scenarioName)};
  </script>
  <script src="./harness.js"></script>
</body>
</html>`;
}

function runChrome(htmlPath) {
	return new Promise((resolve, reject) => {
		execFileCb(
			chromeBinary,
			[
				'--headless=new',
				'--disable-gpu',
				'--no-sandbox',
				'--window-size=1440,1000',
				'--enable-logging=stderr',
				// Real Chromium + React layout needs wall-clock settling beyond the
				// earliest LIVE_VERIFY_RESULT samples when the machine is busy.
				'--virtual-time-budget=35000',
				htmlPath,
			],
			{ maxBuffer: 20 * 1024 * 1024, timeout: 90000 },
			(error, stdout, stderr) => {
				const output = `${stdout}\n${stderr}`;
				if (error && !output.includes('LIVE_VERIFY_RESULT:')) {
					reject(new Error(stderr || error.message));
					return;
				}
				resolve(output);
			},
		);
	});
}

function formatSampleSummary(samples) {
	return samples
		.map((entry, index) => {
			const delta = entry.hiddenCaretDelta
				? `x=${entry.hiddenCaretDelta.x} bottom=${entry.hiddenCaretDelta.bottom}`
				: 'delta=null';
			return `#${index} passed=${entry.passed} ${delta}`;
		})
		.join('; ');
}

function pickSettledMetrics(samples) {
	const passing = samples.filter((entry) => entry.passed === true);
	if (passing.length > 0) {
		return passing.at(-1);
	}
	// Prefer the newest sample with a measured caret over an earlier incomplete one.
	return (
		samples.findLast((entry) => entry.hiddenCaretDelta != null)
		?? samples.at(-1)
	);
}

function assertWrapperNeutralized(metrics, scenario) {
	assert.ok(!metrics.error, `${scenario.name}: ${metrics.error ?? 'unknown error'}`);
	assert.ok(metrics.editableFound, `${scenario.name}: body editable not found`);
	assert.ok(metrics.wrapper, `${scenario.name}: zoom wrapper not found`);

	const { wrapper } = metrics;
	assert.equal(wrapper.inlineTransform, 'none', `${scenario.name}: inline transform should be none`);
	assert.ok(
		!wrapper.inlineTransform.includes('scale(') && !wrapper.inlineTransform.includes('translateX('),
		`${scenario.name}: transform must not use scale/translateX`,
	);
	assert.equal(metrics.transformAncestorsOnCaret, 0, `${scenario.name}: caret must have zero transform ancestors`);
	assert.equal(metrics.hiddenImeRootFound, true, `${scenario.name}: hidden ProseMirror root should exist`);
	assert.equal(metrics.hiddenImeAnchored, true, `${scenario.name}: hidden IME root should be anchored`);
	assert.equal(metrics.compositionStartAnchored, true, `${scenario.name}: composition start should see an anchored root`);
	assert.ok(metrics.hiddenCaretDelta, `${scenario.name}: hidden/visible caret delta should be measured`);
	assert.ok(Math.abs(metrics.hiddenCaretDelta.x) <= 2, `${scenario.name}: hidden caret x should match visible caret`);
	assert.ok(Math.abs(metrics.hiddenCaretDelta.bottom) <= 2, `${scenario.name}: hidden caret bottom should match visible caret`);

	if (scenario.expectZoom || scenario.expectMarginLeft) {
		assert.ok(
			wrapper.inlineZoom || wrapper.inlineMarginLeft,
			`${scenario.name}: expected zoom and/or margin-left compensation`,
		);
	}
}

async function main() {
	if (!existsSync(chromeBinary)) {
		throw new Error(`Chrome not found at ${chromeBinary}`);
	}

	await mkdir(outputDir, { recursive: true });
	const docxBase64 = (await readFile(demoDocxPath)).toString('base64');
	await bundleHarness();

	const results = [];
	const maxAttempts = 5;
	for (const scenario of scenarios) {
		const htmlPath = path.join(outputDir, `${scenario.name}.html`);
		await writeFile(htmlPath, createHarnessHtml(scenario.name, docxBase64));

		let metrics = null;
		let lastError = null;
		for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
			const output = await runChrome(pathToFileURL(htmlPath).href);
			const match = output.match(/LIVE_VERIFY_RESULT:(\{.*\})/g);
			assert.ok(match?.length, `missing LIVE_VERIFY_RESULT for ${scenario.name}\n${output.slice(-800)}`);

			const samples = match.map((line) => JSON.parse(line.replace('LIVE_VERIFY_RESULT:', '')));
			// Chrome can flush a transient sample after the wrapper's inline transform
			// is removed but before its computed transform/caret geometry settles. Use
			// the newest fully passing sample from the verification window instead of
			// treating console delivery order as a lifecycle guarantee.
			metrics = pickSettledMetrics(samples);
			assert.ok(metrics, `missing parsed metrics for ${scenario.name}`);
			try {
				assertWrapperNeutralized(metrics, scenario);
				lastError = null;
				break;
			} catch (error) {
				lastError = error;
				console.warn(
					`[${scenario.label}] attempt ${attempt}/${maxAttempts} unsettled `
						+ `(${formatSampleSummary(samples)}); retrying`,
				);
				metrics = null;
			}
		}
		if (lastError) {
			throw lastError;
		}

		results.push({ scenario, metrics });

		console.log(`[${scenario.label}] PASS`);
		console.log(
			`  wrapper: transform=${metrics.wrapper.inlineTransform}`
				+ ` zoom=${metrics.wrapper.inlineZoom || '(none)'}`
				+ ` margin-left=${metrics.wrapper.inlineMarginLeft || '(none)'}`,
		);
		console.log(`  caret transform ancestors: ${metrics.transformAncestorsOnCaret}`);
		console.log(
			`  hidden IME anchor: x delta=${metrics.hiddenCaretDelta.x}`
				+ ` bottom delta=${metrics.hiddenCaretDelta.bottom}`,
		);
	}

	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
	console.log(`\nLive editor verification passed (${results.length} scenarios).`);
	console.log(`Artifacts: ${outputDir}`);
}

await main();
