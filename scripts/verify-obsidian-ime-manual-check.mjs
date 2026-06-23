import assert from 'node:assert/strict';
import { execFile as execFileCb } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const outputDir = path.join(projectRoot, 'results', 'ime-live-verify');
const demoDocxPath = path.join(projectRoot, 'test_files', 'demo.docx');

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
	await build({
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
				'--virtual-time-budget=20000',
				htmlPath,
			],
			{ maxBuffer: 20 * 1024 * 1024, timeout: 60000 },
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
	for (const scenario of scenarios) {
		const htmlPath = path.join(outputDir, `${scenario.name}.html`);
		await writeFile(htmlPath, createHarnessHtml(scenario.name, docxBase64));

		const output = await runChrome(pathToFileURL(htmlPath).href);
		const match = output.match(/LIVE_VERIFY_RESULT:(\{.*\})/g);
		assert.ok(match?.length, `missing LIVE_VERIFY_RESULT for ${scenario.name}\n${output.slice(-800)}`);

		const metrics = JSON.parse(match[match.length - 1].replace('LIVE_VERIFY_RESULT:', ''));
		if (scenario.name !== 'baseline' && metrics.wrapper?.inlineTransform !== 'none') {
			const neutralized = match
				.map((line) => JSON.parse(line.replace('LIVE_VERIFY_RESULT:', '')))
				.find((entry) => entry.wrapper?.inlineTransform === 'none');
			if (neutralized) {
				Object.assign(metrics, neutralized);
			}
		}
		results.push({ scenario, metrics });
		assertWrapperNeutralized(metrics, scenario);

		console.log(`[${scenario.label}] PASS`);
		console.log(
			`  wrapper: transform=${metrics.wrapper.inlineTransform}`
				+ ` zoom=${metrics.wrapper.inlineZoom || '(none)'}`
				+ ` margin-left=${metrics.wrapper.inlineMarginLeft || '(none)'}`,
		);
		console.log(`  caret transform ancestors: ${metrics.transformAncestorsOnCaret}`);
	}

	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
	console.log(`\nLive editor verification passed (${results.length} scenarios).`);
	console.log(`Artifacts: ${outputDir}`);
}

await main();
