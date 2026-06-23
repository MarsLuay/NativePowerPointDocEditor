import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const outputDir = path.join(projectRoot, 'results', 'ime-neutralizer');

const chromeBinary =
	process.env.CHROME_PATH
	|| '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

const scenarios = [
	{
		name: 'zoom125',
		transform: 'scale(1.25)',
	},
	{
		name: 'outline',
		transform: 'translateX(-176px)',
	},
	{
		name: 'zoom125Outline',
		transform: 'translateX(-176px) scale(1.25)',
	},
];

async function bundleNeutralizerForBrowser() {
	const outfile = path.join(outputDir, 'neutralizer.js');
	await build({
		entryPoints: [path.join(projectRoot, 'src/docxImeTransformNeutralizer.ts')],
		bundle: true,
		format: 'iife',
		globalName: 'DocxImeNeutralizer',
		logLevel: 'silent',
		outfile,
		platform: 'browser',
		target: 'es2020',
	});
	return outfile;
}

function createHarnessHtml(scenario) {
	return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>DOCX IME neutralizer harness</title>
  <style>
    html, body { margin: 0; width: 100%; height: 100%; background: #e2e8f0; }
    .docxidian-host { width: 100%; height: 100%; background: #f8fafc; }
    .ep-root.paged-editor { width: 100%; height: 100%; }
    .editor-transform { width: 816px; margin: 0 auto; background: white; box-shadow: 0 0 0 1px #d1d5db; }
    .layout-page-content { padding: 96px 72px; min-height: 400px; }
    .docx-run-editable { display: inline; min-height: 1em; outline: none; }
  </style>
  <script src="./neutralizer.js"></script>
</head>
<body>
  <div class="docxidian-host">
    <div class="ep-root paged-editor docxidian-editor-harness">
      <div class="editor-transform">
        <div class="paged-editor__pages">
          <div class="layout-page-content">
            <p><span class="docx-run-editable" contenteditable="true">Japanese IME caret probe line.</span></p>
          </div>
        </div>
      </div>
    </div>
  </div>
  <script>
    const {
      attachDocxImeTransformNeutralizer,
      countTransformAncestors,
      findDocxEditorZoomWrapper,
      neutralizeDocxEditorZoomWrapper,
    } = DocxImeNeutralizer;

    const editorRoot = document.querySelector('.docxidian-editor-harness');
    const editable = document.querySelector('.docx-run-editable');
    const wrapper = findDocxEditorZoomWrapper(editorRoot);
    wrapper.style.transform = ${JSON.stringify(scenario.transform)};
    wrapper.style.transformOrigin = 'top center';

    const before = {
      wrapperTransform: wrapper.style.transform,
      transformAncestors: countTransformAncestors(editable),
    };

    const detach = attachDocxImeTransformNeutralizer(editorRoot);
    neutralizeDocxEditorZoomWrapper(wrapper);

    const after = {
      wrapperTransform: wrapper.style.transform,
      wrapperZoom: wrapper.style.zoom || '',
      wrapperMarginLeft: wrapper.style.marginLeft || '',
      transformAncestors: countTransformAncestors(editable),
    };

    detach();
    console.log('NEUTRALIZER_RESULT:' + JSON.stringify({
      scenario: ${JSON.stringify(scenario.name)},
      before,
      after,
    }));
  </script>
</body>
</html>`;
}

function runChromeScreenshot(htmlPath, screenshotPath) {
	return new Promise((resolve, reject) => {
		execFile(
			chromeBinary,
			[
				'--headless=new',
				'--disable-gpu',
				'--no-sandbox',
				'--window-size=1024,768',
				'--screenshot=' + screenshotPath,
				'--virtual-time-budget=5000',
				htmlPath,
			],
			(error, _stdout, stderr) => {
				if (error) {
					reject(new Error(stderr || error.message));
					return;
				}
				resolve();
			},
		);
	});
}

function runChrome(htmlPath) {
	return new Promise((resolve, reject) => {
		execFile(
			chromeBinary,
			[
				'--headless=new',
				'--disable-gpu',
				'--no-sandbox',
				'--enable-logging=stderr',
				'--virtual-time-budget=5000',
				htmlPath,
			],
			{ maxBuffer: 10 * 1024 * 1024 },
			(error, stdout, stderr) => {
				const output = `${stdout}\n${stderr}`;
				if (error && !output.includes('NEUTRALIZER_RESULT:')) {
					reject(new Error(stderr || error.message));
					return;
				}
				resolve(output);
			},
		);
	});
}

async function main() {
	await mkdir(outputDir, { recursive: true });
	await bundleNeutralizerForBrowser();
	const results = [];

	for (const scenario of scenarios) {
		const htmlPath = path.join(outputDir, `${scenario.name}.html`);
		await writeFile(htmlPath, createHarnessHtml(scenario));

		const output = await runChrome(htmlPath);
		const match = output.match(/NEUTRALIZER_RESULT:(\{.*?\})(?:"|, source:)/);
		assert.ok(match, `missing neutralizer result for ${scenario.name}\n${output.slice(-500)}`);
		const result = JSON.parse(match[1]);
		results.push(result);

		assert.equal(result.before.transformAncestors, 1, `${scenario.name}: expected one transform ancestor before`);
		assert.equal(result.after.wrapperTransform, 'none', `${scenario.name}: wrapper transform should be none`);
		assert.equal(result.after.transformAncestors, 0, `${scenario.name}: caret should have no transform ancestors`);

		await runChromeScreenshot(htmlPath, path.join(outputDir, `${scenario.name}.png`));
	}

	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
	console.log('DOCX IME neutralizer smoke passed:', results.map((result) => result.scenario).join(', '));
}

await main();
