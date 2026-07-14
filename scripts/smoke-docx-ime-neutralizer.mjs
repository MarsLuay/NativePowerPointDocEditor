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
    .native-powerpoint-doc-editor-host { width: 100%; height: 100%; background: #f8fafc; }
    [data-native-powerpoint-doc-editor-root] { width: 100%; height: 100%; }
    .editor-transform { width: 816px; margin: 0 auto; background: white; box-shadow: 0 0 0 1px #d1d5db; }
    [data-native-powerpoint-doc-editor-page-content] { padding: 96px 72px; min-height: 400px; }
    .docx-run-editable { display: inline; min-height: 1em; outline: none; }
  </style>
  <script src="./neutralizer.js"></script>
</head>
<body>
  <div class="native-powerpoint-doc-editor-host">
    <div class="native-powerpoint-doc-editor-editor-harness" data-native-powerpoint-doc-editor-root="true">
      <div class="editor-transform">
        <div data-native-powerpoint-doc-editor-pages="true">
          <div data-native-powerpoint-doc-editor-page-content="true">
            <p><span class="docx-run-editable" contenteditable="true">Japanese IME caret probe line.</span></p>
          </div>
        </div>
      </div>
    </div>
  </div>
  <script>
    const {
      attachDocxImeTransformNeutralizer,
      syncDocxImeHiddenProseMirrorAnchor,
      countTransformAncestors,
      findDocxEditorZoomWrapper,
      neutralizeDocxEditorZoomWrapper,
    } = DocxImeNeutralizer;

    const editorRoot = document.querySelector('.native-powerpoint-doc-editor-editor-harness');
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
      wrapperTransform: wrapper.style.transform || 'none',
      wrapperZoom: wrapper.style.zoom || '',
      wrapperMarginLeft: wrapper.style.marginLeft || '',
      transformAncestors: countTransformAncestors(editable),
    };

    const visibleCaret = document.createElement('div');
    visibleCaret.setAttribute('data-native-powerpoint-doc-editor-caret', 'true');
    visibleCaret.style.position = 'fixed';
    visibleCaret.style.left = '300px';
    visibleCaret.style.top = '400px';
    visibleCaret.style.width = '2px';
    visibleCaret.style.height = '24px';
    editorRoot.appendChild(visibleCaret);

    const hiddenRoot = document.createElement('div');
    hiddenRoot.setAttribute('data-native-powerpoint-doc-editor-hidden-prosemirror', 'true');
    hiddenRoot.style.position = 'absolute';
    hiddenRoot.style.left = '-9999px';
    hiddenRoot.style.top = '0';
    hiddenRoot.style.opacity = '0';
    const hiddenPm = document.createElement('div');
    hiddenPm.className = 'ProseMirror';
    hiddenPm.contentEditable = 'true';
    hiddenRoot.appendChild(hiddenPm);
    document.body.appendChild(hiddenRoot);

    const fakeView = {
      dom: hiddenPm,
      state: { selection: { head: 5 } },
      hasFocus: () => true,
      coordsAtPos: () => ({ left: -9400, right: -9398, top: 200, bottom: 220 }),
    };
    const diagnostics = [];
    const anchorSynced = syncDocxImeHiddenProseMirrorAnchor(editorRoot, {
      getEditorView: () => fakeView,
      onDiagnostic: (event) => diagnostics.push(event),
    });
    const anchor = {
      synced: anchorSynced,
      hiddenLeft: hiddenRoot.style.left,
      hiddenTop: hiddenRoot.style.top,
      hiddenAnchored: hiddenRoot.dataset.nativePowerPointDocEditorImeAnchored === 'true',
    };
    const detachDiagnosticNeutralizer = attachDocxImeTransformNeutralizer(editorRoot, {
      getEditorView: () => fakeView,
      onDiagnostic: (event) => diagnostics.push(event),
    });
    document.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: 'に' }));
    document.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '日本' }));
    detachDiagnosticNeutralizer();
    const restoredAnchor = {
      position: hiddenRoot.style.position,
      left: hiddenRoot.style.left,
      top: hiddenRoot.style.top,
      anchored: hiddenRoot.dataset.nativePowerPointDocEditorImeAnchored === 'true',
    };

    detach();
    console.log('NEUTRALIZER_RESULT:' + JSON.stringify({
      scenario: ${JSON.stringify(scenario.name)},
      before,
      after,
      anchor,
      restoredAnchor,
      diagnostics,
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
		assert.equal(result.after.wrapperTransform || 'none', 'none', `${scenario.name}: wrapper transform should be none`);
		assert.equal(result.after.transformAncestors, 0, `${scenario.name}: caret should have no transform ancestors`);
		assert.equal(result.anchor.synced, true, `${scenario.name}: hidden IME anchor should sync`);
		assert.equal(result.anchor.hiddenLeft, '-299px', `${scenario.name}: hidden IME anchor should align x`);
		assert.equal(result.anchor.hiddenTop, '204px', `${scenario.name}: hidden IME anchor should align bottom y`);
		assert.equal(result.anchor.hiddenAnchored, true, `${scenario.name}: hidden IME anchor should be marked`);
		assert.deepEqual(result.restoredAnchor, {
			position: 'absolute',
			left: '-9999px',
			top: '0px',
			anchored: false,
		}, `${scenario.name}: hidden editor positioning should be restored on detach`);
		assert.ok(
			result.diagnostics.some((event) => event.event === 'anchor-state' && event.details?.status === 'synced'),
			`${scenario.name}: diagnostics should record a synced IME anchor`,
		);
		assert.ok(
			result.diagnostics.some((event) => event.event === 'composition-start' && event.details?.anchored === true),
			`${scenario.name}: diagnostics should record composition start`,
		);
		assert.ok(
			result.diagnostics.some((event) => event.event === 'composition-end'),
			`${scenario.name}: diagnostics should record composition end`,
		);

		await runChromeScreenshot(htmlPath, path.join(outputDir, `${scenario.name}.png`));
	}

	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
	console.log('DOCX IME neutralizer smoke passed:', results.map((result) => result.scenario).join(', '));
}

await main();
