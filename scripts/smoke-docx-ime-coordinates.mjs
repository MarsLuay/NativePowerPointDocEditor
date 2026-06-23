import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, '..');
const outputDir = path.join(projectRoot, 'results', 'ime-coordinates');

const chromeBinary =
	process.env.CHROME_PATH
	|| '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

// Matches @eigenpal/docx-editor-react sidebar shift when outline is open.
const SIDEBAR_DOCUMENT_SHIFT = 158;

const scenarios = [
	{
		name: 'baseline',
		workspaceTransform: 'none',
		editorTransform: 'none',
	},
	{
		name: 'zoom125',
		workspaceTransform: 'none',
		editorTransform: 'scale(1.25)',
	},
	{
		name: 'outline',
		workspaceTransform: 'none',
		editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px)`,
	},
	{
		name: 'zoom125Outline',
		workspaceTransform: 'none',
		editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px) scale(1.25)`,
	},
	{
		name: 'obsidianOffset',
		workspaceTransform: 'translate(180px, 96px)',
		editorTransform: 'none',
	},
	{
		name: 'fullStack',
		workspaceTransform: 'translate(180px, 96px)',
		editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px) scale(1.25)`,
	},
];

function createStructuralHarnessHtml(scenario) {
	return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Docx IME structural harness</title>
  <style>
    html, body { margin: 0; width: 100%; height: 100%; background: #e2e8f0; }
    .workspace-shell { width: 100%; height: 100%; padding: 24px; box-sizing: border-box; transform: ${scenario.workspaceTransform}; transform-origin: top left; }
    .docxidian-host { width: 100%; height: 100%; background: #f8fafc; position: relative; }
    .docxidian-fixed-probe { left: 0; position: fixed; top: 0; visibility: hidden; pointer-events: none; }
    .ep-root.paged-editor { width: 100%; height: 100%; }
    .editor-transform { transform: ${scenario.editorTransform}; transform-origin: top center; width: 816px; margin: 0 auto; background: white; box-shadow: 0 0 0 1px #d1d5db; }
    .layout-page-content { padding: 96px 72px; min-height: 400px; }
    .docx-run-editable { display: inline; min-height: 1em; outline: none; }
  </style>
</head>
<body>
  <div class="workspace-shell">
    <div class="docxidian-host">
      <div class="ep-root paged-editor">
        <div class="editor-transform">
          <div class="layout-page-content">
            <p><span class="docx-run-editable" contenteditable="true">Japanese IME caret probe line.</span></p>
          </div>
        </div>
      </div>
    </div>
  </div>
  <script>
    (function () {
      function collectTransformAncestors(node) {
        const ancestors = [];
        let current = node && node.parentElement;
        while (current && current !== document.documentElement) {
          const transform = getComputedStyle(current).transform;
          if (transform && transform !== 'none') {
            ancestors.push({
              tag: current.tagName.toLowerCase(),
              className: current.className || '',
              transform,
            });
          }
          current = current.parentElement;
        }
        return ancestors;
      }

      const hostEl = document.querySelector('.docxidian-host');
      const fixedProbe = document.createElement('div');
      fixedProbe.className = 'docxidian-fixed-probe';
      hostEl.appendChild(fixedProbe);
      const fixedRect = fixedProbe.getBoundingClientRect();
      fixedProbe.remove();

      const editable = document.querySelector('.docx-run-editable');
      editable.focus();
      const textNode = editable.firstChild;
      const range = document.createRange();
      range.setStart(textNode, textNode.textContent.length);
      range.collapse(true);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      const caretRect = range.getBoundingClientRect();
      const transformAncestors = collectTransformAncestors(editable);
      const hasRiskyTransform = transformAncestors.length > 0;
      const hasFixedProbeOffset = Math.abs(fixedRect.left) > 1 || Math.abs(fixedRect.top) > 1;

      document.body.dataset.metrics = encodeURIComponent(JSON.stringify({
        name: ${JSON.stringify(scenario.name)},
        fixedProbeLeft: Math.round(fixedRect.left),
        fixedProbeTop: Math.round(fixedRect.top),
        zoomContainerTransform: ${JSON.stringify(scenario.editorTransform === 'none' ? null : scenario.editorTransform)},
        transformAncestors,
        editableCount: 1,
        caretClientRect: {
          left: caretRect.left,
          top: caretRect.top,
          right: caretRect.right,
          bottom: caretRect.bottom,
        },
        caretScreenPoint: {
          x: caretRect.left + window.screenX,
          y: caretRect.bottom + window.screenY,
        },
        windowScreenOrigin: { x: window.screenX, y: window.screenY },
        imeRisk: hasRiskyTransform || hasFixedProbeOffset ? 'high' : 'low',
      }));
    })();
  </script>
</body>
</html>`;
}

function runChrome(url, userDataDir, dumpPath, timeoutMs = 10000) {
	const args = [
		chromeBinary,
		'--headless=new',
		'--disable-gpu',
		'--no-first-run',
		`--user-data-dir=${userDataDir}`,
		'--window-size=1440,1000',
		'--window-position=120,80',
		'--dump-dom',
		url,
	];
	const command = `${args.map((part) => JSON.stringify(part)).join(' ')} > ${JSON.stringify(dumpPath)} 2>/dev/null`;

	return new Promise((resolve, reject) => {
		execFile('/bin/sh', ['-c', command], { timeout: timeoutMs }, (error) => {
			if (error) {
				reject(error);
				return;
			}
			resolve();
		});
	});
}

async function runScenario(scenario) {
	const htmlPath = path.join(outputDir, `${scenario.name}.html`);
	const dumpPath = path.join(outputDir, `${scenario.name}-dump.html`);
	await writeFile(htmlPath, createStructuralHarnessHtml(scenario));
	const url = pathToFileURL(htmlPath).href;
	const userDataDir = path.join(os.tmpdir(), `docx-ime-coords-${scenario.name}-${process.pid}`);

	await runChrome(url, userDataDir, dumpPath);
	const dump = await readFile(dumpPath, 'utf8');
	const match = dump.match(/data-metrics="([^"]+)"/);
	if (!match) {
		throw new Error(`Scenario ${scenario.name} did not publish data-metrics`);
	}

	const metrics = JSON.parse(decodeURIComponent(match[1]));
	await writeFile(path.join(outputDir, `${scenario.name}-metrics.json`), `${JSON.stringify(metrics, null, 2)}\n`);
	return metrics;
}

async function main() {
	if (!existsSync(chromeBinary)) {
		throw new Error(`Chrome not found at ${chromeBinary}`);
	}

	await mkdir(outputDir, { recursive: true });

	const results = [];
	for (const scenario of scenarios) {
		const metrics = await runScenario(scenario);
		results.push(metrics);
		console.log(
			`[${scenario.name}] imeRisk=${metrics.imeRisk}`
				+ ` transforms=${metrics.transformAncestors.length}`
				+ ` fixedProbe=(${metrics.fixedProbeLeft}, ${metrics.fixedProbeTop})`
				+ ` zoomTransform=${metrics.zoomContainerTransform ?? 'none'}`,
		);
	}

	const summaryPath = path.join(outputDir, 'summary.json');
	await writeFile(summaryPath, `${JSON.stringify(results, null, 2)}\n`);

	const baseline = results.find((result) => result.name === 'baseline');
	const highRisk = results.filter((result) => result.imeRisk === 'high');

	assert.ok(baseline, 'baseline scenario missing');
	assert.equal(baseline.imeRisk, 'low', 'baseline should not have transform-induced IME risk');
	assert.ok(highRisk.length >= 3, `expected multiple high-risk transform scenarios, got ${highRisk.length}`);

	console.log(`\nDocx IME coordinate smoke passed (${results.length} scenarios, ${highRisk.length} high-risk).`);
	console.log(`Artifacts: ${summaryPath}`);
}

main().catch((error) => {
	console.error(error);
	process.exit(1);
});
