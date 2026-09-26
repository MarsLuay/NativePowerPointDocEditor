import { execFile } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.join(__dirname, '..', 'results', 'ime-coordinates');
const chromeBinary = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const SIDEBAR_DOCUMENT_SHIFT = 158;

const scenarios = [
	{ name: 'baseline', workspaceTransform: 'none', editorTransform: 'none' },
	{ name: 'zoom125', workspaceTransform: 'none', editorTransform: 'scale(1.25)' },
	{ name: 'outline', workspaceTransform: 'none', editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px)` },
	{ name: 'zoom125Outline', workspaceTransform: 'none', editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px) scale(1.25)` },
	{ name: 'obsidianOffset', workspaceTransform: 'translate(180px, 96px)', editorTransform: 'none' },
	{ name: 'fullStack', workspaceTransform: 'translate(180px, 96px)', editorTransform: `translateX(-${SIDEBAR_DOCUMENT_SHIFT}px) scale(1.25)` },
];

function createHtml(scenario) {
	return `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{margin:0;height:100%;background:#e2e8f0}
.workspace-shell{width:100%;height:100%;padding:24px;box-sizing:border-box;transform:${scenario.workspaceTransform};transform-origin:top left}
.native-powerpoint-doc-editor-host{width:100%;height:100%;background:#f8fafc;position:relative}
.native-powerpoint-doc-editor-fixed-probe{left:0;position:fixed;top:0;visibility:hidden}
.editor-transform{transform:${scenario.editorTransform};transform-origin:top center;width:816px;margin:0 auto;background:#fff}
[data-native-powerpoint-doc-editor-page-content]{padding:96px 72px;min-height:400px}
</style></head><body><div class="workspace-shell"><div class="native-powerpoint-doc-editor-host"><div class="editor-transform"><div data-native-powerpoint-doc-editor-page-content="true"><span class="docx-run-editable" contenteditable="true">Japanese IME caret probe line.</span></div></div></div></div>
<script>(()=>{const host=document.querySelector('.native-powerpoint-doc-editor-host');const probe=document.createElement('div');probe.className='native-powerpoint-doc-editor-fixed-probe';host.appendChild(probe);const fixedRect=probe.getBoundingClientRect();probe.remove();const editable=document.querySelector('.docx-run-editable');editable.focus();const range=document.createRange();range.selectNodeContents(editable);range.collapse(false);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(range);const ancestors=[];for(let n=editable.parentElement;n&&n!==document.documentElement;n=n.parentElement){const t=getComputedStyle(n).transform;if(t&&t!=='none')ancestors.push({tag:n.tagName.toLowerCase(),className:n.className||'',transform:t});}
const risky=ancestors.length>0||Math.abs(fixedRect.left)>1||Math.abs(fixedRect.top)>1;document.body.dataset.metrics=encodeURIComponent(JSON.stringify({name:'${scenario.name}',fixedProbeLeft:Math.round(fixedRect.left),fixedProbeTop:Math.round(fixedRect.top),zoomContainerTransform:${JSON.stringify(scenario.editorTransform === 'none' ? null : scenario.editorTransform)},transformAncestors:ancestors,imeRisk:risky?'high':'low'}));})();</script></body></html>`;
}

function runChrome(url, dumpPath) {
	const udir = path.join(os.tmpdir(), `docx-ime-${path.basename(dumpPath)}-${Date.now()}`);
	const command = `${JSON.stringify(chromeBinary)} --headless=new --disable-gpu --no-first-run --user-data-dir=${JSON.stringify(udir)} --dump-dom ${JSON.stringify(url)} > ${JSON.stringify(dumpPath)} 2>/dev/null`;
	return new Promise((resolve, reject) => {
		execFile('/bin/sh', ['-c', command], { timeout: 15000 }, (error) => error ? reject(error) : resolve());
	});
}

await mkdir(outputDir, { recursive: true });
const results = [];
for (const scenario of scenarios) {
	const htmlPath = path.join(outputDir, `${scenario.name}.html`);
	const dumpPath = path.join(outputDir, `${scenario.name}-dump.html`);
	await writeFile(htmlPath, createHtml(scenario));
	await new Promise((resolve) => setTimeout(resolve, 1500));
	await runChrome(pathToFileURL(htmlPath).href, dumpPath);
	const dump = await readFile(dumpPath, 'utf8');
	const metrics = JSON.parse(decodeURIComponent(dump.match(/data-metrics="([^"]+)"/)[1]));
	await writeFile(path.join(outputDir, `${scenario.name}-metrics.json`), `${JSON.stringify(metrics, null, 2)}\n`);
	results.push(metrics);
	console.log(`[${scenario.name}] imeRisk=${metrics.imeRisk} transforms=${metrics.transformAncestors.length}`);
}
await writeFile(path.join(outputDir, 'summary.json'), `${JSON.stringify(results, null, 2)}\n`);
