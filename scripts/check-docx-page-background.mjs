import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const stylesPath = path.join(projectRoot, 'styles.css');
const runtimeStylesPath = path.join(projectRoot, 'vendor/docx-editor-runtime/react/dist/styles.css');
const viewPath = path.join(projectRoot, 'src/DocxReactView.tsx');

const [styles, runtimeStyles, viewSource] = await Promise.all([
	readFile(stylesPath, 'utf8'),
	readFile(runtimeStylesPath, 'utf8'),
	readFile(viewPath, 'utf8'),
]);

assert.match(
	styles,
	/--doc-page-bg:\s*var\(--npde-document-bg\);/,
	'NPDE DOCX root must normalize the runtime page-background fallback',
);
const hostOverride = viewSource.match(/const hostCaretOverride = `([\s\S]*?)`\.trim\(\);/)?.[1] ?? '';
assert.match(
	hostOverride,
	/--doc-page-bg:\s*var\(--npde-document-bg\);/,
	'post-runtime DOCX styles must override the dark-only #cccccc scaffold',
);
assert.match(hostOverride, /\.layout-page[\s\S]*?filter:\s*none;/);

function findChrome() {
	const candidates = [
		process.env.CHROME_PATH,
		path.join(process.env.ProgramFiles || '', 'Google/Chrome/Application/chrome.exe'),
		path.join(process.env['ProgramFiles(x86)'] || '', 'Google/Chrome/Application/chrome.exe'),
		path.join(process.env.ProgramFiles || '', 'Microsoft/Edge/Application/msedge.exe'),
		path.join(process.env['ProgramFiles(x86)'] || '', 'Microsoft/Edge/Application/msedge.exe'),
		'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
		'/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
		'/Applications/Chromium.app/Contents/MacOS/Chromium',
		'/usr/bin/google-chrome',
		'/usr/bin/chromium',
		'/usr/bin/chromium-browser',
	].filter(Boolean);
	return candidates.find((candidate) => existsSync(candidate)) ?? null;
}

const chrome = findChrome();
if (!chrome) {
	console.log('DOCX page background contract passed (Chrome/Edge unavailable; browser assertion not run).');
	process.exit(0);
}

const tempDirectory = await mkdtemp(path.join(os.tmpdir(), 'npde-docx-page-background-'));
const htmlPath = path.join(tempDirectory, 'check.html');
const escapedStyles = `${styles}\n${runtimeStyles}`.replace(/<\/style/gi, '<\\/style');
const escapedOverride = hostOverride.replace(/<\/style/gi, '<\\/style');
const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>${escapedStyles}</style>
<style>${escapedOverride}</style>
</head>
<body class="native-powerpoint-doc-editor-theme-resolved-light">
<div class="workspace-leaf-content" data-type="native-powerpoint-doc-editor-docx-view">
  <div class="native-powerpoint-doc-editor-host">
    <div id="docx-root" class="docx-editor-root docx-editor" data-native-powerpoint-doc-editor-root="true">
      <div id="docx-page" class="layout-page" style="background-color: var(--doc-page-bg, #ffffff)"></div>
      <div id="authored-page" class="layout-page" style="background-color: rgb(12, 34, 56)"></div>
    </div>
  </div>
</div>
<div class="native-powerpoint-root">
  <div id="pptx-slide" class="native-powerpoint-slide-surface"></div>
</div>
<script>
(() => {
  const root = document.querySelector('#docx-root');
  const page = document.querySelector('#docx-page');
  const authoredPage = document.querySelector('#authored-page');
  const slide = document.querySelector('#pptx-slide');
  function measure(theme) {
    document.body.className = 'native-powerpoint-doc-editor-theme-resolved-' + theme;
    root.classList.toggle('dark', theme === 'dark');
    const pageStyle = getComputedStyle(page);
    const authoredStyle = getComputedStyle(authoredPage);
    const slideStyle = getComputedStyle(slide);
    return {
      pageBackground: pageStyle.backgroundColor,
      slideBackground: slideStyle.backgroundColor,
      authoredBackground: authoredStyle.backgroundColor,
      pageFilter: pageStyle.filter,
      pageToken: getComputedStyle(root).getPropertyValue('--doc-page-bg').trim(),
    };
  }
  const metrics = { light: measure('light'), dark: measure('dark') };
  document.body.dataset.metrics = encodeURIComponent(JSON.stringify(metrics));
})();
</script>
</body>
</html>`;

await writeFile(htmlPath, html, 'utf8');
try {
	const result = await new Promise((resolve, reject) => {
		const child = spawn(chrome, [
			'--headless=new',
			'--disable-gpu',
			'--no-first-run',
			'--no-default-browser-check',
			'--no-sandbox',
			'--dump-dom',
			`--user-data-dir=${path.join(tempDirectory, 'profile')}`,
			pathToFileURL(htmlPath).href,
		], { stdio: ['ignore', 'pipe', 'pipe'] });
		let stdout = '';
		let stderr = '';
		child.stdout.on('data', (chunk) => { stdout += chunk; });
		child.stderr.on('data', (chunk) => { stderr += chunk; });
		child.once('error', reject);
		child.once('close', (code) => code === 0
			? resolve(stdout)
			: reject(new Error(`Chrome exited with ${code}: ${stderr || stdout}`)));
	});
	const encoded = result.match(/data-metrics="([^"]+)"/)?.[1];
	assert.ok(encoded, 'browser check must publish computed page metrics');
	const metrics = JSON.parse(decodeURIComponent(encoded));
	for (const theme of ['light', 'dark']) {
		const sample = metrics[theme];
		assert.equal(sample.pageBackground, 'rgb(255, 255, 255)', `${theme} DOCX page must be NPDE white`);
		assert.equal(sample.slideBackground, 'rgb(255, 255, 255)', `${theme} PPTX slide must be NPDE white`);
		assert.equal(sample.pageBackground, sample.slideBackground, `${theme} DOCX/PPTX surfaces must match`);
		assert.equal(sample.authoredBackground, 'rgb(12, 34, 56)', `${theme} authored page background must survive`);
		assert.equal(sample.pageFilter, 'none', `${theme} DOCX page inversion must remain disabled`);
		assert.equal(sample.pageToken, '#ffffff', `${theme} DOCX page token must resolve to NPDE document white`);
	}
	console.log('DOCX page background browser check passed for resolved light and dark themes.');
} finally {
	await rm(tempDirectory, { recursive: true, force: true });
}
