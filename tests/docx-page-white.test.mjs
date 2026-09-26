import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const vendorDarkPage = `
.docx-editor-root.dark { --doc-page-bg: #cccccc; }
.docx-editor-root.dark .layout-page { filter: invert(1) hue-rotate(180deg) contrast(0.92); }
.native-powerpoint-slide-surface { background: var(--npde-document-bg); }
`;

function documentHtml(bodyClass, { darkRuntime }) {
	const rootClass = darkRuntime ? 'docx-editor-root docx-editor dark' : 'docx-editor-root docx-editor';
	return `<!doctype html>
<html>
<head><style id="styles"></style></head>
<body class="${bodyClass}">
	<div class="workspace-leaf-content" data-type="native-powerpoint-doc-editor-docx-view">
		<div class="native-powerpoint-doc-editor-host">
			<div class="${rootClass}" data-native-powerpoint-doc-editor-root>
				<div class="layout-page" data-native-powerpoint-doc-editor-page></div>
			</div>
		</div>
	</div>
	<div class="native-powerpoint-root">
		<div class="native-powerpoint-slide-surface"></div>
	</div>
</body>
</html>`;
}

function readPaint(dom) {
	const page = dom.window.document.querySelector('[data-native-powerpoint-doc-editor-page]');
	const slide = dom.window.document.querySelector('.native-powerpoint-slide-surface');
	const root = page.parentElement;
	page.style.backgroundColor = 'var(--doc-page-bg, #ffffff)';
	const pageStyle = dom.window.getComputedStyle(page);
	const rootStyle = dom.window.getComputedStyle(root);
	const slideStyle = dom.window.getComputedStyle(slide);
	return {
		pageToken: rootStyle.getPropertyValue('--doc-page-bg').trim(),
		pageWhite: pageStyle.getPropertyValue('--npde-document-bg').trim(),
		slideWhite: slideStyle.getPropertyValue('--npde-document-bg').trim(),
		pageFilter: pageStyle.filter,
	};
}

async function mount(bodyClass, darkRuntime) {
	const css = await readFile(path.join(projectRoot, 'styles.css'), 'utf8');
	const dom = new JSDOM(documentHtml(bodyClass, { darkRuntime }), { pretendToBeVisual: true });
	dom.window.document.getElementById('styles').textContent = `${vendorDarkPage}\n${css}`;
	return readPaint(dom);
}

function assertSharedWhite(painted, label) {
	assert.equal(painted.pageToken, 'var(--npde-document-bg)', label);
	assert.notEqual(painted.pageToken, '#cccccc', label);
	assert.equal(painted.pageWhite, '#ffffff', label);
	assert.equal(painted.slideWhite, painted.pageWhite, label);
}

test('blank DOCX page and PPTX slide share NPDE white in light and dark', async () => {
	const light = await mount('native-powerpoint-doc-editor-theme-resolved-light', false);
	const dark = await mount('native-powerpoint-doc-editor-theme-resolved-dark', true);
	assertSharedWhite(light, 'light');
	assertSharedWhite(dark, 'dark');
	assert.equal(dark.pageFilter, 'none');
});

test('an authored DOCX page color is not replaced by the blank-page white', async () => {
	const dom = new JSDOM(documentHtml('native-powerpoint-doc-editor-theme-resolved-dark', { darkRuntime: true }), { pretendToBeVisual: true });
	const css = await readFile(path.join(projectRoot, 'styles.css'), 'utf8');
	dom.window.document.getElementById('styles').textContent = `${vendorDarkPage}\n${css}`;
	const page = dom.window.document.querySelector('[data-native-powerpoint-doc-editor-page]');
	page.style.backgroundColor = '#ddeeff';
	assert.equal(dom.window.getComputedStyle(page).backgroundColor, 'rgb(221, 238, 255)');
});
