import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('rendered PDF export uses detached canvases and prefers the SVG renderer', async () => {
	const source = await readFile(path.join(projectRoot, 'src/renderedPdfExport.ts'), 'utf8');
	const renderPageFunction = source.slice(
		source.indexOf('async function renderPageElementToJpeg'),
		source.indexOf('function createRenderedPdfContentStream'),
	);

	assert.match(source, /import \{ createDetachedMeasureCanvas \} from '\.\/powerpoint\/measureCanvas';/);
	assert.doesNotMatch(source, /\bactiveDocument\.createEl\(/);
	assert.match(source, /createDetachedMeasureCanvas\(page\.ownerDocument\)/);
	assert.ok(
		renderPageFunction.indexOf('renderPageElementToSvgJpeg') < renderPageFunction.indexOf('renderPageElementToCanvasJpeg'),
		'SVG rendering should be attempted before html2canvas',
	);
});
