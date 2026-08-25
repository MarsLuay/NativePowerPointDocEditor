import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { loadRenderedPdfExportModule } from './helpers/load-plugin-modules.mjs';

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

test('dataUrlToBytes decodes standard data URLs, raw base64, empty payloads, and binary byte ranges', async () => {
	const { dataUrlToBytes } = await loadRenderedPdfExportModule();

	// Standard data URL with image/jpeg header
	const helloWorldBase64 = Buffer.from('Hello World').toString('base64');
	const jpegDataUrl = `data:image/jpeg;base64,${helloWorldBase64}`;
	const bytes1 = dataUrlToBytes(jpegDataUrl);
	assert.deepEqual(bytes1, new Uint8Array(Buffer.from('Hello World')));

	// Data URL with custom MIME type and parameters
	const textDataUrl = `data:text/plain;charset=utf-8;base64,${helloWorldBase64}`;
	const bytes2 = dataUrlToBytes(textDataUrl);
	assert.deepEqual(bytes2, new Uint8Array(Buffer.from('Hello World')));

	// Raw base64 string without comma prefix
	const bytes3 = dataUrlToBytes(helloWorldBase64);
	assert.deepEqual(bytes3, new Uint8Array(Buffer.from('Hello World')));

	// Empty payloads
	const emptyStringBytes = dataUrlToBytes('');
	assert.deepEqual(emptyStringBytes, new Uint8Array(0));

	const emptyHeaderBytes = dataUrlToBytes('data:image/png;base64,');
	assert.deepEqual(emptyHeaderBytes, new Uint8Array(0));

	// Full 8-bit byte range round-trip (0x00 .. 0xFF)
	const allBytes = new Uint8Array(256);
	for (let i = 0; i < 256; i++) {
		allBytes[i] = i;
	}
	const binaryBase64 = Buffer.from(allBytes).toString('base64');
	const binaryDataUrl = `data:application/octet-stream;base64,${binaryBase64}`;
	const bytes4 = dataUrlToBytes(binaryDataUrl);
	assert.deepEqual(bytes4, allBytes);
});
