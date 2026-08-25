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

test('createRenderedImagePdf builds valid single-page PDF binary with correct structure and xref offsets', async () => {
	const { createRenderedImagePdf } = await loadRenderedPdfExportModule();

	const dummyImageBytes = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46]);
	const pages = [
		{
			imageBytes: dummyImageBytes,
			imageWidth: 800,
			imageHeight: 600,
			pdfWidth: 600,
			pdfHeight: 450,
			textRuns: [],
		},
	];

	const pdfBuffer = createRenderedImagePdf(pages);
	assert.ok(pdfBuffer instanceof ArrayBuffer, 'Returns an ArrayBuffer');

	const pdfText = new TextDecoder('latin1').decode(pdfBuffer);
	assert.ok(pdfText.startsWith('%PDF-1.4\n'), 'PDF starts with %PDF-1.4 header');
	assert.ok(pdfText.includes('%%EOF\n'), 'PDF ends with %%EOF marker');

	assert.ok(pdfText.includes('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj'), 'Object 1 is Catalog');
	assert.ok(pdfText.includes('2 0 obj\n<< /Type /Pages /Count 1 /Kids [4 0 R] >>\nendobj'), 'Object 2 is Pages root');
	assert.ok(pdfText.includes('3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj'), 'Object 3 is Font');

	assert.ok(pdfText.includes('4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 450] /Resources << /XObject << /Im1 6 0 R >> /Font << /Ftxt 3 0 R >> >> /Contents 5 0 R >>\nendobj'), 'Object 4 is Page 1');
	assert.ok(pdfText.includes('5 0 obj\n<< /Length 31 >>\nstream\nq\n600 0 0 450 0 0 cm\n/Im1 Do\nQ\n\nendstream\nendobj'), 'Object 5 is Content stream');
	assert.ok(pdfText.includes('6 0 obj\n<< /Type /XObject /Subtype /Image /Width 800 /Height 600 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length 10 >>\nstream\n'), 'Object 6 is Image object');

	assert.doesNotMatch(pdfText, /\bBT\b/, 'Content stream has no text blocks when textRuns is empty');

	// Verify xref offset table byte accuracy
	const xrefMatch = pdfText.match(/xref\n0 (\d+)\n((?:[0-9]{10} [0-9]{5} [fn]\n)+)/);
	assert.ok(xrefMatch, 'xref table is present');
	const objectCount = Number.parseInt(xrefMatch[1], 10);
	assert.equal(objectCount, 7, 'Total object count is 7 for 1 page');

	const xrefLines = xrefMatch[2].trim().split('\n');
	assert.equal(xrefLines.length, 7);
	assert.equal(xrefLines[0], '0000000000 65535 f');

	for (let objNum = 1; objNum < objectCount; objNum++) {
		const line = xrefLines[objNum];
		const offset = Number.parseInt(line.slice(0, 10), 10);
		const targetHeader = `${objNum} 0 obj`;
		const actualHeader = pdfText.slice(offset, offset + targetHeader.length);
		assert.equal(actualHeader, targetHeader, `Offset for object ${objNum} (${offset}) points to exact object header`);
	}
});

test('createRenderedImagePdf handles multi-page documents with distinct resource object mappings', async () => {
	const { createRenderedImagePdf } = await loadRenderedPdfExportModule();

	const page1Image = new Uint8Array([0x01, 0x02, 0x03]);
	const page2Image = new Uint8Array([0x04, 0x05, 0x06, 0x07]);

	const pages = [
		{
			imageBytes: page1Image,
			imageWidth: 1000,
			imageHeight: 800,
			pdfWidth: 500,
			pdfHeight: 400,
			textRuns: [],
		},
		{
			imageBytes: page2Image,
			imageWidth: 1200,
			imageHeight: 900,
			pdfWidth: 600,
			pdfHeight: 450,
			textRuns: [],
		},
	];

	const pdfBuffer = createRenderedImagePdf(pages);
	const pdfText = new TextDecoder('latin1').decode(pdfBuffer);

	assert.ok(pdfText.includes('<< /Type /Pages /Count 2 /Kids [4 0 R 7 0 R] >>'), 'Pages catalog links to both page objects');

	// Page 1: obj 4 -> Content obj 5 -> Image obj 6 (/Im1)
	assert.ok(pdfText.includes('4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 400] /Resources << /XObject << /Im1 6 0 R >> /Font << /Ftxt 3 0 R >> >> /Contents 5 0 R >>'));
	assert.ok(pdfText.includes('6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1000 /Height 800'));

	// Page 2: obj 7 -> Content obj 8 -> Image obj 9 (/Im2)
	assert.ok(pdfText.includes('7 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 450] /Resources << /XObject << /Im2 9 0 R >> /Font << /Ftxt 3 0 R >> >> /Contents 8 0 R >>'));
	assert.ok(pdfText.includes('9 0 obj\n<< /Type /XObject /Subtype /Image /Width 1200 /Height 900'));

	// Check xref entries
	const xrefMatch = pdfText.match(/xref\n0 (\d+)\n((?:[0-9]{10} [0-9]{5} [fn]\n)+)/);
	assert.ok(xrefMatch);
	const objectCount = Number.parseInt(xrefMatch[1], 10);
	assert.equal(objectCount, 10, 'Total object count is 10 for 2 pages (1 catalog + 1 pages + 1 font + 2*3 page objects)');

	const xrefLines = xrefMatch[2].trim().split('\n');
	for (let objNum = 1; objNum < objectCount; objNum++) {
		const line = xrefLines[objNum];
		const offset = Number.parseInt(line.slice(0, 10), 10);
		const targetHeader = `${objNum} 0 obj`;
		const actualHeader = pdfText.slice(offset, offset + targetHeader.length);
		assert.equal(actualHeader, targetHeader, `Multi-page offset for object ${objNum} (${offset}) matches object header`);
	}
});

test('createRenderedImagePdf formats text runs with WinAnsi character encoding and escaping', async () => {
	const { createRenderedImagePdf } = await loadRenderedPdfExportModule();

	const dummyImageBytes = new Uint8Array([0x00]);
	const pages = [
		{
			imageBytes: dummyImageBytes,
			imageWidth: 100,
			imageHeight: 100,
			pdfWidth: 100,
			pdfHeight: 100,
			textRuns: [
				{
					text: 'Hello World! (Test \\ string)',
					x: 10.5,
					y: 20.25,
					fontSize: 12,
					horizontalScale: 95.5,
				},
				{
					text: 'Bullet • Dash – Quotes “Hello” ™',
					x: 10.5,
					y: 40,
					fontSize: 14,
					horizontalScale: 100,
				},
			],
		},
	];

	const pdfBuffer = createRenderedImagePdf(pages);
	const pdfText = new TextDecoder('latin1').decode(pdfBuffer);

	assert.ok(pdfText.includes('BT\n/Ftxt 1 Tf\n3 Tr'), 'Starts text block BT with font setup');
	assert.ok(pdfText.includes('100 Tz\n0 Tr\nET'), 'Ends text block ET');

	// Check escaping in first run: '(' -> '\(', ')' -> '\)', '\' -> '\\'
	assert.ok(pdfText.includes('95.5 Tz'), 'Applies horizontal scale 95.5');
	assert.ok(pdfText.includes('12 0 0 12 10.5 20.25 Tm'), 'Applies text matrix for font size 12 at x=10.5, y=20.25');
	assert.ok(pdfText.includes('(Hello World! \\(Test \\\\ string\\)) Tj'), 'Escaped parenthesis and backslash in PDF text literal');

	// Check WinAnsi character mapping in second run:
	// '•' -> 0x95 (octal \225)
	// '–' -> 0x96 (octal \226)
	// '“' -> 0x93 (octal \223)
	// '”' -> 0x94 (octal \224)
	// '™' -> 0x99 (octal \231)
	assert.ok(pdfText.includes('100 Tz'), 'Applies horizontal scale 100');
	assert.ok(pdfText.includes('14 0 0 14 10.5 40 Tm'), 'Applies text matrix for font size 14 at x=10.5, y=40');
	assert.ok(pdfText.includes('(Bullet \\225 Dash \\226 Quotes \\223Hello\\224 \\231) Tj'), 'WinAnsi characters properly mapped to octal byte codes');
});

test('dataUrlToBytes decodes standard data URLs, raw base64, empty payloads, and binary byte ranges', async () => {
	const { dataUrlToBytes } = await loadRenderedPdfExportModule();

	const helloWorldBase64 = Buffer.from('Hello World').toString('base64');
	const jpegDataUrl = `data:image/jpeg;base64,${helloWorldBase64}`;
	const bytes1 = dataUrlToBytes(jpegDataUrl);
	assert.deepEqual(bytes1, new Uint8Array(Buffer.from('Hello World')));

	const textDataUrl = `data:text/plain;charset=utf-8;base64,${helloWorldBase64}`;
	const bytes2 = dataUrlToBytes(textDataUrl);
	assert.deepEqual(bytes2, new Uint8Array(Buffer.from('Hello World')));

	const bytes3 = dataUrlToBytes(helloWorldBase64);
	assert.deepEqual(bytes3, new Uint8Array(Buffer.from('Hello World')));

	assert.deepEqual(dataUrlToBytes(''), new Uint8Array(0));
	assert.deepEqual(dataUrlToBytes('data:image/png;base64,'), new Uint8Array(0));

	const allBytes = new Uint8Array(256);
	for (let i = 0; i < 256; i++) {
		allBytes[i] = i;
	}
	const binaryBase64 = Buffer.from(allBytes).toString('base64');
	const binaryDataUrl = `data:application/octet-stream;base64,${binaryBase64}`;
	assert.deepEqual(dataUrlToBytes(binaryDataUrl), allBytes);
});
