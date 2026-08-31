import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

async function loadModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-paragraph-relayout-'));
	const outfile = path.join(outputDirectory, 'docx-paragraph-layout-relayout.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/docxParagraphLayoutRelayout.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	return require(outfile);
}

function paragraphNode(attrs = {}) {
	return {
		type: { name: 'paragraph' },
		attrs,
		descendants() {},
	};
}

function paragraph(defaultTextFormatting) {
	return {
		type: { name: 'paragraph' },
		attrs: { defaultTextFormatting },
		descendants() {},
	};
}

function doc(...nodes) {
	return {
		descendants(callback) {
			nodes.forEach((node, index) => callback(node, index));
		},
	};
}

test('empty paragraph default font formatting invalidates paragraph layout', async () => {
	const { didParagraphTypographyChange } = await loadModule();
	assert.equal(
		didParagraphTypographyChange(
			doc(paragraph({ fontSize: 8 })),
			doc(paragraph({ fontSize: 24 })),
		),
		true,
	);
});

test('unchanged empty paragraph defaults keep layout stable', async () => {
	const { didParagraphTypographyChange } = await loadModule();
	assert.equal(
		didParagraphTypographyChange(
			doc(paragraph({ fontSize: 8, fontFamily: { ascii: 'Arial' } })),
			doc(paragraph({ fontSize: 8, fontFamily: { ascii: 'Arial' } })),
		),
		false,
	);
});

test('getParagraphListLayoutSignature returns empty string when no list attributes exist', async () => {
	const { getParagraphListLayoutSignature } = await loadModule();
	assert.equal(
		getParagraphListLayoutSignature(paragraphNode({ indentLeft: 720, defaultTextFormatting: {} })),
		'',
	);
});

test('getParagraphListLayoutSignature correctly serializes all list layout attributes', async () => {
	const { getParagraphListLayoutSignature } = await loadModule();
	assert.equal(
		getParagraphListLayoutSignature(paragraphNode({
			numPr: { numId: 1, ilvl: 0 },
			listMarker: '1.',
			listMarkerHidden: false,
			listMarkerFontFamily: { ascii: 'Calibri' },
			listMarkerFontSize: 24,
			indentLeft: 720,
			indentFirstLine: 0,
			hangingIndent: 360,
		})),
		[
			JSON.stringify({ numId: 1, ilvl: 0 }),
			'1.',
			'false',
			JSON.stringify({ ascii: 'Calibri' }),
			'24',
			'720',
			'0',
			'360',
		].join('\u001f'),
	);
});

test('getParagraphListLayoutSignature handles partial list attributes with empty strings', async () => {
	const { getParagraphListLayoutSignature } = await loadModule();
	assert.equal(
		getParagraphListLayoutSignature(paragraphNode({
			numPr: { numId: 2, ilvl: 1 },
			indentLeft: 1440,
			hangingIndent: 360,
		})),
		[
			JSON.stringify({ numId: 2, ilvl: 1 }),
			'',
			'',
			'',
			'',
			'1440',
			'',
			'360',
		].join('\u001f'),
	);
});
