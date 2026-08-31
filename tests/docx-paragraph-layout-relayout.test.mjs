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

function paragraph(defaultTextFormatting) {
	return {
		type: { name: 'paragraph' },
		attrs: { defaultTextFormatting },
		descendants() {},
	};
}

function listParagraph(listLayout) {
	return {
		type: { name: 'paragraph' },
		attrs: { ...listLayout },
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

test('changing paragraph list layout invalidates list layout', async () => {
	const { didListLayoutChange } = await loadModule();
	assert.equal(
		didListLayoutChange(
			doc(listParagraph({ numPr: { numId: '1', ilvl: '0' } })),
			doc(listParagraph({ numPr: { numId: '2', ilvl: '0' } })),
		),
		true,
	);
});

test('unchanged paragraph list layout keeps layout stable', async () => {
	const { didListLayoutChange } = await loadModule();
	assert.equal(
		didListLayoutChange(
			doc(listParagraph({ numPr: { numId: '1', ilvl: '0' } })),
			doc(listParagraph({ numPr: { numId: '1', ilvl: '0' } })),
		),
		false,
	);
});
