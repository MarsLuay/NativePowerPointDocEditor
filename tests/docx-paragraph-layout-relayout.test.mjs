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
		descendants(callback) {
			// A paragraph typically descends into text nodes, but we'll leave it empty for this mock.
			return true;
		},
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
			for (let i = 0; i < nodes.length; i++) {
				const node = nodes[i];
				const shouldContinue = callback(node, i);
				if (shouldContinue && node.descendants) {
					node.descendants(callback);
				}
			}
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

test('getDocumentParagraphLayoutSignatures extracts signatures for paragraph nodes only', async () => {
	const { getDocumentParagraphLayoutSignatures } = await loadModule();
	const mixedDoc = doc(
		paragraph({ fontSize: 12 }),
		{ type: { name: 'heading' }, descendants() {} },
		paragraph({ fontSize: 14 }),
	);
	const signatures = getDocumentParagraphLayoutSignatures(mixedDoc);
	assert.equal(signatures.length, 2);
	assert.equal(typeof signatures[0], 'string');
	assert.equal(typeof signatures[1], 'string');
	assert.notEqual(signatures[0], signatures[1]);
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
