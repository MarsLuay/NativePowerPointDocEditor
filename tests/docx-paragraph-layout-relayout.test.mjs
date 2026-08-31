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

function paragraphWithAttrs(attrs) {
	const node = {
		type: { name: 'paragraph' },
		attrs,
		descendants(callback) {
			// child isText check needs descendants to just return nothing or child nodes.
			// The original implementation had descendants() {} doing nothing for child.
			// Let's just restore it doing nothing to represent it has no children.
		},
	};
	return node;
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

test('didParagraphLayoutChange returns true if either list layout or typography changed', async () => {
	const { didParagraphLayoutChange } = await loadModule();

	const base = doc(paragraphWithAttrs({ numPr: '1', styleId: 'Normal' }));
	const listChanged = doc(paragraphWithAttrs({ numPr: '2', styleId: 'Normal' }));
	const typographyChanged = doc(paragraphWithAttrs({ numPr: '1', styleId: 'Heading1' }));
	const bothChanged = doc(paragraphWithAttrs({ numPr: '2', styleId: 'Heading1' }));
	const unchanged = doc(paragraphWithAttrs({ numPr: '1', styleId: 'Normal' }));

	assert.equal(didParagraphLayoutChange(base, unchanged), false);
	assert.equal(didParagraphLayoutChange(base, listChanged), true);
	assert.equal(didParagraphLayoutChange(base, typographyChanged), true);
	assert.equal(didParagraphLayoutChange(base, bothChanged), true);
});
