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
