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
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-font-size-target-'));
	const outfile = path.join(outputDirectory, 'docx-font-size-target.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/docxFontSizeTarget.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	return require(outfile);
}

test('formatting target keeps an explicit text selection', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 20, to: 30 },
		preservedSelection: { from: 40, to: 50 },
		caretParagraph: { from: 10, to: 60 },
	}), {
		range: { from: 20, to: 30 },
		source: 'selection',
	});
});

test('formatting target restores a range captured before toolbar focus', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 815, to: 815 },
		preservedSelection: { from: 869, to: 929 },
		caretParagraph: { from: 815, to: 815 },
	}), {
		range: { from: 869, to: 929 },
		source: 'preserved-selection',
	});
});

test('formatting target expands a caret to its current line', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 875, to: 875 },
		caretParagraph: { from: 869, to: 929 },
	}), {
		range: { from: 869, to: 929 },
		source: 'caret-paragraph',
	});
});

test('formatting target recovers the painted line when PM is on an empty spacer', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 815, to: 815 },
		caretParagraph: { from: 815, to: 815 },
		renderedParagraph: { from: 869, to: 929 },
	}), {
		range: { from: 869, to: 929 },
		source: 'rendered-paragraph',
	});
});

test('formatting target prefers the last painted line over a stale selection elsewhere', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 1, to: 12 },
		preservedSelection: { from: 1, to: 12 },
		caretParagraph: { from: 1, to: 12 },
		renderedParagraph: { from: 869, to: 929 },
	}), {
		range: { from: 869, to: 929 },
		source: 'rendered-paragraph',
	});
});

test('formatting target keeps a clicked empty paragraph instead of the previous text line', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 869, to: 929 },
		preservedSelection: { from: 869, to: 929 },
		caretParagraph: { from: 869, to: 929 },
		renderedParagraph: { from: 815, to: 815 },
	}), {
		range: { from: 815, to: 815 },
		source: 'empty-paragraph',
	});
});

test('formatting target keeps a keyboard caret in a compact resume spacer', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 1188, to: 1188 },
		preservedSelection: { from: 1188, to: 1188 },
		caretParagraph: { from: 1188, to: 1188 },
		renderedParagraph: { from: 1115, to: 1186 },
		preferCurrentSelection: true,
	}), {
		range: { from: 1188, to: 1188 },
		source: 'empty-paragraph',
	});
});

test('font-size step reads the selected text size before a stale toolbar display', async () => {
	const { resolveDocxFontSizeStepBase } = await loadModule();
	assert.equal(resolveDocxFontSizeStepBase([12, 12], 34), 12);
	assert.equal(resolveDocxFontSizeStepBase([], 34), 34);
});

test('font-family target uses the same generic resolver', async () => {
	const { resolveDocxFormattingTarget } = await loadModule();
	assert.deepEqual(resolveDocxFormattingTarget({
		selection: { from: 815, to: 815 },
		preservedSelection: { from: 869, to: 929 },
		caretParagraph: { from: 815, to: 815 },
		renderedParagraph: { from: 869, to: 929 },
	}), {
		range: { from: 869, to: 929 },
		source: 'preserved-selection',
	});
});

test('font-size resolver compatibility alias delegates to the generic resolver', async () => {
	const { resolveDocxFontSizeTarget, resolveDocxFormattingTarget } = await loadModule();
	const input = {
		selection: { from: 20, to: 30 },
		caretParagraph: { from: 10, to: 40 },
	};
	assert.deepEqual(resolveDocxFontSizeTarget(input), resolveDocxFormattingTarget(input));
});
