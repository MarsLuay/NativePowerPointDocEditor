import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

async function loadModule() {
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-text-selection-memory-'));
	const outfile = path.join(outputDirectory, 'docx-text-selection-memory.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/docxTextSelectionMemory.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	return require(outfile);
}

test('toolbar focus retains a selected text range', async () => {
	const { updateDocxPreservedTextSelection } = await loadModule();
	assert.deepEqual(updateDocxPreservedTextSelection(
		{ from: 100, to: 120 },
		null,
		false,
	), { from: 100, to: 120 });
});

test('editor navigation clears a stale text range before an empty paragraph is formatted', async () => {
	const {
		shouldResetDocxPreservedTextSelection,
		updateDocxPreservedTextSelection,
	} = await loadModule();
	const reset = shouldResetDocxPreservedTextSelection({
		eventType: 'pointerdown',
		insideEditorPages: true,
		insideHiddenProseMirror: false,
	});

	assert.equal(reset, true);
	assert.equal(updateDocxPreservedTextSelection(
		{ from: 100, to: 120 },
		null,
		reset,
	), null);
});

test('hidden editor keyboard navigation clears a stale text range', async () => {
	const { shouldResetDocxPreservedTextSelection } = await loadModule();
	assert.equal(shouldResetDocxPreservedTextSelection({
		eventType: 'keydown',
		insideEditorPages: false,
		insideHiddenProseMirror: true,
	}), true);
	assert.equal(shouldResetDocxPreservedTextSelection({
		eventType: 'pointerdown',
		insideEditorPages: false,
		insideHiddenProseMirror: true,
	}), false);
});

test('portaled font choices apply only to the active DOCX leaf', async () => {
	const source = await readFile(path.join(projectRoot, 'src/DocxReactView.tsx'), 'utf8');
	assert.match(source, /editorRoot\?\.closest\('\.workspace-leaf\.mod-active'\)/);
});

test('empty-paragraph font changes preserve existing default marks', async () => {
	const source = await readFile(path.join(projectRoot, 'src/DocxReactView.tsx'), 'utf8');
	const prepareStart = source.indexOf('function prepareFormattingSelection');
	const applyStart = source.indexOf('function applyFontFamilyToEditorView');
	const applyEnd = source.indexOf('function tagFontFamilySelectTrigger');
	const prepareBody = source.slice(prepareStart, applyStart);
	const applyBody = source.slice(applyStart, applyEnd);
	assert.match(prepareBody, /seedEmptyParagraphStoredMarks\(view\)/);
	assert.match(applyBody, /prepareFormattingSelection\(/);
	assert.match(applyBody, /setFontFamily\(fontFamily\)\(view\.state, view\.dispatch\)/);
});
