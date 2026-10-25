import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const require = createRequire(import.meta.url);
const {
	DUPLICATE_ENTER_WINDOW_MS,
	createDuplicateEnterGuard,
} = require(await bundleSource('src/docxDuplicateEnterGuard.ts', 'docx-duplicate-enter-guard.cjs'));

function enter(overrides = {}) {
	return {
		key: 'Enter',
		code: 'Enter',
		repeat: false,
		shiftKey: false,
		altKey: false,
		ctrlKey: false,
		metaKey: false,
		isComposing: false,
		...overrides,
	};
}

test('one ordinary Enter is allowed and the 200ms non-repeat duplicate is suppressed', () => {
	const guard = createDuplicateEnterGuard();

	assert.equal(guard.observeKeyDown(enter(), 1000, 1), 'allow');
	assert.equal(guard.observeKeyDown(enter(), 1000 + 202, 2), 'suppress');
	assert.equal(guard.observeKeyDown(enter(), 1000 + DUPLICATE_ENTER_WINDOW_MS + 1, 2), 'allow');
});

test('keyup then another Enter inside the window is a new line', () => {
	const guard = createDuplicateEnterGuard();

	assert.equal(guard.observeKeyDown(enter(), 1000, 1), 'allow');
	guard.observeKeyUp(enter());
	assert.equal(guard.observeKeyDown(enter(), 1080, 2), 'allow');
	guard.observeKeyUp(enter());
	assert.equal(guard.observeKeyDown(enter(), 1160, 3), 'allow');
	guard.observeKeyUp(enter());
	assert.equal(guard.observeKeyDown(enter(), 1240, 4), 'allow');
	guard.observeKeyUp(enter());
	assert.equal(guard.observeKeyDown(enter(), 1320, 5), 'allow');
});

test('key repeat and modified Enter stay available', () => {
	const guard = createDuplicateEnterGuard();

	assert.equal(guard.observeKeyDown(enter(), 1000, 1), 'allow');
	assert.equal(guard.observeKeyDown(enter({ repeat: true }), 1200, 2), 'allow');
	assert.equal(guard.observeKeyDown(enter({ shiftKey: true }), 1210, 2), 'allow');
	assert.equal(guard.observeKeyDown(enter({ ctrlKey: true }), 1220, 2), 'allow');
});

test('insertParagraph is suppressed only after the keydown already added a paragraph', () => {
	const guard = createDuplicateEnterGuard();

	assert.equal(guard.observeKeyDown(enter(), 1000, 1), 'allow');
	assert.equal(guard.observeBeforeInput({ inputType: 'insertParagraph' }, 1, 1010), 'allow');
	assert.equal(guard.observeBeforeInput({ inputType: 'insertParagraph' }, 2, 1200), 'suppress');
	assert.equal(guard.observeBeforeInput({ inputType: 'insertLineBreak' }, 2, 1210), 'allow');
	assert.equal(guard.observeBeforeInput({ inputType: 'insertText' }, 2, 1220), 'allow');
});

test('DOCX view listens for Enter keyup so rapid taps are not treated as iPad duplicates', async () => {
	const view = await readFile(path.join(projectRoot, 'src/DocxReactView.tsx'), 'utf8');
	assert.match(view, /addEventListener\('keyup', onKeyUp, true\)/);
	assert.match(view, /guard\.observeKeyUp\(event\)/);
	const pptx = await readFile(path.join(projectRoot, 'src/powerpoint/ui/NativePowerPointView.ts'), 'utf8');
	assert.match(pptx, /pendingParagraphSplits/);
	assert.match(pptx, /Queued paragraph split while one is pending/);
});
