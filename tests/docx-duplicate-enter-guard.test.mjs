import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import test from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

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
