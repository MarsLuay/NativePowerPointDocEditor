import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import test from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const {
	CARET_THEME_LOG_MESSAGE,
	CARET_THEME_SUPPRESSED_MESSAGE,
	createCaretThemeDiagnostics,
} = require(await bundleSource('src/docxCaretThemeDiagnostics.ts', 'docx-caret-theme-diagnostics.cjs'));

function snapshot(overrides = {}) {
	return {
		file: 'notes/doc.docx',
		wantDark: false,
		roots: 1,
		docCaretVar: '#000000',
		caretBackground: 'rgb(0, 0, 0)',
		rootHasDarkClass: false,
		...overrides,
	};
}

function collect() {
	const entries = [];
	const diagnostics = createCaretThemeDiagnostics((entry) => {
		entries.push(entry);
	});
	return { diagnostics, entries };
}

test('repeated identical caret syncs emit one state record', () => {
	const { diagnostics, entries } = collect();
	const state = snapshot();

	diagnostics.record(state);
	diagnostics.record(state);
	diagnostics.record({ ...state });

	assert.equal(entries.length, 1);
	assert.equal(entries[0].message, CARET_THEME_LOG_MESSAGE);
	assert.equal(entries[0].data.wantDark, false);
	assert.equal(entries[0].data.docCaretVar, '#000000');
});

test('a caret or theme change emits a new record', () => {
	const { diagnostics, entries } = collect();

	diagnostics.record(snapshot());
	diagnostics.record(snapshot({ wantDark: true, rootHasDarkClass: true, docCaretVar: '#ffffff' }));

	const pinned = entries.filter((entry) => entry.message === CARET_THEME_LOG_MESSAGE);
	assert.equal(pinned.length, 2);
	assert.equal(pinned[1].data.wantDark, true);
	assert.equal(pinned[1].data.docCaretVar, '#ffffff');
	assert.equal(pinned[1].data.rootHasDarkClass, true);
});

test('suppressed identical caret syncs collapse into one summary', () => {
	const { diagnostics, entries } = collect();

	diagnostics.record(snapshot());
	diagnostics.record(snapshot());
	diagnostics.record(snapshot());
	diagnostics.record(snapshot({ caretBackground: 'rgb(255, 255, 255)' }));

	const summaries = entries.filter((entry) => entry.message === CARET_THEME_SUPPRESSED_MESSAGE);
	assert.equal(summaries.length, 1);
	assert.equal(summaries[0].data.suppressed, 2);
	assert.equal(entries.filter((entry) => entry.message === CARET_THEME_LOG_MESSAGE).length, 2);
});
