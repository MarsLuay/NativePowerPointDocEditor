import assert from 'node:assert/strict';
import test from 'node:test';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const grammar = require(await bundleSource('src/harper/pptxGrammarEditing.ts', 'pptx-grammar-editing.cjs'));

function lintForWrld() {
	return [{
		span: { start: 6, end: 10 },
		message: 'Use world',
		suggestions: [{ kind: 'replace', replacement: 'world' }],
	}];
}

function deferredSession(options = {}) {
	const calls = [];
	const logs = [];
	const reviews = [];
	let release = null;
	let enabled = options.enabled ?? true;
	const timers = [];
	const session = grammar.createPptxTextGrammarSession({
		debounceMs: 200,
		getEnabled: () => enabled,
		requestLint: async (text) => {
			calls.push(text);
			return new Promise((resolve) => {
				release = () => resolve(lintForWrld());
			});
		},
		schedule(callback) {
			const timer = { callback, cancelled: false };
			timers.push(timer);
			return { cancel: () => { timer.cancelled = true; } };
		},
		onReview: (diagnostics) => reviews.push(diagnostics),
		log: (data) => {
			logs.push(data);
			assert.equal(JSON.stringify(data).includes('wrld'), false);
		},
	});
	return {
		session,
		calls,
		logs,
		reviews,
		timers,
		setEnabled: (value) => { enabled = value; },
		async flush() {
			const timer = timers.pop();
			assert.ok(timer && !timer.cancelled);
			timer.callback();
			for (let attempt = 0; attempt < 8 && typeof release !== 'function'; attempt += 1) {
				await Promise.resolve();
			}
		},
		finish() {
			const done = release;
			assert.equal(typeof done, 'function');
			done();
			release = null;
		},
		async settled() {
			for (let attempt = 0; attempt < 12; attempt += 1) await Promise.resolve();
		},
	};
}

test('a PowerPoint suggestion uses textarea offsets and survives save and undo', () => {
	const history = [];
	let saved = 'Hello wrld';
	const edited = grammar.applyPptxGrammarEdit({
		text: saved,
		start: 6,
		end: 10,
		suggestion: { kind: 'replace', replacement: 'world' },
		history,
		save: (text) => { saved = text; },
	});
	assert.equal(edited.saved, 'Hello world');
	assert.equal(saved, 'Hello world');
	assert.deepEqual(history, ['Hello wrld']);
	saved = history.pop();
	assert.equal(saved, 'Hello wrld');
});

test('only the active text box is linted and a stale result is dropped', async () => {
	const calls = [];
	const pending = [];
	const logs = [];
	const session = grammar.createPptxTextGrammarSession({
		debounceMs: 0,
		getEnabled: () => true,
		requestLint: (text) => {
			calls.push(text);
			return new Promise((resolve) => pending.push(resolve));
		},
		schedule(callback) {
			callback();
			return { cancel() {} };
		},
		log: (data) => logs.push(data),
	});

	session.noteText('Hello wrld');
	session.noteText('Hello wrld!');
	assert.deepEqual(calls, ['Hello wrld', 'Hello wrld!']);
	pending[0]([]);
	pending[1]([{
		span: { start: 6, end: 10 },
		message: 'Use world',
		suggestions: [{ kind: 'replace', replacement: 'world' }],
	}]);
	for (let attempt = 0; attempt < 12; attempt += 1) await Promise.resolve();
	assert.equal(session.diagnostics().length, 1);
	assert.equal(session.diagnostics()[0].start, 6);
	assert.equal(JSON.stringify(logs).includes('wrld'), false);
	assert.equal(calls.length, 2);
});

test('composition and disable clear PowerPoint grammar diagnostics', async () => {
	const harness = deferredSession();
	harness.session.setComposing(true);
	harness.session.noteText('Hello wrld');
	assert.equal(harness.calls.length, 0);
	harness.session.setComposing(false);
	harness.session.noteText('Hello wrld');
	await harness.flush();
	harness.finish();
	await harness.settled();
	assert.equal(harness.session.diagnostics().length, 1);

	harness.setEnabled(false);
	harness.session.setEnabled(false);
	assert.equal(harness.session.diagnostics().length, 0);
	harness.session.noteText('Hello wrld');
	assert.equal(harness.calls.length, 1);

	harness.setEnabled(true);
	harness.session.noteText('Hello wrld');
	await harness.flush();
	harness.finish();
	await harness.settled();
	const applied = harness.session.apply(harness.session.diagnostics()[0].id, 0);
	assert.equal(applied, 'Hello world');
	harness.session.clear();
	assert.equal(harness.session.diagnostics().length, 0);
});
