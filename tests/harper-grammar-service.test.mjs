import assert from 'node:assert/strict';
import test from 'node:test';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const { createHarperGrammarService } = require(await bundleSource(
	'src/harper/harperGrammarService.ts',
	'harper-grammar-service.cjs',
));

function lintResult(message, replacement) {
	return {
		span: () => ({ start: 8, end: 9 }),
		message: () => message,
		suggestions: () => [{
			kind: () => 0,
			get_replacement_text: () => replacement,
		}],
	};
}

function createHarness() {
	const logs = [];
	const lintCalls = [];
	let setups = 0;
	let defaultConfigs = 0;
	let disposals = 0;
	let created = 0;
	let releaseLint = null;
	const timers = [];
	const service = createHarperGrammarService({
		debounceMs: 250,
		now: () => 1_000,
		log: (entry) => logs.push(entry),
		schedule(callback) {
			const timer = { callback, cancelled: false };
			timers.push(timer);
			return { cancel: () => { timer.cancelled = true; } };
		},
		createLinter: () => {
			created += 1;
			return {
				async setup() { setups += 1; },
				async getDefaultLintConfig() { defaultConfigs += 1; },
				lint(text) {
					lintCalls.push(text);
					return new Promise((resolve) => {
						releaseLint = () => resolve([lintResult('Use an', 'an')]);
					});
				},
				async dispose() { disposals += 1; },
			};
		},
	});
	return {
		service,
		logs,
		lintCalls,
		timers,
		counts: () => ({ created, setups, defaultConfigs, disposals }),
		async flush() {
			const timer = timers.pop();
			assert.ok(timer && !timer.cancelled);
			timer.callback();
			for (let attempt = 0; attempt < 10 && typeof releaseLint !== 'function'; attempt += 1) {
				await Promise.resolve();
			}
		},
		release() {
			const finish = releaseLint;
			assert.equal(typeof finish, 'function');
			finish();
		},
	};
}

test('one Harper worker is reused and source text stays out of logs', async () => {
	const harness = createHarness();
	const first = harness.service.requestLint('SECRET first sentence.');
	const second = harness.service.requestLint('SECRET second sentence.');
	assert.equal(harness.timers.filter((timer) => !timer.cancelled).length, 1);
	await harness.flush();
	harness.release();
	const result = await second;
	assert.equal(await first, null);
	assert.equal(result.length, 1);
	assert.equal(result[0].span.start, 8);
	assert.equal(result[0].suggestions[0].replacement, 'an');
	assert.equal(harness.lintCalls.length, 1);
	assert.equal(harness.lintCalls[0], 'SECRET second sentence.');
	assert.deepEqual(harness.counts(), { created: 1, setups: 1, defaultConfigs: 1, disposals: 0 });

	const third = harness.service.requestLint('SECRET third sentence.');
	await harness.flush();
	harness.release();
	await third;
	assert.deepEqual(harness.counts(), { created: 1, setups: 1, defaultConfigs: 1, disposals: 0 });
	assert.equal(JSON.stringify(harness.logs).includes('SECRET'), false);
	assert.equal(typeof harness.logs[0].data.durationMs, 'number');
});

test('a slower Harper result is dropped when a newer request supersedes it', async () => {
	const harness = createHarness();
	const stale = harness.service.requestLint('SECRET stale sentence.');
	await harness.flush();
	const current = harness.service.requestLint('SECRET current sentence.');
	assert.equal(await stale, null);
	await harness.flush();
	harness.release();
	harness.release();
	assert.deepEqual(await current, [{
		span: { start: 8, end: 9 },
		message: 'Use an',
		suggestions: [{ kind: 'replace', replacement: 'an' }],
	}]);
});

test('disable and dispose cancel pending Harper work and free the worker', async () => {
	const harness = createHarness();
	const pending = harness.service.requestLint('SECRET pending sentence.');
	harness.service.disable();
	assert.equal(await pending, null);
	assert.equal(await harness.service.requestLint('SECRET later sentence.'), null);
	await harness.service.dispose();
	await harness.service.dispose();
	assert.equal(harness.counts().disposals, 0);
	const ready = createHarness();
	const lint = ready.service.requestLint('SECRET ready sentence.');
	await ready.flush();
	ready.release();
	await lint;
	await ready.service.dispose();
	assert.equal(ready.counts().disposals, 1);
	assert.equal(await ready.service.requestLint('SECRET after dispose.'), null);
});

test('Harper worker initialization failure is logged without the document text', async () => {
	const logs = [];
	const service = createHarperGrammarService({
		debounceMs: 0,
		schedule(callback) {
			callback();
			return { cancel() {} };
		},
		log: (entry) => logs.push(entry),
		createLinter: () => ({
			setup() { return Promise.reject(new Error('wasm failed')); },
			getDefaultLintConfig() { return Promise.resolve({}); },
			lint() { return Promise.resolve([]); },
			dispose() { return Promise.resolve(); },
		}),
	});
	const result = await service.requestLint('SECRET document');
	assert.equal(result, null);
	assert.equal(logs[0].level, 'error');
	assert.equal(logs[0].data.phase, 'initialize');
	assert.equal(JSON.stringify(logs).includes('SECRET document'), false);
	assert.match(logs[0].data.error, /wasm failed/);
});
