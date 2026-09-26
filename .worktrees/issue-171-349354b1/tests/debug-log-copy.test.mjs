import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

let modulePromise;
async function loadCopyModule() {
	modulePromise ??= bundleSource('src/debugLogCopy.ts', 'debug-log-copy.cjs').then((outfile) => import(`file://${outfile}`));
	return modulePromise;
}

function diagnostics(overrides = {}) {
	return {
		obsidianVersion: '1.8.10',
		obsidianApiVersion: '1.8.10',
		platform: 'Linux x86_64',
		appMode: 'desktop',
		runtime: { electron: '30.0.0', chromium: '124.0.0', node: '20.0.0' },
		userAgent: 'Mozilla/5.0 test',
		devicePixelRatio: 2,
		...overrides,
	};
}

function input(scope, logs, extra = {}) {
	return {
		generatedAt: '2026-01-01T00:00:00.000Z',
		scope,
		activeDocxPath: scope === 'docx' ? 'notes/test.docx' : undefined,
		plugin: { id: 'native-powerpoint-doc-editor', version: '1.1.15', dir: 'native-powerpoint-doc-editor' },
		settings: { locale: 'en', debugLogging: true },
		docxEditorBundle: 'main.js',
		logStats: {
			debugLoggingEnabled: true,
			maxRetainedEntries: 2000,
			retainedEntries: logs.length,
			totalEntries: logs.length,
			droppedEntries: 0,
		},
		diagnostics: diagnostics(),
		logs,
		...extra,
	};
}

function log(index, message = `log-${index}`) {
	return {
		time: `2026-01-01T00:00:${String(index % 60).padStart(2, '0')}.000Z`,
		level: 'info',
		area: 'diagnostics',
		message,
		data: { index },
	};
}

test('short copied logs preserve metadata and remain valid for every scope', async () => {
	const { buildCopiedLogPayload, MAX_COPIED_LOG_CHARACTERS } = await loadCopyModule();
	for (const scope of ['docx', 'pptx', 'all']) {
		const payload = buildCopiedLogPayload(input(scope, [log(0), log(1)]));
		const serialized = JSON.stringify(payload, null, 2);
		assert.ok(serialized.length <= MAX_COPIED_LOG_CHARACTERS);
		assert.equal(JSON.parse(serialized).scope, scope);
		assert.equal(payload.logs.at(-1).message, 'log-1');
		assert.equal(payload.diagnostics.obsidianVersion, '1.8.10');
		assert.equal(payload.diagnostics.appMode, 'desktop');
		assert.equal(payload.diagnostics.runtime.electron, '30.0.0');
	}
});

test('oversized copied logs retain the newest tail and valid JSON', async () => {
	const { buildCopiedLogPayload, MAX_COPIED_LOG_CHARACTERS } = await loadCopyModule();
	const logs = Array.from({ length: 2000 }, (_, index) => log(index, `event-${index}-${'x'.repeat(80)}`));
	const payload = buildCopiedLogPayload(input('all', logs));
	const serialized = JSON.stringify(payload, null, 2);
	assert.ok(serialized.length <= MAX_COPIED_LOG_CHARACTERS);
	assert.equal(JSON.parse(serialized).scope, 'all');
	assert.equal(payload.logs.at(-1).message, logs.at(-1).message);
	assert.equal(payload.logs[0].message, logs[logs.length - payload.logs.length].message);
	assert.equal(payload.logRetention.truncated, true);
});

test('missing optional environment fields stay explicit and machine-readable', async () => {
	const { buildCopiedLogPayload, MAX_COPIED_LOG_CHARACTERS } = await loadCopyModule();
	const payload = buildCopiedLogPayload(input('docx', [log(0)], {
		diagnostics: diagnostics({
			obsidianVersion: null,
			obsidianApiVersion: null,
			platform: null,
			appMode: 'unknown',
			runtime: { electron: null, chromium: null, node: null },
			userAgent: null,
			devicePixelRatio: null,
		}),
	}));
	assert.ok(JSON.stringify(payload).length <= MAX_COPIED_LOG_CHARACTERS);
	assert.deepEqual(payload.diagnostics.runtime, { electron: null, chromium: null, node: null });
	assert.equal(payload.diagnostics.userAgent, null);
	assert.equal(payload.diagnostics.devicePixelRatio, null);
});

test('a single oversized newest event is compacted without invalidating JSON', async () => {
	const { buildCopiedLogPayload, MAX_COPIED_LOG_CHARACTERS } = await loadCopyModule();
	const payload = buildCopiedLogPayload(input('pptx', [log(0, 'old'), log(1, 'newest-' + 'z'.repeat(100_000))]));
	const serialized = JSON.stringify(payload, null, 2);
	assert.ok(serialized.length <= MAX_COPIED_LOG_CHARACTERS);
	assert.equal(JSON.parse(serialized).logs.at(-1).time, '2026-01-01T00:00:01.000Z');
	assert.ok(payload.logs.at(-1).message.endsWith('z'.repeat(100)));
	assert.equal(payload.logRetention.truncated, true);
});
