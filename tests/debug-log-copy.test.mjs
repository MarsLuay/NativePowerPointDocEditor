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
		editorDiagnostics: {
			docx: [],
			pptx: [],
		},
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

test('copy-time editor diagnostics preserve bounded live DOCX and PowerPoint state', async () => {
	const { buildCopiedLogPayload, MAX_COPIED_LOG_CHARACTERS } = await loadCopyModule();
	const payload = buildCopiedLogPayload(input('all', [log(0)], {
		editorDiagnostics: {
			docx: [{ path: 'notes/report.docx', dirty: true, editorMounted: true }],
			pptx: [{ path: 'slides/deck.pptx', slideCount: 12, currentSlide: 3, selectedShapeCount: 2 }],
		},
	}));
	assert.ok(JSON.stringify(payload).length <= MAX_COPIED_LOG_CHARACTERS);
	assert.equal(payload.editorDiagnostics.docx[0].dirty, true);
	assert.equal(payload.editorDiagnostics.pptx[0].currentSlide, 3);
});

test('copied logs include the resolved runtime frame profile', async () => {
	const { buildCopiedLogPayload } = await loadCopyModule();
	const frameTiming = {
		measuredRefreshHz: 120,
		measuredFrameBudgetMs: 8.33,
		resolvedRefreshHz: 120,
		resolvedFrameBudgetMs: 8.33,
		sampleCount: 32,
		confidence: 'stable',
		thresholdSource: 'measured-raf',
		platform: 'windows',
		appMode: 'desktop',
		thresholds: {
			synchronousWorkMs: 8,
			lateFrameGapMs: 12.5,
			missedFrameGapMs: 16.67,
			substantialStallMs: 25,
		},
	};
	const payload = buildCopiedLogPayload(input('all', [log(0)], {
		diagnostics: diagnostics({ frameTiming }),
	}));
	assert.deepEqual(payload.diagnostics.frameTiming, frameTiming);
});

test('copied logs include continuous interaction scroll and zoom summaries', async () => {
	const { buildCopiedLogPayload } = await loadCopyModule();
	const continuousInteractions = [
		{
			scope: 'docx',
			interactionType: 'docx-scroll',
			frameCount: 24,
			totalDurationMs: 400,
			lateFrameCount: 0,
			missedFrameEstimate: 0,
			longestFrameGapMs: 16.7,
			frameIntervalsMs: { p50: 16.6, p95: 16.7, max: 16.7 },
			rafSchedulingDelayMs: { p50: 0.5, p95: 1.0, max: 1.2 },
			synchronousWorkMs: { p50: 2.1, p95: 3.5, max: 4.0 },
			eventLoopDelayMs: { p50: 0.2, p95: 0.4, max: 0.5 },
			worstFrames: [],
			frameBudgetMs: 16.67,
			resolvedRefreshHz: 60,
		},
	];
	const payload = buildCopiedLogPayload(input('all', [log(0)], {
		diagnostics: diagnostics({ continuousInteractions }),
	}));
	assert.deepEqual(payload.diagnostics.continuousInteractions, continuousInteractions);
});

test('copied logs include session resource and memory diagnostics', async () => {
	const { buildCopiedLogPayload } = await loadCopyModule();
	const resourceDiagnostics = {
		timestamp: 12345678,
		label: 'test-snapshot',
		memory: {
			supported: false,
			provider: 'unsupported',
			usedBytes: null,
			totalBytes: null,
			limitBytes: null,
		},
		resources: {
			mountedDocxViews: 1,
			mountedPptxViews: 0,
			activeTimers: 2,
			activeAnimationFrames: 0,
			activeMutationObservers: 1,
			activeResizeObservers: 0,
			thumbnailCacheEntries: 0,
			domNodeCount: 150,
			registeredListeners: 5,
		},
	};
	const payload = buildCopiedLogPayload(input('all', [log(0)], {
		diagnostics: diagnostics({ resourceDiagnostics }),
	}));
	assert.deepEqual(payload.diagnostics.resourceDiagnostics, resourceDiagnostics);
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

test('DOCX input traces survive newer unrelated log spam', async () => {
	const { buildCopiedLogPayload } = await loadCopyModule();
	const inputTrace = {
		time: '2026-01-01T00:00:00.000Z',
		level: 'debug',
		area: 'text-input',
		message: 'DOCX input event observed',
		data: { correlationId: 'docx-input-plugin-1:correlation-1', key: 'Enter' },
	};
	const spam = Array.from({ length: 2000 }, (_, index) => log(index + 1, `caret-${index}-${'y'.repeat(80)}`));
	const payload = buildCopiedLogPayload(input('docx', [inputTrace, ...spam]));
	assert.equal(payload.logs.some((entry) => entry.message === 'DOCX input event observed'), true);
	assert.equal(payload.logs.at(-1).message, spam.at(-1).message);
});
