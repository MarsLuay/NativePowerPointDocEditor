import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const projectRoot = path.resolve(import.meta.dirname, '..');
const require = createRequire(import.meta.url);

let cachedModule;

async function loadWordCountModule() {
	if (cachedModule) return cachedModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-word-count-'));
	const outfile = path.join(outputDirectory, 'document-word-count.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/documentWordCount.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
	});
	cachedModule = require(outfile);
	return cachedModule;
}

test('countDocumentWords counts whitespace-delimited document words', async () => {
	const { countDocumentWords } = await loadWordCountModule();

	assert.equal(countDocumentWords(''), 0);
	assert.equal(countDocumentWords('  one\ttwo\nthree  '), 3);
	assert.equal(countDocumentWords('well-known Unicode\u00a0text'), 3);
});

test('formatDocumentWordCount shows the selection count when one exists', async () => {
	const { formatDocumentWordCount } = await loadWordCountModule();

	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: null }), '1,200 words');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 1 }), '1 selected word');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 4 }), '4 selected words');
});
