import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadDocumentWordCountModule } from './helpers/load-plugin-modules.mjs';

test('countDocumentWords counts whitespace-delimited document words', async () => {
	const { countDocumentWords } = await loadDocumentWordCountModule();

	assert.equal(countDocumentWords(''), 0);
	assert.equal(countDocumentWords('  one\ttwo\nthree  '), 3);
	assert.equal(countDocumentWords('well-known Unicode\u00a0text'), 3);
});

test('formatDocumentWordCount shows the selection count when one exists', async () => {
	const { formatDocumentWordCount } = await loadDocumentWordCountModule();

	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: null }), '1,200 words');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 1 }), '1 selected word');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 4 }), '4 selected words');
});
