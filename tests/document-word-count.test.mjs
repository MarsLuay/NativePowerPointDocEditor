import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadDocumentWordCountModule } from './helpers/load-plugin-modules.mjs';

test('countDocumentWords counts whitespace-delimited document words', async () => {
	const { countDocumentWords } = await loadDocumentWordCountModule();

	assert.equal(countDocumentWords(''), 0);
	assert.equal(countDocumentWords('   \n\t '), 0, 'only whitespace returns 0');
	assert.equal(countDocumentWords('  one\ttwo\nthree  '), 3);
	assert.equal(countDocumentWords('well-known Unicode\u00a0text'), 3);
	assert.equal(countDocumentWords('Hello, world!'), 2, 'punctuation is retained and forms words');
	assert.equal(countDocumentWords('Complex text 👨‍👩‍👧‍👦 with emojis!'), 5, 'handles emojis correctly');
});

test('formatDocumentWordCount shows the selection count when one exists', async () => {
	const { formatDocumentWordCount } = await loadDocumentWordCountModule();

	assert.equal(formatDocumentWordCount({ totalWords: 0, selectedWords: null }), '0 words');
	assert.equal(formatDocumentWordCount({ totalWords: 1, selectedWords: null }), '1 word');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: null }), '1,200 words');

	assert.equal(formatDocumentWordCount({ totalWords: 10, selectedWords: 0 }), '0 selected words');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 1 }), '1 selected word');
	assert.equal(formatDocumentWordCount({ totalWords: 1200, selectedWords: 4 }), '4 selected words');
});
