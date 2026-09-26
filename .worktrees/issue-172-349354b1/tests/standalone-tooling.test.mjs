import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

const root = new URL('../', import.meta.url);

test('ESLint logic rules resolve inside a standalone plugin checkout', async () => {
	const config = await readFile(new URL('eslint.config.mts', root), 'utf8');
	assert.doesNotMatch(config, /from ['"]\.\.\/\.\.\/scripts\//);
	await readFile(new URL('scripts/lib/obsidian-logic-eslint-rules.mjs', root), 'utf8');
});
