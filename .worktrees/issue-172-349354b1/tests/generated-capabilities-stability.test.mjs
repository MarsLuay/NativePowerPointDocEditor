import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const execFileAsync = promisify(execFile);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputPath = path.join(projectRoot, 'ai', 'capabilities.json');

test('AI capability generation is byte-stable when semantics are unchanged', async () => {
	await execFileAsync(process.execPath, ['scripts/generate-ai-capabilities.mjs'], { cwd: projectRoot });
	const first = await readFile(outputPath, 'utf8');
	await execFileAsync(process.execPath, ['scripts/generate-ai-capabilities.mjs'], { cwd: projectRoot });
	const second = await readFile(outputPath, 'utf8');

	assert.equal(second, first);
});
