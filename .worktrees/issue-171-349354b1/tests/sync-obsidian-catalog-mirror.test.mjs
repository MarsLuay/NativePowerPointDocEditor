import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const syncScript = path.join(projectRoot, 'scripts', 'sync-obsidian-catalog-mirror.mjs');

test('catalog mirror sync preserves a destination git checkout when source uses a gitfile', () => {
	const destination = fs.mkdtempSync(path.join(os.tmpdir(), 'npde-catalog-mirror-'));
	try {
		assert.equal(spawnSync('git', ['init', '--quiet'], { cwd: destination }).status, 0);
		const result = spawnSync(process.execPath, [syncScript, destination], {
			cwd: projectRoot,
			encoding: 'utf8',
		});
		assert.equal(result.status, 0, result.stderr || result.stdout);
		assert.equal(fs.statSync(path.join(destination, '.git')).isDirectory(), true);
		assert.equal(
			spawnSync('git', ['rev-parse', '--is-inside-work-tree'], {
				cwd: destination,
				encoding: 'utf8',
			}).stdout.trim(),
			'true',
		);
	} finally {
		fs.rmSync(destination, { recursive: true, force: true });
	}
});
