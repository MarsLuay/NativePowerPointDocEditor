import assert from 'node:assert/strict';
import { test } from 'node:test';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { deployPluginArtifacts } from '../scripts/lib/deploy-plugin-artifacts.mjs';

async function writeTree(root, files) {
	for (const [relativePath, contents] of Object.entries(files)) {
		const target = path.join(root, relativePath);
		await mkdir(path.dirname(target), { recursive: true });
		await writeFile(target, contents);
	}
}

test('deployPluginArtifacts replaces generated trees and preserves runtime data', async () => {
	const root = await mkdtemp(path.join(tmpdir(), 'npde-deploy-test-'));
	const sourceDir = path.join(root, 'source');
	const targetDir = path.join(root, 'plugin');
	await writeTree(sourceDir, {
		'main.js': 'new-main',
		'styles.css': 'new-styles',
		'manifest.json': '{"version":"test"}',
		'ai/capabilities.json': 'new-capabilities',
		'locales/en.json': 'new-locale',
	});
	await writeTree(targetDir, {
		'main.js': 'old-main',
		'styles.css': 'old-styles',
		'manifest.json': '{"version":"old"}',
		'ai/stale.json': 'stale',
		'locales/old.json': 'old-locale',
		'data.json': '{"userSetting":true}',
	});

	await deployPluginArtifacts({
		sourceDir,
		targetDir,
		files: ['main.js', 'styles.css', 'manifest.json'],
		directories: ['ai', 'locales'],
	});

	assert.equal(await readFile(path.join(targetDir, 'main.js'), 'utf8'), 'new-main');
	assert.equal(await readFile(path.join(targetDir, 'ai/capabilities.json'), 'utf8'), 'new-capabilities');
	await assert.rejects(readFile(path.join(targetDir, 'ai/stale.json')));
	assert.equal(await readFile(path.join(targetDir, 'data.json'), 'utf8'), '{"userSetting":true}');
});

test('deployPluginArtifacts leaves the installed plugin untouched when staging fails', async () => {
	const root = await mkdtemp(path.join(tmpdir(), 'npde-deploy-failure-test-'));
	const sourceDir = path.join(root, 'source');
	const targetDir = path.join(root, 'plugin');
	await writeTree(sourceDir, { 'main.js': 'new-main' });
	await writeTree(targetDir, { 'main.js': 'old-main', 'data.json': '{"userSetting":true}' });

	await assert.rejects(deployPluginArtifacts({
		sourceDir,
		targetDir,
		files: ['main.js', 'manifest.json'],
		directories: [],
	}));

	assert.equal(await readFile(path.join(targetDir, 'main.js'), 'utf8'), 'old-main');
	assert.equal(await readFile(path.join(targetDir, 'data.json'), 'utf8'), '{"userSetting":true}');
});
