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

const stubObsidianPlugin = {
	name: 'stub-obsidian',
	setup(buildContext) {
		buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: 'obsidian', namespace: 'stub-obsidian' }));
		buildContext.onLoad({ filter: /.*/, namespace: 'stub-obsidian' }, () => ({
			contents: 'module.exports = { Notice: class Notice {} };',
			loader: 'js',
		}));
	},
};

async function loadRegisterAiCommandsModule() {
	if (cachedModule) return cachedModule;
	const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-ai-reg-test-'));
	const outfile = path.join(outputDirectory, 'registerAiCommands.cjs');
	await build({
		entryPoints: [path.join(projectRoot, 'src/ai/registerAiCommands.ts')],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		plugins: [stubObsidianPlugin],
	});
	cachedModule = require(outfile);
	return cachedModule;
}

test('registerAiCommands registers all standard and legacy commands', async () => {
	const { registerAiCommands } = await loadRegisterAiCommandsModule();

	const registeredCommands = [];
	const mockPlugin = {
		addCommand: (cmd) => {
			registeredCommands.push(cmd);
		},
		app: {
			workspace: {
				getActiveFile: () => null,
			},
		},
	};

	const mockGetI18n = () => null;
	const mockGetAi = () => undefined;

	registerAiCommands({
		plugin: mockPlugin,
		getI18n: mockGetI18n,
		getAi: mockGetAi,
	});

	assert.equal(registeredCommands.length, 10);

	const commandIds = registeredCommands.map((c) => c.id);
	assert.ok(commandIds.includes('npde-ai-capabilities'));
	assert.ok(commandIds.includes('npde-ai-describe'));
	assert.ok(commandIds.includes('npde-ai-apply'));
	assert.ok(commandIds.includes('npde-ai-validate'));
	assert.ok(commandIds.includes('npde-ai-save'));
	assert.ok(commandIds.includes('npde-ai-undo'));
	assert.ok(commandIds.includes('npde-ai-redo'));

	assert.ok(commandIds.includes('npde-ai-list-capabilities'));
	assert.ok(commandIds.includes('npde-ai-describe-document'));
	assert.ok(commandIds.includes('npde-ai-apply-operations'));

	// Test callback execution when AI is disabled
	const capCmd = registeredCommands.find((c) => c.id === 'npde-ai-capabilities');
	await capCmd.callback();
});
