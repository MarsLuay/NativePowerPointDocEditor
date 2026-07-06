import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import Module, { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

async function loadModule(entry, external = []) {
	const outdir = await mkdtemp(path.join(tmpdir(), 'npde-editor-language-test-'));
	const outfile = path.join(outdir, 'module.cjs');
	await build({
		absWorkingDir: projectRoot,
		entryPoints: [entry],
		bundle: true,
		external,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		loader: { '.json': 'json' },
	});
	return import(pathToFileURL(outfile).href);
}

function loadWithObsidianStub(outfile) {
	const originalLoad = Module._load;
	Module._load = function load(request, parent, isMain) {
		if (request === 'obsidian') {
			return {
				App: class App {},
				Notice: class Notice {},
				PluginSettingTab: class PluginSettingTab {},
				Setting: class Setting {},
			};
		}
		return originalLoad.call(this, request, parent, isMain);
	};
	try {
		return require(outfile);
	} finally {
		Module._load = originalLoad;
	}
}

async function loadSettingsModule() {
	const outdir = await mkdtemp(path.join(tmpdir(), 'npde-editor-language-settings-'));
	const outfile = path.join(outdir, 'settings.cjs');
	await build({
		absWorkingDir: projectRoot,
		entryPoints: ['src/settings.ts'],
		bundle: true,
		external: ['obsidian'],
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		loader: { '.json': 'json' },
	});
	return loadWithObsidianStub(outfile);
}

test('DEFAULT_SETTINGS does not define editorLanguage', async () => {
	const { DEFAULT_SETTINGS } = await loadSettingsModule();
	assert.equal('editorLanguage' in DEFAULT_SETTINGS, false);
});

test('readNativePowerPointDocEditorSettings drops legacy editorLanguage', async () => {
	const { readNativePowerPointDocEditorSettings } = await loadSettingsModule();
	const result = readNativePowerPointDocEditorSettings({
		authorName: 'Legacy User',
		editorLanguage: 'pl',
		autosave: true,
	});

	assert.equal(result.hadLegacyEditorLanguage, true);
	assert.equal(result.shouldPersistSettings, true);
	assert.equal('editorLanguage' in result.settings, false);
	assert.equal(result.settings.authorName, 'Legacy User');
});

test('resolveAutomaticLocale matches exact, base, then English', async () => {
	const { resolveAutomaticLocale } = await loadModule('src/i18n/localeResolver.ts');
	const available = ['en', 'pl'];

	assert.equal(resolveAutomaticLocale('pl', available), 'pl');
	assert.equal(resolveAutomaticLocale('pl-PL', available), 'pl');
	assert.equal(resolveAutomaticLocale('de', available), 'en');
	assert.equal(resolveAutomaticLocale('en-US', available), 'en');
});

test('DOCX editor language stays automatic even when plugin UI falls back', async () => {
	const { resolveAutomaticLocale } = await loadModule('src/i18n/localeResolver.ts');
	const { resolveAutomaticDocxEditorLanguage } = await loadModule('src/locales.ts');

	const pluginLocale = resolveAutomaticLocale('pt-BR', ['en', 'pl']);
	const docxLanguage = resolveAutomaticDocxEditorLanguage('pt-BR');

	assert.equal(pluginLocale, 'en');
	assert.equal(docxLanguage, 'pt-BR');
	assert.equal(resolveAutomaticDocxEditorLanguage('tr'), 'tr');
	assert.equal(resolveAutomaticDocxEditorLanguage('de'), 'en');
});
