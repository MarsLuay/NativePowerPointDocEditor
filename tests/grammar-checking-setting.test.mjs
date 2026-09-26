import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { bundleSource } from './helpers/load-plugin-modules.mjs';

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const harper = require(await bundleSource('src/harper/harperGrammarService.ts', 'harper-grammar-setting.cjs'));
const coordinatorModule = require(await bundleSource('src/harper/grammarCheckingCoordinator.ts', 'grammar-checking-coordinator.cjs'));

test('old settings migrate grammar checking on and keep other preferences', async () => {
	const source = await readFile(path.join(projectRoot, 'src/settings.ts'), 'utf8');
	assert.match(source, /enableGrammarChecking: true/);
	assert.match(source, /raw\.enableGrammarChecking !== false/);
	assert.match(source, /raw\.enableGrammarChecking !== normalizedEnableGrammarChecking/);
	assert.match(source, /enableGrammarChecking: normalizedEnableGrammarChecking/);
	const savedFields = source.slice(source.indexOf('const settings: NativePowerPointDocEditorSettings'), source.indexOf('const shouldPersistSettings'));
	assert.match(savedFields, /authorName: readString\(raw\.authorName/);
	assert.match(savedFields, /showRuler: raw\.showRuler === true/);
});

test('grammar checking is one settings-tab toggle and is not copied into the DOCX menu', async () => {
	const catalog = await readFile(path.join(projectRoot, 'src/i18n/settingsCatalog.ts'), 'utf8');
	const tabSections = catalog.slice(catalog.indexOf('export function getNativePowerPointDocEditorSettingsTabSections'), catalog.indexOf('export function getDocxEditorSettingSectionLabels'));
	assert.match(tabSections, /\['editorTheme', 'showRuler', 'enableGrammarChecking', 'defaultZoom'\]/);
	const docxMenu = catalog.slice(catalog.indexOf('export function getDocxEditorSettingsMenuSections'));
	assert.doesNotMatch(docxMenu, /enableGrammarChecking/);
	const english = JSON.parse(await readFile(path.join(projectRoot, 'locales/en/settings.json'), 'utf8'));
	const polish = JSON.parse(await readFile(path.join(projectRoot, 'locales/pl/settings.json'), 'utf8'));
	assert.equal(english.grammar.enableGrammarChecking.name, 'Grammar checking');
	assert.equal(typeof polish.grammar.enableGrammarChecking.name, 'string');
});

test('disabling grammar checking clears diagnostics and reuses one worker after reenable', async () => {
	let created = 0;
	let cleared = 0;
	const service = harper.createHarperGrammarService({
		debounceMs: 0,
		now: () => 50,
		schedule(callback) {
			callback();
			return { cancel() {} };
		},
		createLinter: () => {
			created += 1;
			return {
				setup: async () => {},
				getDefaultLintConfig: async () => ({}),
				lint: async () => [{
					span: () => ({ start: 0, end: 1 }),
					message: () => 'Check',
					suggestions: () => [],
				}],
				dispose: async () => {},
			};
		},
	});
	const coordinator = coordinatorModule.createGrammarCheckingCoordinator({
		service,
		clearDiagnostics: () => { cleared += 1; },
	});

	const first = await coordinator.requestLint('Hello wrld');
	assert.equal(first.length, 1);
	coordinator.setEnabled(false);
	assert.equal(cleared, 1);
	assert.equal(await coordinator.requestLint('Hello wrld'), null);
	coordinator.setEnabled(true);
	const second = await coordinator.requestLint('Hello wrld');
	assert.equal(second.length, 1);
	assert.equal(created, 1);
	await coordinator.dispose();
	assert.equal(await coordinator.requestLint('Hello wrld'), null);
});

test('settings tab registers the grammar toggle and applies it through the plugin', async () => {
	const source = await readFile(path.join(projectRoot, 'src/settings.ts'), 'utf8');
	assert.match(source, /control: \{ type: 'toggle', key: 'enableGrammarChecking' \}/);
	assert.match(source, /case 'enableGrammarChecking'/);
	assert.match(source, /this\.plugin\.applyGrammarChecking\(\)/);
	const main = await readFile(path.join(projectRoot, 'src/main.ts'), 'utf8');
	assert.match(main, /this\.setupGrammarChecking\(\)/);
	assert.match(main, /this\.grammarChecking\?\.dispose\(\)/);
});
