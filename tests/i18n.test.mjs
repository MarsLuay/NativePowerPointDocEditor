import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function loadModule(entry) {
	const outdir = await mkdtemp(path.join(tmpdir(), 'npde-i18n-test-'));
	const outfile = path.join(outdir, 'module.cjs');
	await build({
		absWorkingDir: projectRoot,
		entryPoints: [entry],
		bundle: true,
		format: 'cjs',
		logLevel: 'silent',
		outfile,
		platform: 'node',
		target: 'node22',
		loader: { '.json': 'json' },
	});
	return import(pathToFileURL(outfile).href);
}

test('localeCandidates falls back through BCP-47 chain', async () => {
	const { localeCandidates } = await loadModule('src/i18n/localeResolver.ts');
	assert.deepEqual(localeCandidates('pt-BR'), ['pt-BR', 'pt', 'en']);
	assert.deepEqual(localeCandidates('en-US'), ['en-US', 'en']);
});

test('resolveAutomaticLocale matches exact, base, then English', async () => {
	const { resolveAutomaticLocale } = await loadModule('src/i18n/localeResolver.ts');
	const available = ['en', 'pl'];

	assert.equal(resolveAutomaticLocale('pl', available), 'pl');
	assert.equal(resolveAutomaticLocale('pl-PL', available), 'pl');
	assert.equal(resolveAutomaticLocale('de', available), 'en');
	assert.equal(resolveAutomaticLocale('en-US', available), 'en');
});

test('resolveAutomaticLocale ignores path-like and invalid locale entries', async () => {
	const { resolveAutomaticLocale } = await loadModule('src/i18n/localeResolver.ts');
	const available = [
		'.obsidian/plugins/native-powerpoint-doc-editor/locales/en',
		'.obsidian/plugins/native-powerpoint-doc-editor/locales/pl',
		'not a locale',
	];

	assert.equal(resolveAutomaticLocale('pl-PL', available), 'en');
});

test('getLocaleDirection marks RTL locales', async () => {
	const { getLocaleDirection } = await loadModule('src/i18n/localeResolver.ts');
	assert.equal(getLocaleDirection('he'), 'rtl');
	assert.equal(getLocaleDirection('en'), 'ltr');
});

test('mergeNamespaceMessages flattens nested JSON', async () => {
	const { mergeNamespaceMessages } = await loadModule('src/i18n/localeLoader.ts');
	const merged = mergeNamespaceMessages('docx', {
		find: { title: 'Find' },
	});
	assert.equal(merged['docx:find.title'], 'Find');
});

test('listInstalledLocales normalizes Obsidian adapter folder paths', async () => {
	const { listInstalledLocales } = await loadModule('src/i18n/localeLoader.ts');
	const adapter = {
		exists: async () => true,
		read: async () => '',
		list: async () => [
			{ name: '.obsidian/plugins/native-powerpoint-doc-editor/locales/en', type: 'folder' },
			{ name: '.obsidian/plugins/native-powerpoint-doc-editor/locales/pl', type: 'folder' },
			{ name: '.obsidian/plugins/native-powerpoint-doc-editor/locales/readme.txt', type: 'file' },
		],
	};

	assert.deepEqual(await listInstalledLocales(adapter, '.obsidian/plugins/native-powerpoint-doc-editor'), ['en', 'pl']);
});

test('formatMessage replaces placeholders and plurals', async () => {
	const { formatMessage } = await loadModule('src/i18n/messageFormat.ts');
	assert.equal(
		formatMessage('{current} of {total}', { current: 2, total: 5 }),
		'2 of 5',
	);

	const template = '{count, plural, =0 {No matches} one {# match} other {# matches}}';
	assert.equal(formatMessage(template, { count: 0 }), 'No matches');
	assert.equal(formatMessage(template, { count: 1 }), '1 match');
	assert.equal(formatMessage(template, { count: 4 }), '4 matches');
});
