import type { DataAdapter, Plugin } from 'obsidian';

import { createPluginI18nService, PluginI18nService } from './I18nService';
import {
	listInstalledLocales,
	loadBundledPluginMessages,
	loadLocale,
	type LocaleFileAdapter,
} from './localeLoader';
import { getObsidianLocale } from './obsidianLocale';
import { resolveAutomaticLocale } from './localeResolver';

type PluginWithI18n = Plugin & { i18n: PluginI18nService | null };

let activePlugin: PluginWithI18n | null = null;

function createVaultLocaleFileAdapter(dataAdapter: DataAdapter): LocaleFileAdapter {
	return {
		exists: (path) => dataAdapter.exists(path),
		read: (path) => dataAdapter.read(path),
		list: async (path) => {
			const listed = await dataAdapter.list(path);
			return [
				...listed.files.map((name) => ({ name, type: 'file' as const })),
				...listed.folders.map((name) => ({ name, type: 'folder' as const })),
			];
		},
	};
}

export async function resolvePluginLocale(plugin: Plugin): Promise<string> {
	const pluginDir = plugin.manifest.dir;
	if (!pluginDir) {
		return resolveAutomaticLocale(getObsidianLocale(), ['en']);
	}

	const adapter = createVaultLocaleFileAdapter(plugin.app.vault.adapter);
	const availableLocales = await listInstalledLocales(adapter, pluginDir);
	return resolveAutomaticLocale(getObsidianLocale(), availableLocales);
}

export async function initPluginI18n(
	plugin: PluginWithI18n,
	requestedLocale: string,
): Promise<PluginI18nService> {
	const pluginDir = plugin.manifest.dir;
	if (!pluginDir) {
		throw new Error('Plugin manifest dir is required to initialize i18n');
	}

	const adapter = createVaultLocaleFileAdapter(plugin.app.vault.adapter);
	const loaded = await loadLocale(adapter, pluginDir, requestedLocale);
	const englishFallback = loadBundledPluginMessages('en');
	const service = createPluginI18nService(loaded, englishFallback);

	plugin.i18n = service;
	activePlugin = plugin;
	return service;
}

export function getPluginI18n(): PluginI18nService | null {
	return activePlugin?.i18n ?? null;
}
