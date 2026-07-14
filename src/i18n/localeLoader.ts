import { BUNDLED_LOCALE_JSON, BUNDLED_LOCALES } from './bundledLocaleRegistry';
import { loadEigenpalMessages, type LoadedLocale } from './eigenpalAdapter';
import { getLocaleDirection, localeCandidates } from './localeResolver';

export const LOCALE_NAMESPACES = [
	'common',
	'settings',
	'docx',
	'powerpoint',
	'errors',
	'accessibility',
] as const;

export type LocaleNamespace = (typeof LOCALE_NAMESPACES)[number];

export type PluginMessages = Record<string, string>;

export interface LocaleFileAdapter {
	exists(path: string): Promise<boolean>;
	read(path: string): Promise<string>;
	list(path: string): Promise<Array<{ name: string; type: 'file' | 'folder' }>>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function flattenMessages(
	value: Record<string, unknown>,
	prefix = '',
): Record<string, string> {
	const result: Record<string, string> = {};

	for (const [key, nestedValue] of Object.entries(value)) {
		const fullKey = prefix ? `${prefix}.${key}` : key;
		if (typeof nestedValue === 'string') {
			result[fullKey] = nestedValue;
			continue;
		}

		if (isRecord(nestedValue)) {
			Object.assign(result, flattenMessages(nestedValue, fullKey));
		}
	}

	return result;
}

function getLocaleEntryName(name: string): string {
	return name.replace(/[\\/]+$/, '').split(/[\\/]/).pop() ?? name;
}

export function mergeNamespaceMessages(
	namespace: string,
	json: Record<string, unknown>,
): PluginMessages {
	const flat = flattenMessages(json);
	const result: PluginMessages = {};

	for (const [key, value] of Object.entries(flat)) {
		result[`${namespace}:${key}`] = value;
	}

	return result;
}

export function loadBundledPluginMessages(locale: string): PluginMessages {
	const localeJson = BUNDLED_LOCALE_JSON[locale as keyof typeof BUNDLED_LOCALE_JSON];
	if (!localeJson) {
		return {};
	}

	return LOCALE_NAMESPACES.reduce<PluginMessages>((messages, namespace) => {
		Object.assign(messages, mergeNamespaceMessages(namespace, localeJson[namespace]));
		return messages;
	}, {});
}

async function loadLocaleNamespaceMessages(
	adapter: LocaleFileAdapter,
	pluginDir: string,
	locale: string,
	namespace: LocaleNamespace,
): Promise<PluginMessages> {
	const filePath = `${pluginDir}/locales/${locale}/${namespace}.json`;
	if (!(await adapter.exists(filePath))) {
		return {};
	}

	try {
		const raw = await adapter.read(filePath);
		const parsed = JSON.parse(raw) as unknown;
		if (!isRecord(parsed)) {
			return {};
		}

		return mergeNamespaceMessages(namespace, parsed);
	} catch {
		return {};
	}
}

async function loadLocaleMessagesFromAdapter(
	adapter: LocaleFileAdapter,
	pluginDir: string,
	locale: string,
): Promise<PluginMessages> {
	const messages: PluginMessages = {};

	for (const namespace of LOCALE_NAMESPACES) {
		Object.assign(messages, await loadLocaleNamespaceMessages(adapter, pluginDir, locale, namespace));
	}

	return messages;
}

export async function loadPluginMessagesFromAdapter(
	adapter: LocaleFileAdapter,
	pluginDir: string,
	locale: string,
): Promise<PluginMessages> {
	const candidates = [...localeCandidates(locale)].reverse();
	let merged: PluginMessages = {};

	for (const candidate of candidates) {
		const candidateMessages = await loadLocaleMessagesFromAdapter(adapter, pluginDir, candidate);
		merged = { ...merged, ...candidateMessages };
	}

	return merged;
}

export function loadBundledPluginMessagesWithFallback(locale: string): PluginMessages {
	let merged: PluginMessages = {};

	for (const candidate of [...localeCandidates(locale)].reverse()) {
		merged = { ...merged, ...loadBundledPluginMessages(candidate) };
	}

	return merged;
}

export async function loadLocale(
	adapter: LocaleFileAdapter,
	pluginDir: string,
	requestedLocale: string,
): Promise<LoadedLocale> {
	const canonical = localeCandidates(requestedLocale)[0] ?? 'en';
	const [adapterMessages, eigenpalMessages] = await Promise.all([
		loadPluginMessagesFromAdapter(adapter, pluginDir, canonical),
		loadEigenpalMessages(canonical),
	]);

	return {
		locale: canonical,
		direction: getLocaleDirection(canonical),
		pluginMessages: {
			...loadBundledPluginMessagesWithFallback(canonical),
			...adapterMessages,
		},
		eigenpalMessages,
	};
}

export async function listInstalledLocales(
	adapter: LocaleFileAdapter,
	pluginDir: string,
): Promise<string[]> {
	const localesDir = `${pluginDir}/locales`;
	const locales = [...BUNDLED_LOCALES];
	if (await adapter.exists(localesDir)) {
		const entries = await adapter.list(localesDir);
		locales.push(...entries
			.filter((entry) => entry.type === 'folder')
			.map((entry) => getLocaleEntryName(entry.name))
			.filter((name) => name.length > 0 && name !== '.' && name !== '..'));
	}

	return [...new Set(locales)];
}
