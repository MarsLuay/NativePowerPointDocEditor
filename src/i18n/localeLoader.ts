import accessibilityEn from '../../locales/en/accessibility.json';
import commonEn from '../../locales/en/common.json';
import docxEn from '../../locales/en/docx.json';
import errorsEn from '../../locales/en/errors.json';
import powerpointEn from '../../locales/en/powerpoint.json';
import settingsEn from '../../locales/en/settings.json';
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

const BUNDLED_EN_LOCALE_JSON: Record<LocaleNamespace, Record<string, unknown>> = {
	common: commonEn,
	settings: settingsEn,
	docx: docxEn,
	powerpoint: powerpointEn,
	errors: errorsEn,
	accessibility: accessibilityEn,
};

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
	if (locale !== 'en') {
		return {};
	}

	return LOCALE_NAMESPACES.reduce<PluginMessages>((messages, namespace) => {
		Object.assign(messages, mergeNamespaceMessages(namespace, BUNDLED_EN_LOCALE_JSON[namespace]));
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

export async function loadLocale(
	adapter: LocaleFileAdapter,
	pluginDir: string,
	requestedLocale: string,
): Promise<LoadedLocale> {
	const canonical = localeCandidates(requestedLocale)[0] ?? 'en';
	const [pluginMessages, eigenpalMessages] = await Promise.all([
		loadPluginMessagesFromAdapter(adapter, pluginDir, canonical),
		loadEigenpalMessages(canonical),
	]);

	return {
		locale: canonical,
		direction: getLocaleDirection(canonical),
		pluginMessages,
		eigenpalMessages,
	};
}

export async function listInstalledLocales(
	adapter: LocaleFileAdapter,
	pluginDir: string,
): Promise<string[]> {
	const localesDir = `${pluginDir}/locales`;
	if (!(await adapter.exists(localesDir))) {
		return ['en'];
	}

	const entries = await adapter.list(localesDir);
	const locales = entries
		.filter((entry) => entry.type === 'folder')
		.map((entry) => getLocaleEntryName(entry.name))
		.filter((name) => name.length > 0 && name !== '.' && name !== '..');

	return locales.length > 0 ? [...new Set(locales)] : ['en'];
}
