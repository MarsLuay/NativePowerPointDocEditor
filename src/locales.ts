import type { Translations } from '@eigenpal/docx-editor-i18n';

import { loadEigenpalMessages } from './i18n/eigenpalAdapter';
import { localeCandidates } from './i18n/localeResolver';

const SUPPORTED_DOCX_EDITOR_LANGUAGES = ['en', 'pl', 'pt-BR', 'tr', 'he', 'zh-CN'] as const;

export type NativePowerPointDocEditorLanguage = (typeof SUPPORTED_DOCX_EDITOR_LANGUAGES)[number];

export type LocaleCode = string;

export const DEFAULT_LANGUAGE: NativePowerPointDocEditorLanguage = 'en';

const SUPPORTED_LANGUAGES = new Set<string>(SUPPORTED_DOCX_EDITOR_LANGUAGES);

const localeCache = new Map<string, Translations>();
const localePromises = new Map<string, Promise<Translations | undefined>>();

export function isNativePowerPointDocEditorLanguage(value: string): value is NativePowerPointDocEditorLanguage {
	return SUPPORTED_LANGUAGES.has(value);
}

export function resolveAutomaticDocxEditorLanguage(locale: string): NativePowerPointDocEditorLanguage {
	for (const candidate of localeCandidates(locale)) {
		if (isNativePowerPointDocEditorLanguage(candidate)) {
			return candidate;
		}
	}

	return DEFAULT_LANGUAGE;
}

export function getDocxEditorLocale(language: NativePowerPointDocEditorLanguage): Translations | undefined {
	return localeCache.get(language);
}

export function preloadDocxEditorLocale(language: NativePowerPointDocEditorLanguage): void {
	void loadDocxEditorLocale(language);
}

export function loadDocxEditorLocale(language: LocaleCode): Promise<Translations | undefined> {
	const cached = localeCache.get(language);
	if (cached) {
		return Promise.resolve(cached);
	}

	const pending = localePromises.get(language);
	if (pending) {
		return pending;
	}

	const promise = loadEigenpalMessages(language)
		.then((translations) => {
			if (translations) {
				localeCache.set(language, translations);
			}
			return translations;
		})
		.catch(() => undefined)
		.finally(() => {
			localePromises.delete(language);
		});

	localePromises.set(language, promise);
	return promise;
}
