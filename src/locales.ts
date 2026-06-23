import type { Translations } from '@eigenpal/docx-editor-i18n';

export type DocxidianLanguage = 'en' | 'pl' | 'pt-BR' | 'tr' | 'he' | 'zh-CN';

export interface DocxidianLanguageOption {
	code: DocxidianLanguage;
	label: string;
}

export const DEFAULT_LANGUAGE: DocxidianLanguage = 'en';

export const DOCXIDIAN_LANGUAGE_OPTIONS: DocxidianLanguageOption[] = [
	{ code: 'en', label: 'English' },
	{ code: 'pl', label: 'Polski' },
	{ code: 'pt-BR', label: 'Portugues do Brasil' },
	{ code: 'tr', label: 'Turkce' },
	{ code: 'he', label: 'Hebrew' },
	{ code: 'zh-CN', label: 'Simplified Chinese' },
];

const SUPPORTED_LANGUAGES = new Set<string>(DOCXIDIAN_LANGUAGE_OPTIONS.map((option) => option.code));

const localeCache = new Map<DocxidianLanguage, Translations>();
const localePromises = new Map<DocxidianLanguage, Promise<Translations | undefined>>();

const localeLoaders: Record<DocxidianLanguage, () => Promise<Translations>> = {
	en: async () => (await import('@eigenpal/docx-editor-i18n/en')).default,
	pl: async () => (await import('@eigenpal/docx-editor-i18n/pl')).default,
	'pt-BR': async () => (await import('@eigenpal/docx-editor-i18n/pt-BR')).default,
	tr: async () => (await import('@eigenpal/docx-editor-i18n/tr')).default,
	he: async () => (await import('@eigenpal/docx-editor-i18n/he')).default,
	'zh-CN': async () => (await import('@eigenpal/docx-editor-i18n/zh-CN')).default,
};

export function isDocxidianLanguage(value: string): value is DocxidianLanguage {
	return SUPPORTED_LANGUAGES.has(value);
}

export function normalizeDocxidianLanguage(value: unknown): DocxidianLanguage {
	return typeof value === 'string' && isDocxidianLanguage(value) ? value : DEFAULT_LANGUAGE;
}

export function getDocxEditorLocale(language: DocxidianLanguage): Translations | undefined {
	return localeCache.get(language);
}

export function preloadDocxEditorLocale(language: DocxidianLanguage): void {
	void loadDocxEditorLocale(language);
}

export function loadDocxEditorLocale(language: DocxidianLanguage): Promise<Translations | undefined> {
	const cached = localeCache.get(language);
	if (cached) {
		return Promise.resolve(cached);
	}

	const pending = localePromises.get(language);
	if (pending) {
		return pending;
	}

	const promise = localeLoaders[language]()
		.then((translations) => {
			localeCache.set(language, translations);
			return translations;
		})
		.catch(() => undefined)
		.finally(() => {
			localePromises.delete(language);
		});

	localePromises.set(language, promise);
	return promise;
}
