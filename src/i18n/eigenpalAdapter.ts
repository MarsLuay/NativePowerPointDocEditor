import type { Translations } from '@eigenpal/docx-editor-i18n';

import type { PluginMessages } from './localeLoader';
import { localeCandidates } from './localeResolver';

export interface LoadedLocale {
	locale: string;
	direction: 'ltr' | 'rtl';
	pluginMessages: PluginMessages;
	eigenpalMessages: Translations | undefined;
}

const eigenpalLocaleLoaders: Record<string, () => Promise<Translations>> = {
	en: async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/en.mjs')).default,
	pl: async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/pl.mjs')).default,
	'pt-BR': async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/pt-BR.mjs')).default,
	tr: async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/tr.mjs')).default,
	he: async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/he.mjs')).default,
	'zh-CN': async () => (await import('../vendor/eigenpal/docx-editor-i18n/dist/zh-CN.mjs')).default,
};

export async function loadEigenpalMessages(locale: string): Promise<Translations | undefined> {
	for (const candidate of localeCandidates(locale)) {
		const loader = eigenpalLocaleLoaders[candidate];
		if (!loader) {
			continue;
		}

		try {
			return await loader();
		} catch {
			continue;
		}
	}

	return undefined;
}
