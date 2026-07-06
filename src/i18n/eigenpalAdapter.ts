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
	en: async () => (await import('@eigenpal/docx-editor-i18n/en')).default,
	pl: async () => (await import('@eigenpal/docx-editor-i18n/pl')).default,
	'pt-BR': async () => (await import('@eigenpal/docx-editor-i18n/pt-BR')).default,
	tr: async () => (await import('@eigenpal/docx-editor-i18n/tr')).default,
	he: async () => (await import('@eigenpal/docx-editor-i18n/he')).default,
	'zh-CN': async () => (await import('@eigenpal/docx-editor-i18n/zh-CN')).default,
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
