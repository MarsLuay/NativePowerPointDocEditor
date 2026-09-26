import { loadDocxEditorLocale } from '../docx/runtime/bridge.mjs';
import type { Translations } from '../docx/runtime/contract';
import type { PluginMessages } from './localeLoader';
import { localeCandidates } from './localeResolver';

export interface LoadedLocale {
	locale: string;
	direction: 'ltr' | 'rtl';
	pluginMessages: PluginMessages;
	docxEditorMessages: Translations | undefined;
}

export async function loadDocxEditorMessages(locale: string): Promise<Translations | undefined> {
	for (const candidate of localeCandidates(locale)) {
		try {
			const messages = await loadDocxEditorLocale(candidate);
			if (messages) return messages;
		} catch {
			continue;
		}
	}

	return undefined;
}
