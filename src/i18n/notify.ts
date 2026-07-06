import { Notice } from 'obsidian';

import type { I18nService, MessageKey } from './I18nService';

type NoticeValues = Record<string, string | number | boolean>;

export function showI18nNotice(
	i18n: I18nService | null | undefined,
	key: MessageKey | string,
	values?: NoticeValues,
): void {
	new Notice(i18n?.t(key, values) ?? key);
}
