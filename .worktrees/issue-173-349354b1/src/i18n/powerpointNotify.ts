import { Notice } from 'obsidian';

import { getPluginI18n } from './pluginI18n';
import type { TranslateFn, TranslateNoticeFn } from './translate';
import { createTranslateNotice } from './translate';

export type { TranslateFn, TranslateNoticeFn, TranslateValues } from './translate';

export const pptT: TranslateFn = (key, values) => getPluginI18n()?.t(key, values) ?? key;

export const pptNotice: TranslateNoticeFn = (key, values, duration) => {
	new Notice(pptT(key, values), duration);
};

export function createPowerPointTranslate(t: TranslateFn = pptT): {
	t: TranslateFn;
	notice: TranslateNoticeFn;
} {
	return {
		t,
		notice: createTranslateNotice(t),
	};
}
