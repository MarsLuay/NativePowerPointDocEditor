import { Notice } from 'obsidian';

export type TranslateValues = Record<string, string | number | boolean>;

export type TranslateFn = (key: string, values?: TranslateValues) => string;

export type TranslateNoticeFn = (
	key: string,
	values?: TranslateValues,
	duration?: number,
) => void;

export function createTranslateNotice(t: TranslateFn): TranslateNoticeFn {
	return (key, values, duration) => {
		new Notice(t(key, values), duration);
	};
}
