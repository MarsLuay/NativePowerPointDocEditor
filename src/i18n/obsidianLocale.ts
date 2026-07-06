import { getLanguage } from 'obsidian';

/** ISO code for Obsidian's configured app language. */
export function getObsidianLocale(): string {
	try {
		return getLanguage() || 'en';
	} catch {
		return 'en';
	}
}
