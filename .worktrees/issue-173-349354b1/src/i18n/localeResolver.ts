const RTL_LANGUAGES = new Set(['he', 'ar', 'fa', 'ur']);

function getBaseLanguage(locale: string): string | undefined {
	const separator = locale.indexOf('-');
	if (separator === -1) {
		return undefined;
	}

	const base = locale.slice(0, separator);
	return base.length > 0 ? base : undefined;
}

function canonicalizeLocale(locale: string): string | null {
	try {
		return Intl.getCanonicalLocales(locale)[0] ?? null;
	} catch {
		return null;
	}
}

export function localeCandidates(requested: string): string[] {
	const canonical = canonicalizeLocale(requested);
	if (!canonical) {
		return ['en'];
	}

	const candidates: string[] = [canonical];

	try {
		const intlLocale = new Intl.Locale(canonical);
		if (intlLocale.language && intlLocale.language !== canonical && !candidates.includes(intlLocale.language)) {
			candidates.push(intlLocale.language);
		}
	} catch {
		const base = getBaseLanguage(canonical);
		if (base && !candidates.includes(base)) {
			candidates.push(base);
		}
	}

	const fallbackBase = getBaseLanguage(canonical);
	if (fallbackBase && !candidates.includes(fallbackBase)) {
		candidates.push(fallbackBase);
	}

	if (!candidates.includes('en')) {
		candidates.push('en');
	}

	return candidates;
}

export function getLocaleDirection(locale: string): 'ltr' | 'rtl' {
	const canonical = canonicalizeLocale(locale) ?? locale;

	try {
		const intlLocale = new Intl.Locale(canonical);
		if (RTL_LANGUAGES.has(intlLocale.language)) {
			return 'rtl';
		}
	} catch {
		const base = getBaseLanguage(canonical) ?? canonical;
		if (RTL_LANGUAGES.has(base)) {
			return 'rtl';
		}
	}

	return 'ltr';
}

/**
 * Pick the best installed plugin locale for an Obsidian locale:
 * exact canonical match → base-language match → English.
 */
export function resolveAutomaticLocale(
	obsidianLocale: string,
	availableLocales: readonly string[],
): string {
	const automaticLocales = availableLocales.filter(
		(locale) => !locale.endsWith('-XA') && !locale.endsWith('-XB'),
	);
	const normalizedAvailable = automaticLocales
		.map(canonicalizeLocale)
		.filter((locale): locale is string => Boolean(locale));
	const availableSet = new Set(normalizedAvailable);

	if (availableSet.size === 0) {
		return 'en';
	}

	const candidates = localeCandidates(obsidianLocale);

	for (const candidate of candidates) {
		const canonical = canonicalizeLocale(candidate);
		if (canonical && availableSet.has(canonical)) {
			return canonical;
		}
	}

	for (const candidate of candidates) {
		const canonical = canonicalizeLocale(candidate);
		const base = canonical ? getBaseLanguage(canonical) : undefined;
		if (!base) {
			continue;
		}

		for (const locale of normalizedAvailable) {
			if (locale === base || locale.startsWith(`${base}-`)) {
				return locale;
			}
		}
	}

	if (availableSet.has('en')) {
		return 'en';
	}

	return normalizedAvailable[0] ?? 'en';
}
