import fs from 'node:fs';
import path from 'node:path';
import {
	LOCALES_DIR,
	EN_LOCALE_DIR,
	PSEUDO_LOCALES,
	extractPlaceholders,
	readLocaleCatalog,
	toMessageKey,
} from './lib/i18n-utils.mjs';

const errors = [];
const warnings = [];

function addError(message) {
	errors.push(message);
}

function addWarning(message) {
	warnings.push(message);
}

function listLocaleDirectories() {
	if (!fs.existsSync(LOCALES_DIR)) {
		return [];
	}

	return fs
		.readdirSync(LOCALES_DIR, { withFileTypes: true })
		.filter((entry) => entry.isDirectory())
		.map((entry) => entry.name)
		.filter((name) => !PSEUDO_LOCALES.has(name))
		.sort((left, right) => left.localeCompare(right));
}

function validateEnglishCatalog(enCatalog) {
	if (enCatalog.size === 0) {
		addError(`English catalog is empty or missing at ${path.relative(process.cwd(), EN_LOCALE_DIR)}.`);
		return;
	}

	for (const namespaceFile of enCatalog.values()) {
		for (const duplicateKey of namespaceFile.duplicateKeys) {
			addError(`${path.relative(process.cwd(), namespaceFile.filePath)}: duplicate key "${duplicateKey}" in English catalog.`);
		}
	}
}

function validateLocale(localeCode, enCatalog) {
	const localeDir = path.join(LOCALES_DIR, localeCode);
	const localeCatalog = readLocaleCatalog(localeDir);
	const isEnglish = localeCode === 'en';

	for (const [namespace, enNamespace] of enCatalog) {
		const localeNamespace = localeCatalog.get(namespace);

		if (!localeNamespace) {
			const message = `${localeCode}: missing namespace file "${namespace}.json".`;
			if (isEnglish) {
				addError(message);
			} else {
				addWarning(message);
			}
			continue;
		}

		for (const duplicateKey of localeNamespace.duplicateKeys) {
			const message = `${path.relative(process.cwd(), localeNamespace.filePath)}: duplicate key "${duplicateKey}".`;
			if (isEnglish) {
				addError(message);
			} else {
				addWarning(message);
			}
		}

		for (const [keyPath, englishValue] of enNamespace.flat) {
			const messageKey = toMessageKey(namespace, keyPath);
			if (!localeNamespace.flat.has(keyPath)) {
				const message = `${localeCode}: missing key "${messageKey}".`;
				if (isEnglish) {
					addError(message);
				} else {
					addWarning(message);
				}
				continue;
			}

			const localeValue = localeNamespace.flat.get(keyPath);
			const englishPlaceholders = extractPlaceholders(englishValue);
			const localePlaceholders = extractPlaceholders(localeValue);

			for (const placeholder of englishPlaceholders) {
				if (!localePlaceholders.has(placeholder)) {
					const message = `${localeCode}: placeholder "{${placeholder}}" missing in "${messageKey}".`;
					if (isEnglish) {
						addError(message);
					} else {
						addWarning(message);
					}
				}
			}

			for (const placeholder of localePlaceholders) {
				if (!englishPlaceholders.has(placeholder)) {
					const message = `${localeCode}: unknown placeholder "{${placeholder}}" in "${messageKey}".`;
					if (isEnglish) {
						addError(message);
					} else {
						addWarning(message);
					}
				}
			}
		}
	}

	for (const [namespace, localeNamespace] of localeCatalog) {
		const enNamespace = enCatalog.get(namespace);
		if (!enNamespace) {
			addError(`${localeCode}: unknown namespace file "${namespace}.json" (not present in English catalog).`);
			continue;
		}

		for (const keyPath of localeNamespace.flat.keys()) {
			if (!enNamespace.flat.has(keyPath)) {
				const messageKey = toMessageKey(namespace, keyPath);
				const message = `${localeCode}: unknown key "${messageKey}" (not present in English catalog).`;
				if (isEnglish) {
					addError(message);
				} else {
					addWarning(message);
				}
			}
		}
	}
}

function main() {
	const enCatalog = readLocaleCatalog(EN_LOCALE_DIR);
	validateEnglishCatalog(enCatalog);

	const localeCodes = listLocaleDirectories();
	if (localeCodes.length === 0 && enCatalog.size > 0) {
		localeCodes.push('en');
	}

	for (const localeCode of localeCodes) {
		validateLocale(localeCode, enCatalog);
	}

	console.log('Locale validation summary');
	console.log(`  locales checked: ${localeCodes.length}`);
	console.log(`  english namespaces: ${enCatalog.size}`);
	console.log(`  warnings: ${warnings.length}`);
	console.log(`  errors: ${errors.length}`);

	for (const warning of warnings) {
		console.warn(`WARN ${warning}`);
	}

	for (const error of errors) {
		console.error(`ERROR ${error}`);
	}

	if (errors.length > 0) {
		process.exitCode = 1;
	}
}

main();
