import fs from 'node:fs';
import path from 'node:path';
import {
	EN_LOCALE_DIR,
	LOCALES_DIR,
	ensureDirectory,
	listLocaleJsonFiles,
	pseudoLocalizeArXB,
	pseudoLocalizeEnXA,
	transformLeafValues,
} from './lib/i18n-utils.mjs';

function writePseudoLocale(targetLocale, transform) {
	const targetDir = path.join(LOCALES_DIR, targetLocale);
	const sourceFiles = listLocaleJsonFiles(EN_LOCALE_DIR);

	if (sourceFiles.length === 0) {
		console.log(`Skipped ${targetLocale}: no English locale files found.`);
		return 0;
	}

	ensureDirectory(targetDir);

	for (const sourcePath of sourceFiles) {
		const parsed = JSON.parse(fs.readFileSync(sourcePath, 'utf8'));
		const transformed = transformLeafValues(parsed, transform);
		const targetPath = path.join(targetDir, path.basename(sourcePath));
		fs.writeFileSync(targetPath, `${JSON.stringify(transformed, null, 2)}\n`, 'utf8');
	}

	console.log(`Wrote ${sourceFiles.length} pseudo-locale file(s) to locales/${targetLocale}/.`);
	return sourceFiles.length;
}

function main() {
	const enXAFiles = writePseudoLocale('en-XA', pseudoLocalizeEnXA);
	const arXBFiles = writePseudoLocale('ar-XB', pseudoLocalizeArXB);
	console.log(`Generated pseudo locales (${enXAFiles + arXBFiles} files total).`);
}

main();
