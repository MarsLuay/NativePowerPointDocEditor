import fs from 'node:fs';
import path from 'node:path';
import { PROJECT_ROOT } from './lib/i18n-utils.mjs';

const SOURCE_ROOT = path.join(PROJECT_ROOT, 'src');
const SKIP_DIR_NAMES = new Set(['node_modules', 'editor', 'docx-editor', 'build', 'dist']);

const SCAN_PATTERNS = [
	{ label: 'Notice', regex: /new\s+Notice\s*\(\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'aria-label', regex: /\.setAttribute\s*\(\s*['"]aria-label['"]\s*,\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'title', regex: /\.setAttribute\s*\(\s*['"]title['"]\s*,\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'placeholder', regex: /\.(?:setAttribute\s*\(\s*['"]placeholder['"]|placeholder)\s*(?:,\s*)?\s*(?:=\s*)?(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'data-tooltip', regex: /['"]data-tooltip['"]\s*:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'textContent', regex: /\.textContent\s*=\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'setText', regex: /\.setText\s*\(\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'setName', regex: /\.setName\s*\(\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'setDesc', regex: /\.setDesc\s*\(\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'label', regex: /\blabel:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'name', regex: /\bname:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'description', regex: /\bdescription:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'emptyState', regex: /\bemptyState:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
	{ label: 'text', regex: /\btext:\s*(['"`])((?:\\.|(?!\1)[\s\S])*?)\1/g },
];

/** Precise allowlist: file (posix) + exact literal value */
const PRECISE_ALLOWLIST = new Map([
	['src/DocxView.tsx', new Set(['›'])],
	['src/powerpoint/ui/NativePowerPointView.ts', new Set(['@page { size: landscape; margin: 12mm; }'])],
]);

function collectSourceFiles(root) {
	const files = [];

	function visit(current) {
		for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
			const fullPath = path.join(current, entry.name);
			if (entry.isDirectory()) {
				if (!SKIP_DIR_NAMES.has(entry.name)) {
					visit(fullPath);
				}
				continue;
			}

			if (entry.isFile() && ['.ts', '.tsx'].includes(path.extname(entry.name))) {
				files.push(fullPath);
			}
		}
	}

	visit(root);
	return files;
}

function isAllowlisted(filePath) {
	const relativePath = path.relative(PROJECT_ROOT, filePath).replaceAll('\\', '/');
	return (
		relativePath.startsWith('src/i18n/')
		|| relativePath.startsWith('locales/')
		|| relativePath.includes('.generated.')
		|| relativePath.endsWith('.generated.ts')
	);
}

function isPreciselyAllowlisted(relativePath, literalValue) {
	return PRECISE_ALLOWLIST.get(relativePath)?.has(literalValue) ?? false;
}

function unescapeString(value) {
	return value
		.replace(/\\n/g, '\n')
		.replace(/\\r/g, '\r')
		.replace(/\\t/g, '\t')
		.replace(/\\'/g, "'")
		.replace(/\\"/g, '"')
		.replace(/\\\\/g, '\\');
}

function shouldIgnore(line, literalValue) {
	if (/debugLog\b/.test(line)) {
		return true;
	}
	if (/classList\.add\b/.test(line)) {
		return true;
	}
	if (/throw\s+new\s+Error\b/.test(line)) {
		return true;
	}
	if (/translate(?:X|Y|3d)?\s*\(/.test(line)) {
		return true;
	}
	if (/\bpptT\s*\(/.test(line) || /\bthis\.t\s*\(/.test(line) || /\bhost\.t\s*\(/.test(line)) {
		return true;
	}
	if (/\bdocxT\s*\(/.test(line) || /\bi18n\.t\s*\(/.test(line)) {
		return true;
	}

	const trimmed = literalValue.trim();
	if (trimmed.length === 0) {
		return true;
	}
	if (/^[.#\[]/.test(trimmed)) {
		return true;
	}
	if (/^(\.\.?\/|\/|[a-zA-Z]:\\)/.test(trimmed)) {
		return true;
	}
	if (trimmed.includes('/') && !/\s/.test(trimmed) && !trimmed.includes('://')) {
		return true;
	}
	if (/^[\d\s%pxem-]+$/.test(trimmed)) {
		return true;
	}

	return false;
}

function lineForIndex(text, index) {
	return text.slice(0, index).split('\n').length;
}

function scanFile(filePath) {
	const source = fs.readFileSync(filePath, 'utf8');
	const relativePath = path.relative(PROJECT_ROOT, filePath).replaceAll('\\', '/');
	const findings = [];

	for (const pattern of SCAN_PATTERNS) {
		pattern.regex.lastIndex = 0;
		for (const match of source.matchAll(pattern.regex)) {
			const literalValue = unescapeString(match[2] ?? '');
			const lineNumber = lineForIndex(source, match.index ?? 0);
			const line = source.split('\n')[lineNumber - 1] ?? '';

			if (shouldIgnore(line, literalValue)) {
				continue;
			}
			if (isPreciselyAllowlisted(relativePath, literalValue)) {
				continue;
			}

			findings.push({
				filePath,
				line: lineNumber,
				kind: pattern.label,
				text: literalValue,
			});
		}
	}

	return findings;
}

function main() {
	const files = collectSourceFiles(SOURCE_ROOT).filter((filePath) => !isAllowlisted(filePath));
	const findings = files.flatMap((filePath) => scanFile(filePath));

	findings.sort((left, right) => {
		const pathCompare = left.filePath.localeCompare(right.filePath);
		if (pathCompare !== 0) {
			return pathCompare;
		}
		return left.line - right.line;
	});

	for (const finding of findings) {
		const relativePath = path.relative(PROJECT_ROOT, finding.filePath).replaceAll('\\', '/');
		console.log(`${relativePath}:${finding.line}\t[${finding.kind}]\t${finding.text}`);
	}

	console.log(`Found ${findings.length} likely untranslated user-facing string literal(s).`);
	process.exit(findings.length > 0 ? 1 : 0);
}

main();
