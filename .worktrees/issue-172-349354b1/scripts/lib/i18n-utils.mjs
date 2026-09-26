import fs from 'node:fs';
import path from 'node:path';

export const PROJECT_ROOT = path.resolve(import.meta.dirname, '../..');
export const LOCALES_DIR = path.join(PROJECT_ROOT, 'locales');
export const EN_LOCALE_DIR = path.join(LOCALES_DIR, 'en');

export const PSEUDO_LOCALES = new Set(['en-XA', 'ar-XB']);

const PLACEHOLDER_PATTERN = /\{([a-zA-Z_][a-zA-Z0-9_]*)(?:\s*,\s*(?:plural|select|selectordinal)[^}]*)?\}/g;

export function listLocaleJsonFiles(localeDir) {
	if (!fs.existsSync(localeDir)) {
		return [];
	}
	return fs
		.readdirSync(localeDir, { withFileTypes: true })
		.filter((entry) => entry.isFile() && entry.name.endsWith('.json'))
		.map((entry) => path.join(localeDir, entry.name))
		.sort((left, right) => left.localeCompare(right));
}

export function namespaceFromFile(filePath) {
	return path.basename(filePath, '.json');
}

export function flattenJsonObject(value, prefix = '') {
	const entries = new Map();
	if (value === null || typeof value !== 'object' || Array.isArray(value)) {
		if (prefix) {
			entries.set(prefix, value);
		}
		return entries;
	}

	for (const [key, nestedValue] of Object.entries(value)) {
		const nextPrefix = prefix ? `${prefix}.${key}` : key;
		if (nestedValue !== null && typeof nestedValue === 'object' && !Array.isArray(nestedValue)) {
			for (const [nestedKey, nestedEntryValue] of flattenJsonObject(nestedValue, nextPrefix)) {
				entries.set(nestedKey, nestedEntryValue);
			}
		} else {
			entries.set(nextPrefix, nestedValue);
		}
	}

	return entries;
}

export function readNamespaceFile(filePath) {
	const raw = fs.readFileSync(filePath, 'utf8');
	const namespace = namespaceFromFile(filePath);
	let parsed;

	try {
		parsed = JSON.parse(raw);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		throw new Error(`${filePath}: invalid JSON (${message})`);
	}

	return {
		namespace,
		filePath,
		raw,
		flat: flattenJsonObject(parsed),
		duplicateKeys: findDuplicateKeysInJson(raw),
	};
}

export function readLocaleCatalog(localeDir) {
	const catalog = new Map();

	for (const filePath of listLocaleJsonFiles(localeDir)) {
		const namespaceFile = readNamespaceFile(filePath);
		catalog.set(namespaceFile.namespace, namespaceFile);
	}

	return catalog;
}

export function toMessageKey(namespace, keyPath) {
	return `${namespace}:${keyPath}`;
}

export function collectMessageKeys(catalog) {
	const keys = [];

	for (const namespaceFile of catalog.values()) {
		for (const keyPath of namespaceFile.flat.keys()) {
			keys.push(toMessageKey(namespaceFile.namespace, keyPath));
		}
	}

	return keys.sort((left, right) => left.localeCompare(right));
}

export function extractPlaceholders(text) {
	if (typeof text !== 'string') {
		return new Set();
	}

	const placeholders = new Set();
	for (const match of text.matchAll(PLACEHOLDER_PATTERN)) {
		placeholders.add(match[1]);
	}

	return placeholders;
}

export function findDuplicateKeysInJson(rawText) {
	const duplicates = [];
	const stack = [new Set()];
	let index = 0;

	while (index < rawText.length) {
		const char = rawText[index];

		if (char === '{') {
			stack.push(new Set());
			index += 1;
			continue;
		}

		if (char === '}') {
			if (stack.length > 1) {
				stack.pop();
			}
			index += 1;
			continue;
		}

		if (char === '"') {
			index += 1;
			let key = '';

			while (index < rawText.length) {
				const current = rawText[index];
				if (current === '\\') {
					index += 2;
					continue;
				}
				if (current === '"') {
					break;
				}
				key += current;
				index += 1;
			}

			index += 1;
			while (index < rawText.length && /\s/.test(rawText[index])) {
				index += 1;
			}

			if (rawText[index] === ':') {
				const currentLevel = stack[stack.length - 1];
				if (currentLevel.has(key)) {
					duplicates.push(key);
				}
				currentLevel.add(key);
			}

			continue;
		}

		index += 1;
	}

	return duplicates;
}

const EN_XA_MAP = {
	a: 'ą',
	A: 'Ą',
	c: 'ć',
	C: 'Ć',
	d: 'ḑ',
	D: 'Ḑ',
	e: 'ę',
	E: 'Ę',
	g: 'ğ',
	G: 'Ğ',
	i: 'į',
	I: 'Į',
	l: 'ſ',
	L: 'Ł',
	n: 'ñ',
	N: 'Ñ',
	o: 'ǫ',
	O: 'Ǫ',
	s: 'ş',
	S: 'Ş',
	t: 'ţ',
	T: 'Ţ',
	u: 'ų',
	U: 'Ų',
	y: 'ÿ',
	Y: 'Ÿ',
};

function accentizeLiteral(text) {
	return [...text]
		.map((char) => EN_XA_MAP[char] ?? char)
		.join('');
}

function tokenizePreservingBracedSegments(text) {
	const parts = [];
	let index = 0;

	while (index < text.length) {
		if (text[index] === '{') {
			let depth = 0;
			let end = index;

			for (; end < text.length; end += 1) {
				const char = text[end];
				if (char === '{') {
					depth += 1;
				} else if (char === '}') {
					depth -= 1;
					if (depth === 0) {
						end += 1;
						break;
					}
				}
			}

			parts.push({ type: 'preserved', value: text.slice(index, end) });
			index = end;
			continue;
		}

		const nextBrace = text.indexOf('{', index);
		const end = nextBrace === -1 ? text.length : nextBrace;
		parts.push({ type: 'literal', value: text.slice(index, end) });
		index = end;
	}

	return parts.length > 0 ? parts : [{ type: 'literal', value: text }];
}

function transformPreservingBracedSegments(text, transformLiteral) {
	return tokenizePreservingBracedSegments(text)
		.map((part) => (part.type === 'literal' ? transformLiteral(part.value) : part.value))
		.join('');
}

export function pseudoLocalizeEnXA(text) {
	if (typeof text !== 'string') {
		return text;
	}

	return `[${transformPreservingBracedSegments(text, accentizeLiteral)}]`;
}

export function pseudoLocalizeArXB(text) {
	if (typeof text !== 'string') {
		return text;
	}

	return `\u202B${transformPreservingBracedSegments(text, (literal) => [...literal].reverse().join(''))}`;
}

export function transformLeafValues(value, transform) {
	if (typeof value === 'string') {
		return transform(value);
	}

	if (value === null || typeof value !== 'object' || Array.isArray(value)) {
		return value;
	}

	const next = {};
	for (const [key, nestedValue] of Object.entries(value)) {
		next[key] = transformLeafValues(nestedValue, transform);
	}
	return next;
}

export function ensureDirectory(dirPath) {
	fs.mkdirSync(dirPath, { recursive: true });
}
