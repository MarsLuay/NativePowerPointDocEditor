export type DocxPartKind = 'body' | 'header' | 'footer' | 'footnotes' | 'endnotes';

export type DocxStableBlockKind = 'paragraph' | 'run' | 'table' | 'cell';

export interface DocxStableLocation {
	part: DocxPartKind;
	partNumber: number | null;
	kind: DocxStableBlockKind;
	paragraphIndex: number;
	runIndex: number | null;
	tableIndex: number | null;
	rowIndex: number | null;
	colIndex: number | null;
}

interface PartRule {
	pattern: RegExp;
	part: DocxPartKind;
	getPartNumber: (match: RegExpExecArray) => number | null;
}

interface SuffixRule {
	pattern: RegExp;
	kind: DocxStableBlockKind;
	getIndices: (match: RegExpExecArray) => {
		paragraphIndex: number;
		runIndex: number | null;
		tableIndex: number | null;
		rowIndex: number | null;
		colIndex: number | null;
	};
}

const PART_RULES: PartRule[] = [
	{
		pattern: /^body\/(.+)$/,
		part: 'body',
		getPartNumber: () => null,
	},
	{
		pattern: /^header\/(\d+)\/(.+)$/,
		part: 'header',
		getPartNumber: (m) => Number(m[1]),
	},
	{
		pattern: /^footer\/(\d+)\/(.+)$/,
		part: 'footer',
		getPartNumber: (m) => Number(m[1]),
	},
	{
		pattern: /^footnotes\/fn\[(\d+)\]\/(.+)$/,
		part: 'footnotes',
		getPartNumber: (m) => Number(m[1]),
	},
	{
		pattern: /^endnotes\/en\[(\d+)\]\/(.+)$/,
		part: 'endnotes',
		getPartNumber: (m) => Number(m[1]),
	},
];

const SUFFIX_RULES: SuffixRule[] = [
	{
		pattern: /^p\[(\d+)\]$/,
		kind: 'paragraph',
		getIndices: (m) => ({
			paragraphIndex: Number(m[1]),
			runIndex: null,
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		}),
	},
	{
		pattern: /^p\[(\d+)\]\/r\[(\d+)\]$/,
		kind: 'run',
		getIndices: (m) => ({
			paragraphIndex: Number(m[1]),
			runIndex: Number(m[2]),
			tableIndex: null,
			rowIndex: null,
			colIndex: null,
		}),
	},
	{
		pattern: /^tbl\[(\d+)\]$/,
		kind: 'table',
		getIndices: (m) => ({
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(m[1]),
			rowIndex: null,
			colIndex: null,
		}),
	},
	{
		pattern: /^tbl\[(\d+)\]\/tr\[(\d+)\]\/tc\[(\d+)\]$/,
		kind: 'cell',
		getIndices: (m) => ({
			paragraphIndex: 0,
			runIndex: null,
			tableIndex: Number(m[1]),
			rowIndex: Number(m[2]),
			colIndex: Number(m[3]),
		}),
	},
];

export function docxIdPrefix(part: DocxPartKind, partNumber: number | null): string {
	switch (part) {
		case 'body':
			return 'body';
		case 'header':
			return `header/${partNumber ?? 1}`;
		case 'footer':
			return `footer/${partNumber ?? 1}`;
		case 'footnotes':
			return `footnotes/fn[${partNumber ?? 0}]`;
		case 'endnotes':
			return `endnotes/en[${partNumber ?? 0}]`;
	}
}

export function parseStableLocation(id: string): DocxStableLocation | null {
	for (const partRule of PART_RULES) {
		const partMatch = partRule.pattern.exec(id);
		if (!partMatch) {
			continue;
		}

		const rest = partMatch[partMatch.length - 1] ?? '';
		for (const suffixRule of SUFFIX_RULES) {
			const suffixMatch = suffixRule.pattern.exec(rest);
			if (!suffixMatch) {
				continue;
			}

			return {
				part: partRule.part,
				partNumber: partRule.getPartNumber(partMatch),
				kind: suffixRule.kind,
				...suffixRule.getIndices(suffixMatch),
			};
		}
		return null;
	}

	return null;
}

export function paragraphIdForLocation(location: DocxStableLocation): string {
	const prefix = docxIdPrefix(location.part, location.partNumber);
	return `${prefix}/p[${location.paragraphIndex}]`;
}

export function runIdForLocation(location: DocxStableLocation): string {
	if (location.runIndex === null) {
		throw new Error('runIdForLocation requires runIndex.');
	}
	return `${paragraphIdForLocation(location)}/r[${location.runIndex}]`;
}

export function tableIdForLocation(location: DocxStableLocation): string {
	const prefix = docxIdPrefix(location.part, location.partNumber);
	return `${prefix}/tbl[${location.tableIndex ?? 0}]`;
}

export function cellIdForLocation(location: DocxStableLocation): string {
	return `${tableIdForLocation(location)}/tr[${location.rowIndex ?? 0}]/tc[${location.colIndex ?? 0}]`;
}
