type MessageValues = Record<string, string | number | boolean>;

const SIMPLE_PLACEHOLDER_PATTERN = /\{(\w+)\}/g;

interface BracedContent {
	content: string;
	end: number;
}

interface PluralBranch {
	category: string;
	text: string;
}

function readBracedContent(template: string, start: number): BracedContent | null {
	if (template[start] !== '{') {
		return null;
	}

	let depth = 0;
	for (let index = start; index < template.length; index += 1) {
		const character = template[index];
		if (character === '{') {
			depth += 1;
		} else if (character === '}') {
			depth -= 1;
			if (depth === 0) {
				return {
					content: template.slice(start + 1, index),
					end: index + 1,
				};
			}
		}
	}

	return null;
}

function parsePluralBranches(body: string): PluralBranch[] {
	const branches: PluralBranch[] = [];
	let index = 0;

	while (index < body.length) {
		while (index < body.length && /\s/.test(body[index] ?? '')) {
			index += 1;
		}

		if (index >= body.length) {
			break;
		}

		let category = '';
		if (body[index] === '=') {
			index += 1;
			while (index < body.length && /\d/.test(body[index] ?? '')) {
				category += body[index] ?? '';
				index += 1;
			}
			category = `=${category}`;
		} else {
			while (index < body.length && /[a-zA-Z]/.test(body[index] ?? '')) {
				category += body[index] ?? '';
				index += 1;
			}
		}

		while (index < body.length && /\s/.test(body[index] ?? '')) {
			index += 1;
		}

		const branchText = readBracedContent(body, index);
		if (!category || !branchText) {
			break;
		}

		branches.push({ category, text: branchText.content });
		index = branchText.end;
	}

	return branches;
}

function tryParsePlural(
	template: string,
	start: number,
	values: MessageValues,
	locale: string,
): { text: string; end: number } | null {
	const headerMatch = template.slice(start).match(/^\{(\w+),\s*plural,\s*/);
	if (!headerMatch) {
		return null;
	}

	const variableName = headerMatch[1];
	if (!variableName) {
		return null;
	}

	const bodyStart = start + headerMatch[0].length;
	let depth = 1;
	let closingIndex = -1;
	for (let index = bodyStart; index < template.length; index += 1) {
		const character = template[index];
		if (character === '{') {
			depth += 1;
		} else if (character === '}') {
			depth -= 1;
			if (depth === 0) {
				closingIndex = index;
				break;
			}
		}
	}
	if (closingIndex === -1) {
		return null;
	}

	const branches = parsePluralBranches(template.slice(bodyStart, closingIndex));
	if (branches.length === 0) {
		return null;
	}

	const rawCount = values[variableName];
	const count = typeof rawCount === 'number'
		? rawCount
		: typeof rawCount === 'string'
			? Number(rawCount)
			: Number(rawCount);

	const pluralRules = new Intl.PluralRules(locale);
	const pluralCategory = Number.isFinite(count) ? pluralRules.select(count) : 'other';
	const exactCategory = `=${Number.isFinite(count) ? count : 0}`;

	const selectedBranch =
		branches.find((branch) => branch.category === exactCategory)
		?? branches.find((branch) => branch.category === pluralCategory)
		?? branches.find((branch) => branch.category === 'other')
		?? branches[branches.length - 1];

	if (!selectedBranch) {
		return null;
	}

	const branchValues: MessageValues = {
		...values,
		[variableName]: Number.isFinite(count) ? count : 0,
	};

	const branchText = selectedBranch.text.replace(/#/g, String(Number.isFinite(count) ? count : 0));
	return {
		text: formatMessage(branchText, branchValues, locale),
		end: closingIndex + 1,
	};
}

function formatPlurals(template: string, values: MessageValues, locale: string): string {
	let result = '';
	let index = 0;

	while (index < template.length) {
		const open = template.indexOf('{', index);
		if (open === -1) {
			result += template.slice(index);
			break;
		}

		result += template.slice(index, open);
		const parsedPlural = tryParsePlural(template, open, values, locale);
		if (parsedPlural) {
			result += parsedPlural.text;
			index = parsedPlural.end;
			continue;
		}

		result += '{';
		index = open + 1;
	}

	return result;
}

export function formatMessage(
	template: string,
	values?: MessageValues,
	locale = 'en',
): string {
	if (!values) {
		return template;
	}

	const withPlurals = formatPlurals(template, values, locale);
	return withPlurals.replace(SIMPLE_PLACEHOLDER_PATTERN, (_match, name: string) => {
		const value = values[name];
		return value === undefined ? `{${name}}` : String(value);
	});
}
