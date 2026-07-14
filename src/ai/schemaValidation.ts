import type { JsonSchema } from './types';

export interface SchemaValidationIssue {
	path: string;
	message: string;
}

function typeName(value: unknown): string {
	if (value === null) return 'null';
	if (Array.isArray(value)) return 'array';
	return typeof value;
}

function matchesSchemaType(value: unknown, expectedType: NonNullable<JsonSchema['type']>): boolean {
	const actual = typeName(value);
	if (expectedType === 'integer') {
		return actual === 'number' && Number.isInteger(value);
	}
	return actual === expectedType;
}

function validateValue(
	value: unknown,
	schema: JsonSchema,
	path: string,
	issues: SchemaValidationIssue[],
): void {
	const expectedType = schema.type;
	if (expectedType && !matchesSchemaType(value, expectedType)) {
		issues.push({ path, message: `Expected ${expectedType}, got ${typeName(value)}.` });
		return;
	}

	if (expectedType === 'object' && value && typeof value === 'object' && !Array.isArray(value)) {
		const record = value as Record<string, unknown>;
		for (const key of schema.required ?? []) {
			if (!(key in record)) {
				issues.push({ path: path ? `${path}.${key}` : key, message: 'Required property is missing.' });
			}
		}

		for (const [key, childValue] of Object.entries(record)) {
			const childSchema = schema.properties?.[key];
			if (!childSchema) {
				if (schema.additionalProperties === false) {
					issues.push({ path: path ? `${path}.${key}` : key, message: 'Unknown property.' });
				}
				continue;
			}
			validateValue(childValue, childSchema, path ? `${path}.${key}` : key, issues);
		}
		return;
	}

	if (expectedType === 'array' && Array.isArray(value) && schema.items) {
		value.forEach((item, index) => {
			validateValue(item, schema.items as JsonSchema, `${path}[${index}]`, issues);
		});
		return;
	}

	if (schema.enum && !schema.enum.includes(value as never)) {
		issues.push({ path, message: `Value must be one of: ${schema.enum.join(', ')}.` });
	}

	if (typeof value === 'number' && schema.minimum !== undefined && value < schema.minimum) {
		issues.push({ path, message: `Value must be >= ${schema.minimum}.` });
	}
}

export function validateAgainstSchema(
	value: unknown,
	schema: JsonSchema,
	rootPath = '',
): SchemaValidationIssue[] {
	const issues: SchemaValidationIssue[] = [];
	validateValue(value, schema, rootPath, issues);
	return issues;
}
