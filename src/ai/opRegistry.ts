import { OP_CATALOG } from './opCatalog';
import { createAiError, AI_ERROR_CODES } from './errors';
import { validateAgainstSchema } from './schemaValidation';
import type { AiErrorDetail } from './errors';
import type { DocumentOp, OpDefinition } from './types';

const OP_BY_ID = new Map(OP_CATALOG.map((op) => [op.id, op]));

export function getOpDefinition(opId: string): OpDefinition | undefined {
	return OP_BY_ID.get(opId);
}

export function listOpDefinitions(): readonly OpDefinition[] {
	return OP_CATALOG;
}

export function validateDocumentOp(op: DocumentOp, index: number): AiErrorDetail[] {
	const errors: AiErrorDetail[] = [];
	if (!op || typeof op !== 'object' || Array.isArray(op)) {
		errors.push(createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Operation must be an object.', {
			field: `ops[${index}]`,
		}));
		return errors;
	}

	const opId = op.op;
	if (typeof opId !== 'string' || opId.trim().length === 0) {
		errors.push(createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'Operation id is required.', {
			field: `ops[${index}].op`,
		}));
		return errors;
	}

	const definition = getOpDefinition(opId);
	if (!definition) {
		errors.push(createAiError(AI_ERROR_CODES.UNKNOWN_OP, `Unknown operation: ${opId}.`, { op: opId }));
		return errors;
	}

	const { op: _ignored, ...payload } = op;
	const issues = validateAgainstSchema(payload, definition.parameters, `ops[${index}]`);
	for (const issue of issues) {
		errors.push(createAiError(AI_ERROR_CODES.SCHEMA_INVALID, issue.message, {
			op: opId,
			field: issue.path,
		}));
	}

	return errors;
}

export function validateDocumentOps(ops: DocumentOp[]): AiErrorDetail[] {
	if (!Array.isArray(ops)) {
		return [createAiError(AI_ERROR_CODES.SCHEMA_INVALID, 'ops must be an array.', { field: 'ops' })];
	}

	const errors: AiErrorDetail[] = [];
	ops.forEach((op, index) => {
		errors.push(...validateDocumentOp(op, index));
	});
	return errors;
}
