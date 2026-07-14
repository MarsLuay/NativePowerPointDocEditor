export const AI_ERROR_CODES = {
	AI_DISABLED: 'AI_DISABLED',
	SCHEMA_INVALID: 'SCHEMA_INVALID',
	UNKNOWN_OP: 'UNKNOWN_OP',
	NOT_IMPLEMENTED: 'NOT_IMPLEMENTED',
	FILE_NOT_FOUND: 'FILE_NOT_FOUND',
	UNSUPPORTED_FORMAT: 'UNSUPPORTED_FORMAT',
	VALIDATION_FAILED: 'VALIDATION_FAILED',
	SHAPE_NOT_FOUND: 'SHAPE_NOT_FOUND',
	SLIDE_NOT_FOUND: 'SLIDE_NOT_FOUND',
	OBJECT_NOT_EDITABLE: 'OBJECT_NOT_EDITABLE',
	BLOCK_NOT_FOUND: 'BLOCK_NOT_FOUND',
} as const;

export type AiErrorCode = (typeof AI_ERROR_CODES)[keyof typeof AI_ERROR_CODES];

export interface AiErrorDetail {
	code: AiErrorCode | string;
	message: string;
	op?: string;
	path?: string;
	field?: string;
	extension?: string;
}

export function createAiError(
	code: AiErrorCode | string,
	message: string,
	details: Partial<Pick<AiErrorDetail, 'op' | 'path' | 'field' | 'extension'>> = {},
): AiErrorDetail {
	return { code, message, ...details };
}

export function isAiErrorDetail(error: unknown): error is AiErrorDetail {
	return Boolean(error && typeof error === 'object' && 'code' in error && 'message' in error);
}
