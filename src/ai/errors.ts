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
const AI_ERROR_CODE_VALUES: ReadonlySet<string> = new Set(Object.values(AI_ERROR_CODES));

export interface AiErrorDetail {
	code: AiErrorCode;
	message: string;
	op?: string;
	path?: string;
	field?: string;
	extension?: string;
}

type AiErrorContext = Partial<Pick<AiErrorDetail, 'op' | 'path' | 'field' | 'extension'>>;

export class AiError extends Error implements AiErrorDetail {
	readonly code: AiErrorCode;
	readonly op?: string;
	readonly path?: string;
	readonly field?: string;
	readonly extension?: string;

	constructor(code: AiErrorCode, message: string, details: AiErrorContext = {}) {
		super(message);
		this.name = 'AiError';
		this.code = code;
		Object.assign(this, details);
	}

	toJSON(): AiErrorDetail {
		return {
			code: this.code,
			message: this.message,
			...(this.op === undefined ? {} : { op: this.op }),
			...(this.path === undefined ? {} : { path: this.path }),
			...(this.field === undefined ? {} : { field: this.field }),
			...(this.extension === undefined ? {} : { extension: this.extension }),
		};
	}
}

export function createAiError(
	code: AiErrorCode,
	message: string,
	details: AiErrorContext = {},
): AiError {
	return new AiError(code, message, details);
}

export function isAiErrorDetail(error: unknown): error is AiErrorDetail {
	return Boolean(
		error
		&& typeof error === 'object'
		&& 'code' in error
		&& typeof error.code === 'string'
		&& AI_ERROR_CODE_VALUES.has(error.code)
		&& 'message' in error
		&& typeof error.message === 'string',
	);
}
