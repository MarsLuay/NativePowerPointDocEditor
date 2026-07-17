import { debugLog } from '../logger';
import { isPowerPointExtension } from '../powerpoint/extensions';
import type { AiRuntime } from './aiRuntime';
import { buildCapabilityManifest } from './capabilities';
import { createAiError, AI_ERROR_CODES } from './errors';
import type { AiErrorDetail } from './errors';
import { getOpDefinition, validateDocumentOps } from './opRegistry';
import type {
	ApplyOptions,
	ApplyResult,
	CapabilityManifest,
	DescribeResult,
	DocumentOp,
} from './types';

export interface AiCoreOptions {
	getEnabled: () => boolean;
	pluginVersion: string;
	runtime: AiRuntime | null;
}

type PptxDocumentServiceModule = typeof import('./pptxDocumentService');
type DocxDocumentServiceModule = typeof import('./docxDocumentService');

export class AiCore {
	private pptxService: InstanceType<PptxDocumentServiceModule['PptxDocumentService']> | null = null;
	private pptxServicePromise: Promise<InstanceType<PptxDocumentServiceModule['PptxDocumentService']> | null> | null = null;
	private docxService: InstanceType<DocxDocumentServiceModule['DocxDocumentService']> | null = null;
	private docxServicePromise: Promise<InstanceType<DocxDocumentServiceModule['DocxDocumentService']> | null> | null = null;

	constructor(private readonly options: AiCoreOptions) {}

	isEnabled(): boolean {
		return this.options.getEnabled();
	}

	private disabledError() {
		return createAiError(
			AI_ERROR_CODES.AI_DISABLED,
			'Enable AI-Interfacing in plugin settings to use the agent API.',
		);
	}

	private missingRuntimeError() {
		return createAiError(
			AI_ERROR_CODES.NOT_IMPLEMENTED,
			'AI runtime is not available.',
		);
	}

	private async resolvePptxService(): Promise<InstanceType<PptxDocumentServiceModule['PptxDocumentService']> | null> {
		if (!this.options.runtime) {
			return null;
		}
		if (this.pptxService) {
			return this.pptxService;
		}
		if (!this.pptxServicePromise) {
			this.pptxServicePromise = import('./pptxDocumentService').then(({ PptxDocumentService }) => {
				if (!this.options.runtime) return null;
			 this.pptxService = new PptxDocumentService(this.options.runtime);
				return this.pptxService;
			});
		}
		return this.pptxServicePromise;
	}

	private async resolveDocxService(): Promise<InstanceType<DocxDocumentServiceModule['DocxDocumentService']> | null> {
		if (!this.options.runtime) {
			return null;
		}
		if (this.docxService) {
			return this.docxService;
		}
		if (!this.docxServicePromise) {
			this.docxServicePromise = import('./docxDocumentService').then(({ DocxDocumentService }) => {
				if (!this.options.runtime) return null;
				this.docxService = new DocxDocumentService(this.options.runtime);
				return this.docxService;
			});
		}
		return this.docxServicePromise;
	}

	buildManifest(enabled = this.isEnabled()): CapabilityManifest {
		return buildCapabilityManifest({
			pluginVersion: this.options.pluginVersion,
			enabled,
		});
	}

	listCapabilities(): CapabilityManifest | { ok: false; errors: AiErrorDetail[] } {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}
		return this.buildManifest(true);
	}

	validateOps(ops: DocumentOp[]): { ok: boolean; errors: AiErrorDetail[] } {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}
		const errors = validateDocumentOps(ops);
		return { ok: errors.length === 0, errors };
	}

	async describe(path: string): Promise<DescribeResult> {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}
		if (!this.options.runtime) {
			return { ok: false, errors: [this.missingRuntimeError()] };
		}

		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (isPowerPointExtension(extension)) {
			const pptxService = await this.resolvePptxService();
			if (!pptxService) {
				return { ok: false, errors: [this.missingRuntimeError()] };
			}
			return pptxService.describe(path);
		}
		if (extension === 'docx') {
			const docxService = await this.resolveDocxService();
			if (!docxService) {
				return { ok: false, errors: [this.missingRuntimeError()] };
			}
			return docxService.describe(path);
		}

		return {
			ok: false,
			errors: [createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported file: ${path}.`, { path })],
		};
	}

	async apply(path: string, ops: DocumentOp[], options: ApplyOptions = {}): Promise<ApplyResult> {
		const startedAt = performance.now();
		if (!this.isEnabled()) {
			return { ok: false, warnings: [], errors: [this.disabledError()] };
		}

		const validation = this.validateOps(ops);
		if (!validation.ok) {
			debugLog('agent', 'AI apply rejected by schema validation', {
				path,
				dryRun: options.dryRun === true,
				errorCount: validation.errors.length,
				ms: Math.round(performance.now() - startedAt),
			});
			return { ok: false, warnings: [], errors: validation.errors };
		}

		const pptxOps = ops.filter((op) => getOpDefinition(String(op.op))?.namespace === 'pptx');
		const docxOps = ops.filter((op) => getOpDefinition(String(op.op))?.namespace === 'docx');
		if (pptxOps.length > 0 && docxOps.length > 0) {
			return {
				ok: false,
				dryRun: options.dryRun === true,
				warnings: [],
				errors: [
					createAiError(
						AI_ERROR_CODES.SCHEMA_INVALID,
						'Cannot mix PPTX and DOCX operations in one batch.',
						{ path },
					),
				],
			};
		}

		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (docxOps.length > 0) {
			if (extension !== 'docx') {
				return {
					ok: false,
					dryRun: options.dryRun === true,
					warnings: [],
					errors: [createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported file: ${path}.`, { path })],
				};
			}
			const docxService = await this.resolveDocxService();
			if (!docxService) {
				return { ok: false, dryRun: options.dryRun === true, warnings: [], errors: [this.missingRuntimeError()] };
			}
			return docxService.apply(path, ops, options);
		}

		if (!isPowerPointExtension(extension)) {
			return {
				ok: false,
				dryRun: options.dryRun === true,
				warnings: [],
				errors: [createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported file: ${path}.`, { path })],
			};
		}

		const pptxService = await this.resolvePptxService();
		if (!pptxService) {
			return { ok: false, dryRun: options.dryRun === true, warnings: [], errors: [this.missingRuntimeError()] };
		}

		return pptxService.apply(path, ops, options);
	}

	async saveSession(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}

		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (extension === 'docx') {
			const docxService = await this.resolveDocxService();
			if (!docxService) {
				return { ok: false, errors: [this.missingRuntimeError()] };
			}
			return docxService.save(path);
		}

		const pptxService = await this.resolvePptxService();
		if (!pptxService) {
			return { ok: false, errors: [this.missingRuntimeError()] };
		}
		return pptxService.save(path);
	}

	async closeSession(path: string): Promise<void> {
		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (extension === 'docx') {
			const docxService = await this.resolveDocxService();
			if (!docxService) return;
			await docxService.close(path);
			return;
		}
		const pptxService = await this.resolvePptxService();
		if (!pptxService) return;
		await pptxService.close(path);
	}

	async undoAgentEdit(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}

		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (extension === 'docx') {
			const docxService = await this.resolveDocxService();
			if (!docxService) {
				return { ok: false, errors: [this.missingRuntimeError()] };
			}
			return docxService.undo(path);
		}

		const pptxService = await this.resolvePptxService();
		if (!pptxService) {
			return { ok: false, errors: [this.missingRuntimeError()] };
		}
		return pptxService.undo(path);
	}

	async redoAgentEdit(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		if (!this.isEnabled()) {
			return { ok: false, errors: [this.disabledError()] };
		}

		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (extension === 'docx') {
			const docxService = await this.resolveDocxService();
			if (!docxService) {
				return { ok: false, errors: [this.missingRuntimeError()] };
			}
			return docxService.redo(path);
		}

		const pptxService = await this.resolvePptxService();
		if (!pptxService) {
			return { ok: false, errors: [this.missingRuntimeError()] };
		}
		return pptxService.redo(path);
	}
}
