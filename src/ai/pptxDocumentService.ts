import { debugLog } from '../logger';
import { isPowerPointExtension } from '../powerpoint/extensions';
import type { AiRuntime } from './aiRuntime';
import { AI_EDIT_UNDO_LABEL, aiUndoStore } from './aiUndoStore';
import { AI_ERROR_CODES, createAiError } from './errors';
import { coalescePptxOps } from './coalescePptxOps';
import { describePptxFromEngine } from './pptxDescribe';
import { executePptxOp } from './pptxOpExecutor';
import { isAiErrorDetail } from './errors';
import { savePptxToVault } from './pptxSave';
import {
	exportPresentationToPdfBytes,
	toExportPdfFailure,
	writeExportPdfArtifact,
	type ExportPdfOptions,
	type ExportPdfResult,
} from './pptxExportPdf';
import { PptxSessionManager } from './pptxSessionManager';
import { getOpDefinition, validateDocumentOps } from './opRegistry';
import type { ApplyOptions, ApplyPreviewChange, ApplyResult, DescribeResult, DocumentOp } from './types';

export type { ExportPdfOptions, ExportPdfResult };

function isDocxPath(path: string): boolean {
	return path.toLowerCase().endsWith('.docx');
}

export class PptxDocumentService {
	private readonly sessions: PptxSessionManager;

	constructor(private readonly runtime: AiRuntime) {
		this.sessions = new PptxSessionManager(runtime);
	}

	get sessionManager(): PptxSessionManager {
		return this.sessions;
	}

	private assertPptxPath(path: string): void {
		const extension = path.split('.').pop()?.toLowerCase() ?? '';
		if (isDocxPath(path)) {
			throw createAiError(
				AI_ERROR_CODES.UNSUPPORTED_FORMAT,
				'Use the DOCX AI describe path for .docx files.',
				{ path },
			);
		}
		if (!isPowerPointExtension(extension)) {
			throw createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported file format: ${path}.`, { path });
		}
	}

	async describe(path: string): Promise<DescribeResult> {
		const startedAt = performance.now();
		try {
			this.assertPptxPath(path);
			const lease = await this.sessions.acquire(path, { requireEditable: false });
			const snapshot = describePptxFromEngine(lease.engine, lease.file.path);
			debugLog('agent', 'AI describe completed', {
				path: lease.file.path,
				slideCount: snapshot.slideCount,
				ms: Math.round(performance.now() - startedAt),
			});
			return { ok: true, errors: [], snapshot };
		} catch (error) {
			if (isAiErrorDetail(error)) {
				return { ok: false, errors: [error] };
			}
			return {
				ok: false,
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}

	async apply(path: string, ops: DocumentOp[], options: ApplyOptions = {}): Promise<ApplyResult> {
		const startedAt = performance.now();
		const dryRun = options.dryRun === true;
		const validationErrors = validateDocumentOps(ops);
		if (validationErrors.length > 0) {
			return { ok: false, dryRun, warnings: [], errors: validationErrors };
		}

		const unsupported = ops
			.map((op) => getOpDefinition(String(op.op)))
			.filter((definition) => definition && definition.namespace !== 'pptx');
		if (unsupported.length > 0) {
			return {
				ok: false,
				dryRun,
				warnings: [],
				errors: unsupported.map((definition) =>
					createAiError(
						AI_ERROR_CODES.NOT_IMPLEMENTED,
						`Operation ${definition?.id} is not implemented yet.`,
						{ op: definition?.id, path },
					),
				),
			};
		}

		try {
			this.assertPptxPath(path);
			const lease = await this.sessions.acquire(path);
			const changed = new Set<string>();
			const preview: ApplyPreviewChange[] = [];
			const warnings: string[] = [];
			const affectedSlideIndices = new Set<number>();

			const runBatch = async (engine: typeof lease.engine) => {
				// Coalesce consecutive same-slide deletes so renumbering cannot
				// make later shapeIndex payloads miss their targets.
				const batchOps = coalescePptxOps(ops);
				for (const op of batchOps) {
					const result = await executePptxOp(
						{
							engine,
							vault: this.runtime.vault,
							filePath: lease.file.path,
							dryRun,
						},
						op,
					);
					for (const id of result.changedIds) changed.add(id);
					preview.push(...result.preview);
					warnings.push(...result.warnings);
					for (const slideIndex of result.affectedSlideIndices) {
						affectedSlideIndices.add(slideIndex);
					}
				}
				return [...affectedSlideIndices];
			};

			if (lease.mode === 'view') {
				if (!dryRun && !lease.view.canAgentEdit()) {
					return {
						ok: false,
						dryRun,
						warnings: [],
						errors: [
							createAiError(
								AI_ERROR_CODES.OBJECT_NOT_EDITABLE,
								'The open PowerPoint view is read-only.',
								{ path: lease.file.path },
							),
						],
					};
				}

				if (dryRun) {
					await runBatch(lease.engine);
				} else {
					// Keep a durable pre-edit snapshot for AI undo even after session.close()
					// or view history clear on reload. View history still gets Ctrl+Z.
					const rollbackBuffer = await lease.engine.export();
					await lease.view.runAgentEditBatch('AI edit', (engine) => runBatch(engine));
					aiUndoStore.record(lease.file.path, {
						label: AI_EDIT_UNDO_LABEL,
						before: { kind: 'pptx', buffer: rollbackBuffer.slice(0), currentSlide: 0 },
					});
				}
			} else {
				const rollbackBuffer = dryRun ? null : await lease.engine.export();
				try {
					await runBatch(lease.engine);
				} catch (error) {
					if (rollbackBuffer) {
						await lease.engine.restoreSnapshot(rollbackBuffer);
					}
					throw error;
				}
				if (!dryRun && rollbackBuffer) {
					aiUndoStore.record(lease.file.path, {
						label: AI_EDIT_UNDO_LABEL,
						before: { kind: 'pptx', buffer: rollbackBuffer.slice(0), currentSlide: 0 },
					});
				}
			}

			debugLog('agent', 'AI apply completed', {
				path: lease.file.path,
				dryRun,
				opCount: ops.length,
				changedCount: changed.size,
				ms: Math.round(performance.now() - startedAt),
			});

			return {
				ok: true,
				dryRun,
				changed: [...changed],
				undoLabel: dryRun ? undefined : AI_EDIT_UNDO_LABEL,
				canUndo: !dryRun && aiUndoStore.canUndo(lease.file.path),
				preview: preview.length > 0 ? preview : undefined,
				warnings,
				errors: [],
			};
		} catch (error) {
			debugLog('agent', 'AI apply failed', {
				path,
				dryRun,
				error: isAiErrorDetail(error) ? error.message : String(error),
				ms: Math.round(performance.now() - startedAt),
			});
			if (isAiErrorDetail(error)) {
				return { ok: false, dryRun, warnings: [], errors: [error] };
			}
			return {
				ok: false,
				dryRun,
				warnings: [],
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}

	async save(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		try {
			this.assertPptxPath(path);
			const lease = await this.sessions.acquire(path);
			if (lease.mode === 'view') {
				const saved = await lease.view.saveCurrentPresentation();
				const saveError = lease.view.getAgentSaveError();
				if (!saved) {
					debugLog('agent', 'AI view session save failed', {
						path: lease.file.path,
						mode: 'view',
						error: saveError,
					});
				}
				return saved
					? { ok: true, errors: [] }
					: {
						ok: false,
						errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, saveError ?? 'Save failed in the open PowerPoint view.', {
							path: lease.file.path,
						})],
					};
			}

			const { output, sourcePackage } = await savePptxToVault(
				this.runtime.vault,
				lease.file,
				lease.engine,
				lease.session.sourceBuffer,
				lease.session.sourcePackage,
			);
			this.sessions.updateHeadlessSessionAfterSave(lease.session, output);
			lease.session.sourcePackage = sourcePackage;
			return { ok: true, errors: [] };
		} catch (error) {
			if (isAiErrorDetail(error)) {
				return { ok: false, errors: [error] };
			}
			return {
				ok: false,
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}

	async exportPdf(path: string, options: ExportPdfOptions = {}): Promise<ExportPdfResult> {
		const startedAt = performance.now();
		try {
			this.assertPptxPath(path);
			const lease = await this.sessions.acquire(path, { requireEditable: false });
			const { bytes, slideCount } = await exportPresentationToPdfBytes(lease.engine, options);
			const written = await writeExportPdfArtifact(this.runtime.vault, lease.file, bytes, options);
			debugLog('agent', 'AI exportPdf completed', {
				path: lease.file.path,
				outputPath: written.path,
				slideCount,
				bytes: written.bytes,
				ms: Math.round(performance.now() - startedAt),
			});
			return {
				ok: true,
				path: written.path,
				bytes: written.bytes,
				slideCount,
				errors: [],
			};
		} catch (error) {
			return toExportPdfFailure(error, path);
		}
	}

	async close(path: string): Promise<void> {
		// Keep AI undo/redo snapshots after session.close() so agents can undo
		// apply→save→close the same way Ctrl+Z undoes in the editor.
		await this.sessions.release(path);
	}

	private async exportCurrentPptxBuffer(path: string): Promise<ArrayBuffer> {
		const lease = await this.sessions.acquire(path, { requireEditable: false });
		return lease.engine.export();
	}

	private async restorePptxUndoBuffer(path: string, buffer: ArrayBuffer): Promise<void> {
		const openView = this.runtime.findOpenPptxView(this.runtime.normalizePath(path));
		if (openView) {
			await openView.reloadFromAgentBuffer(buffer);
			return;
		}
		await this.sessions.restoreHeadlessBuffer(path, buffer);
	}

	async undo(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		try {
			this.assertPptxPath(path);
			const normalized = this.runtime.normalizePath(path);
			const entry = aiUndoStore.popUndo(normalized);
			if (!entry || entry.before.kind !== 'pptx') {
				const openView = this.runtime.findOpenPptxView(normalized);
				if (openView?.canUndoAgentEdit()) {
					const undone = await openView.undoAgentEdit();
					if (!undone) {
						return {
							ok: false,
							errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to undo.', { path })],
						};
					}
					const saved = await this.save(path);
					return saved.ok ? { ok: true, errors: [] } : saved;
				}
				return {
					ok: false,
					errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to undo.', { path })],
				};
			}

			const currentBuffer = await this.exportCurrentPptxBuffer(path);
			aiUndoStore.pushRedo(normalized, {
				label: entry.label,
				before: { kind: 'pptx', buffer: currentBuffer.slice(0), currentSlide: 0 },
			});
			await this.restorePptxUndoBuffer(path, entry.before.buffer);
			const saved = await this.save(path);
			return saved.ok ? { ok: true, errors: [] } : saved;
		} catch (error) {
			if (isAiErrorDetail(error)) {
				return { ok: false, errors: [error] };
			}
			return {
				ok: false,
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}

	async redo(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		try {
			this.assertPptxPath(path);
			const normalized = this.runtime.normalizePath(path);
			const entry = aiUndoStore.popRedo(normalized);
			if (!entry || entry.before.kind !== 'pptx') {
				const openView = this.runtime.findOpenPptxView(normalized);
				if (openView?.canRedoAgentEdit()) {
					const redone = await openView.redoAgentEdit();
					if (!redone) {
						return {
							ok: false,
							errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to redo.', { path })],
						};
					}
					const saved = await this.save(path);
					return saved.ok ? { ok: true, errors: [] } : saved;
				}
				return {
					ok: false,
					errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to redo.', { path })],
				};
			}

			const currentBuffer = await this.exportCurrentPptxBuffer(path);
			aiUndoStore.record(normalized, {
				label: entry.label,
				before: { kind: 'pptx', buffer: currentBuffer.slice(0), currentSlide: 0 },
			});
			await this.restorePptxUndoBuffer(path, entry.before.buffer);
			const saved = await this.save(path);
			return saved.ok ? { ok: true, errors: [] } : saved;
		} catch (error) {
			if (isAiErrorDetail(error)) {
				return { ok: false, errors: [error] };
			}
			return {
				ok: false,
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}
}
