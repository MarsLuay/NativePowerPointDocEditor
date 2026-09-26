import { debugLog } from '../logger';
import type { AiRuntime } from './aiRuntime';
import { describeDocxFromBuffer } from './docxDescribe';
import { executeDocxOp } from './docxOpExecutor';
import { saveDocxToVault, validateDocxDocumentXmlLight } from './docxSave';
import { DocxSessionManager } from './docxSessionManager';
import { AI_EDIT_UNDO_LABEL, aiUndoStore } from './aiUndoStore';
import { AI_ERROR_CODES, createAiError, isAiErrorDetail } from './errors';
import { getOpDefinition, validateDocumentOps } from './opRegistry';
import type { ApplyOptions, ApplyResult, DescribeResult, DocumentOp } from './types';

export class DocxDocumentService {
	private readonly sessions: DocxSessionManager;
	private readonly operationLocks = new Map<string, Promise<void>>();

	constructor(private readonly runtime: AiRuntime) {
		this.sessions = new DocxSessionManager(runtime);
	}

	resolveDocxFile(path: string) {
		return this.sessions.resolveDocxFile(path);
	}

	private async withPathLock<T>(path: string, operation: () => Promise<T>): Promise<T> {
		const normalized = this.runtime.normalizePath(path);
		const previous = this.operationLocks.get(normalized) ?? Promise.resolve();
		let release!: () => void;
		const current = new Promise<void>((resolve) => {
			release = resolve;
		});
		this.operationLocks.set(normalized, current);
		await previous;
		try {
			return await operation();
		} finally {
			release();
			if (this.operationLocks.get(normalized) === current) {
				this.operationLocks.delete(normalized);
			}
		}
	}

	async describe(path: string): Promise<DescribeResult> {
		return this.withPathLock(path, () => this.describeUnlocked(path));
	}

	private async describeUnlocked(path: string): Promise<DescribeResult> {
		const startedAt = performance.now();
		try {
			const lease = await this.sessions.acquire(path);
			const buffer = lease.mode === 'view'
				? await lease.view.exportBufferForAgent()
				: await lease.patch.export();
			if (!buffer) {
				throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `DOCX view has no exportable document for ${path}.`, { path });
			}
			const file = lease.file;
			const snapshot = await describeDocxFromBuffer(buffer, file.path);
			debugLog('agent', 'AI DOCX describe completed', {
				path: file.path,
				blockCount: snapshot.blockCount,
				ms: Math.round(performance.now() - startedAt),
			});
			return { ok: true, errors: [], snapshot };
		} catch (error) {
			debugLog('agent', 'AI DOCX describe failed', {
				path,
				error: isAiErrorDetail(error) ? error.message : String(error),
				ms: Math.round(performance.now() - startedAt),
			});
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
		return this.withPathLock(path, () => this.applyUnlocked(path, ops, options));
	}

	private async applyUnlocked(path: string, ops: DocumentOp[], options: ApplyOptions = {}): Promise<ApplyResult> {
		const startedAt = performance.now();
		const dryRun = options.dryRun === true;
		const validationErrors = validateDocumentOps(ops);
		if (validationErrors.length > 0) {
			return { ok: false, dryRun, warnings: [], errors: validationErrors };
		}

		const unsupported = ops
			.map((op) => getOpDefinition(String(op.op)))
			.filter((definition) => definition && definition.namespace !== 'docx');
		if (unsupported.length > 0) {
			return {
				ok: false,
				dryRun,
				warnings: [],
				errors: unsupported.map((definition) =>
					createAiError(
						AI_ERROR_CODES.NOT_IMPLEMENTED,
						`Operation ${definition?.id} is not a DOCX operation.`,
						{ op: definition?.id, path },
					),
				),
			};
		}

		try {
			const lease = await this.sessions.acquire(path);
			const beforeBuffer = lease.sourceBuffer.slice(0);
			const patch = dryRun ? await lease.patch.clone() : lease.patch;
			const originalXml = patch.getDocumentXml();
			let documentXml = originalXml;
			const changed = new Set<string>();
			const created = new Set<string>();
			const preview: ApplyResult['preview'] = [];
			const warnings: string[] = [];

			for (const op of ops) {
				patch.setDocumentXml(documentXml);
				const result = await executeDocxOp(
					{
						session: patch,
						vault: this.runtime.vault,
						filePath: lease.file.path,
						dryRun,
					},
					op,
				);
				documentXml = result.documentXml;
				for (const id of result.changedIds) changed.add(id);
				for (const id of result.createdIds) created.add(id);
				preview.push(...(result.preview ?? []));
				warnings.push(...result.warnings);
			}

			validateDocxDocumentXmlLight(documentXml);

			if (!dryRun) {
				lease.patch.setDocumentXml(documentXml);
				aiUndoStore.record(lease.file.path, {
					label: AI_EDIT_UNDO_LABEL,
					before: { kind: 'docx', buffer: beforeBuffer },
				});
				if (lease.mode === 'view') {
					const output = await lease.patch.export();
					await lease.view.reloadFromAgentBuffer(output);
				}
			}

			debugLog('agent', 'AI DOCX apply completed', {
				path: lease.file.path,
				mode: lease.mode,
				dryRun,
				opIds: ops.map((op) => String(op.op)),
				changedIds: [...changed],
				opCount: ops.length,
				changedCount: changed.size,
				ms: Math.round(performance.now() - startedAt),
			});

			return {
				ok: true,
				dryRun,
				changed: [...changed],
				created: created.size > 0 ? [...created] : undefined,
				undoLabel: dryRun ? undefined : AI_EDIT_UNDO_LABEL,
				canUndo: !dryRun && aiUndoStore.canUndo(lease.file.path),
				preview: preview.length > 0 ? preview : undefined,
				warnings,
				errors: [],
			};
		} catch (error) {
			debugLog('agent', 'AI DOCX apply failed', {
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
		return this.withPathLock(path, () => this.saveUnlocked(path));
	}

	private async saveUnlocked(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		const startedAt = performance.now();
		try {
			const lease = await this.sessions.acquire(path);
			const output = await saveDocxToVault(
				this.runtime.vault,
				lease.file,
				lease.patch,
				lease.sourceBuffer,
			);

			if (lease.mode === 'headless') {
				this.sessions.updateAfterSave(lease.session, output);
			} else {
				await lease.view.reloadFromAgentBuffer(output);
			}
			debugLog('agent', 'AI DOCX save completed', {
				path: lease.file.path,
				mode: lease.mode,
				bytes: output.byteLength,
				ms: Math.round(performance.now() - startedAt),
			});

			return { ok: true, errors: [] };
		} catch (error) {
			debugLog('agent', 'AI DOCX save failed', {
				path,
				error: isAiErrorDetail(error) ? error.message : String(error),
				ms: Math.round(performance.now() - startedAt),
			});
			if (isAiErrorDetail(error)) {
				return { ok: false, errors: [error] };
			}
			return {
				ok: false,
				errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, String(error), { path })],
			};
		}
	}

	async close(path: string): Promise<void> {
		// Keep AI undo/redo snapshots after session.close() so agents can undo
		// apply→save→close without falling back to git restore.
		await this.withPathLock(path, () => this.sessions.release(path));
	}

	async undo(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		return this.withPathLock(path, () => this.undoUnlocked(path));
	}

	private async undoUnlocked(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		try {
			const normalized = this.runtime.normalizePath(path);
			const entry = aiUndoStore.popUndo(normalized);
			if (!entry || entry.before.kind !== 'docx') {
				const openView = this.runtime.findOpenDocxView(normalized);
				if (openView?.canUndoAgentEdit()) {
					const undone = await openView.undoAgentEdit();
					if (!undone) {
						return {
							ok: false,
							errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to undo.', { path })],
						};
					}
					const saved = await this.saveUnlocked(path);
					return saved.ok ? { ok: true, errors: [] } : saved;
				}
				return {
					ok: false,
					errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to undo.', { path })],
				};
			}

			const lease = await this.sessions.acquire(path);
			const currentBuffer = lease.mode === 'view'
				? await lease.view.exportBufferForAgent()
				: await lease.patch.export();
			if (currentBuffer) {
				aiUndoStore.pushRedo(normalized, {
					label: entry.label,
					before: { kind: 'docx', buffer: currentBuffer.slice(0) },
				});
			}
			await this.sessions.restoreHeadlessBuffer(path, entry.before.buffer);
			if (lease.mode === 'view') {
				await lease.view.reloadFromAgentBuffer(entry.before.buffer);
			}
			const saved = await this.saveUnlocked(path);
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
		return this.withPathLock(path, () => this.redoUnlocked(path));
	}

	private async redoUnlocked(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		try {
			const normalized = this.runtime.normalizePath(path);
			const entry = aiUndoStore.popRedo(normalized);
			if (!entry || entry.before.kind !== 'docx') {
				const openView = this.runtime.findOpenDocxView(normalized);
				if (openView?.canRedoAgentEdit()) {
					const redone = await openView.redoAgentEdit();
					if (!redone) {
						return {
							ok: false,
							errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to redo.', { path })],
						};
					}
					const saved = await this.saveUnlocked(path);
					return saved.ok ? { ok: true, errors: [] } : saved;
				}
				return {
					ok: false,
					errors: [createAiError(AI_ERROR_CODES.VALIDATION_FAILED, 'Nothing to redo.', { path })],
				};
			}

			const lease = await this.sessions.acquire(path);
			const currentBuffer = lease.mode === 'view'
				? await lease.view.exportBufferForAgent()
				: await lease.patch.export();
			if (currentBuffer) {
				aiUndoStore.record(normalized, {
					label: entry.label,
					before: { kind: 'docx', buffer: currentBuffer.slice(0) },
				});
			}
			await this.sessions.restoreHeadlessBuffer(path, entry.before.buffer);
			if (lease.mode === 'view') {
				await lease.view.reloadFromAgentBuffer(entry.before.buffer);
			}
			const saved = await this.saveUnlocked(path);
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
