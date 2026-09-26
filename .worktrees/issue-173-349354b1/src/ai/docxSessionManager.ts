import { TFile } from 'obsidian';
import type { AiRuntime, DocxViewAgentBridge } from './aiRuntime';
import { DocxPatchSession } from './docxPatchSession';
import { AI_ERROR_CODES, createAiError } from './errors';

export type { AiRuntime, DocxViewAgentBridge } from './aiRuntime';

export interface HeadlessDocxSession {
	file: TFile;
	patch: DocxPatchSession;
	sourceBuffer: ArrayBuffer;
}

export type DocxEditingLease =
	| {
		mode: 'view';
		file: TFile;
		patch: DocxPatchSession;
		sourceBuffer: ArrayBuffer;
		view: DocxViewAgentBridge;
	}
	| {
		mode: 'headless';
		file: TFile;
		patch: DocxPatchSession;
		sourceBuffer: ArrayBuffer;
		session: HeadlessDocxSession;
	};

export class DocxSessionManager {
	private readonly sessions = new Map<string, HeadlessDocxSession>();

	constructor(private readonly runtime: AiRuntime) {}

	resolveDocxFile(path: string): TFile {
		const normalized = this.runtime.normalizePath(path);
		const file = this.runtime.vault.getAbstractFileByPath(normalized);
		if (!(file instanceof TFile)) {
			throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `File not found: ${normalized}.`, { path: normalized });
		}
		if (file.extension.toLowerCase() !== 'docx') {
			throw createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Not a DOCX file: ${normalized}.`, { path: normalized });
		}
		return file;
	}

	async acquire(path: string): Promise<DocxEditingLease> {
		const file = this.resolveDocxFile(path);
		const normalized = file.path;
		const openView = this.runtime.findOpenDocxView(normalized);
		if (openView) {
			this.sessions.delete(normalized);
			if (!openView.canAgentEdit()) {
				throw createAiError(AI_ERROR_CODES.OBJECT_NOT_EDITABLE, `DOCX view is not ready for agent edits: ${normalized}.`, {
					path: normalized,
				});
			}

			const sourceBuffer = await openView.exportBufferForAgent();
			if (!sourceBuffer) {
				throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `DOCX view has no exportable document for ${normalized}.`, {
					path: normalized,
				});
			}

			const patch = await DocxPatchSession.load(sourceBuffer);
			return { mode: 'view', file, patch, sourceBuffer, view: openView };
		}

		let session = this.sessions.get(normalized);
		if (!session) {
			const sourceBuffer = await this.runtime.vault.readBinary(file);
			const patch = await DocxPatchSession.load(sourceBuffer);
			session = { file, patch, sourceBuffer };
			this.sessions.set(normalized, session);
		}

		return {
			mode: 'headless',
			file: session.file,
			patch: session.patch,
			sourceBuffer: session.sourceBuffer,
			session,
		};
	}

	async release(path: string): Promise<void> {
		this.sessions.delete(this.runtime.normalizePath(path));
	}

	updateAfterSave(session: HeadlessDocxSession, output: ArrayBuffer): void {
		session.sourceBuffer = output.slice(0);
	}

	async restoreHeadlessBuffer(path: string, buffer: ArrayBuffer): Promise<HeadlessDocxSession> {
		const file = this.resolveDocxFile(path);
		const normalized = file.path;
		const patch = await DocxPatchSession.load(buffer.slice(0));
		const session: HeadlessDocxSession = {
			file,
			patch,
			sourceBuffer: buffer.slice(0),
		};
		this.sessions.set(normalized, session);
		return session;
	}
}
