import { TFile } from 'obsidian';
import type { PresentationEngine } from '../PresentationEngine';
import { PresentationEngine as PresentationEngineClass } from '../PresentationEngine';
import { inspectPowerPointPackage, type PowerPointPackageInspection } from '../PowerPointPackage';
import {
	isEditablePowerPointExtension,
	isLegacyPowerPointExtension,
	isModernPowerPointExtension,
	isPowerPointExtension,
} from '../powerpoint/extensions';
import type { AiRuntime, PptxViewAgentBridge } from './aiRuntime';
import { AI_ERROR_CODES, createAiError } from './errors';

export type { AiRuntime, PptxViewAgentBridge } from './aiRuntime';

export interface HeadlessPptxSession {
	file: TFile;
	engine: PresentationEngine;
	sourceBuffer: ArrayBuffer;
	sourcePackage: PowerPointPackageInspection;
}

export type PptxEditingLease =
	| {
		mode: 'view';
		file: TFile;
		engine: PresentationEngine;
		view: PptxViewAgentBridge;
	}
	| {
		mode: 'headless';
		file: TFile;
		engine: PresentationEngine;
		session: HeadlessPptxSession;
	};

export class PptxSessionManager {
	private readonly headlessSessions = new Map<string, HeadlessPptxSession>();

	constructor(private readonly runtime: AiRuntime) {}

	resolvePptxFile(path: string, options: { requireEditable?: boolean } = {}): TFile {
		const requireEditable = options.requireEditable !== false;
		const normalized = this.runtime.normalizePath(path);
		const file = this.runtime.vault.getAbstractFileByPath(normalized);
		if (!(file instanceof TFile)) {
			throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `File not found: ${normalized}.`, { path: normalized });
		}
		if (!isPowerPointExtension(file.extension)) {
			throw createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Not a PowerPoint file: ${normalized}.`, {
				path: normalized,
			});
		}
		if (isLegacyPowerPointExtension(file.extension)) {
			throw createAiError(
				AI_ERROR_CODES.UNSUPPORTED_FORMAT,
				`Legacy binary PowerPoint (.ppt/.pps/.pot) is not supported: ${normalized}. Save as .pptx first.`,
				{ path: normalized, extension: file.extension },
			);
		}
		if (!isModernPowerPointExtension(file.extension)) {
			throw createAiError(AI_ERROR_CODES.UNSUPPORTED_FORMAT, `Unsupported PowerPoint format: ${normalized}.`, {
				path: normalized,
			});
		}
		if (requireEditable && !isEditablePowerPointExtension(file.extension)) {
			throw createAiError(AI_ERROR_CODES.OBJECT_NOT_EDITABLE, `PowerPoint file is view-only: ${normalized}.`, {
				path: normalized,
			});
		}
		return file;
	}

	async acquire(path: string, options: { requireEditable?: boolean } = {}): Promise<PptxEditingLease> {
		const requireEditable = options.requireEditable !== false;
		const file = this.resolvePptxFile(path, { requireEditable });
		const normalized = file.path;
		const openView = this.runtime.findOpenPptxView(normalized);
		if (openView) {
			this.headlessSessions.delete(normalized);
			const engine = openView.getPresentationEngineForAgent();
			if (!engine) {
				throw createAiError(AI_ERROR_CODES.FILE_NOT_FOUND, `PowerPoint view has no loaded engine for ${normalized}.`, {
					path: normalized,
				});
			}
			return { mode: 'view', file, engine, view: openView };
		}

		let session = this.headlessSessions.get(normalized);
		if (!session) {
			const sourceBuffer = await this.runtime.vault.readBinary(file);
			const engine = await PresentationEngineClass.load(sourceBuffer);
			session = {
				file,
				engine,
				sourceBuffer,
				sourcePackage: inspectPowerPointPackage(sourceBuffer),
			};
			this.headlessSessions.set(normalized, session);
		}

		return { mode: 'headless', file, engine: session.engine, session };
	}

	async release(path: string): Promise<void> {
		this.headlessSessions.delete(this.runtime.normalizePath(path));
	}

	updateHeadlessSessionAfterSave(session: HeadlessPptxSession, output: ArrayBuffer): void {
		session.sourceBuffer = output.slice(0);
		session.sourcePackage = inspectPowerPointPackage(output);
	}

	async restoreHeadlessBuffer(path: string, buffer: ArrayBuffer): Promise<HeadlessPptxSession> {
		const file = this.resolvePptxFile(path);
		const normalized = file.path;
		const engine = await PresentationEngineClass.load(buffer.slice(0));
		const session: HeadlessPptxSession = {
			file,
			engine,
			sourceBuffer: buffer.slice(0),
			sourcePackage: inspectPowerPointPackage(buffer),
		};
		this.headlessSessions.set(normalized, session);
		return session;
	}
}
