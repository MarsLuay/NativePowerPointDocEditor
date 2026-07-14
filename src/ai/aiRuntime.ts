import type { Vault } from 'obsidian';
import type { PresentationEngine } from '../PresentationEngine';

export interface DocxViewAgentBridge {
	getLoadedDocumentPath(): string | null;
	canAgentEdit(): boolean;
	exportBufferForAgent(): Promise<ArrayBuffer | null>;
	reloadFromAgentBuffer(buffer: ArrayBuffer): Promise<void>;
	saveCurrentDocument(): Promise<boolean>;
	canUndoAgentEdit(): boolean;
	canRedoAgentEdit(): boolean;
	undoAgentEdit(): Promise<boolean>;
	redoAgentEdit(): Promise<boolean>;
}

export interface PptxViewAgentBridge {
	getLoadedPresentationPath(): string | null;
	getPresentationEngineForAgent(): PresentationEngine | null;
	canAgentEdit(): boolean;
	runAgentEditBatch(
		label: string,
		mutate: (engine: PresentationEngine) => Promise<number[]>,
	): Promise<void>;
	saveCurrentPresentation(): Promise<boolean>;
	canUndoAgentEdit(): boolean;
	canRedoAgentEdit(): boolean;
	undoAgentEdit(): Promise<boolean>;
	redoAgentEdit(): Promise<boolean>;
}

export interface AiRuntime {
	vault: Vault;
	normalizePath(path: string): string;
	findOpenDocxView(path: string): DocxViewAgentBridge | null;
	findOpenPptxView(path: string): PptxViewAgentBridge | null;
}

export function createAiRuntime(options: AiRuntime): AiRuntime {
	return options;
}
