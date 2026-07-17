import type { AiCore } from './aiCore';
import { AI_API_VERSION, PLUGIN_ID } from './types';
import type {
	ApplyOptions,
	ApplyResult,
	CapabilityManifest,
	DescribeResult,
	DocumentOp,
} from './types';

export interface NpdeAiApiInfo {
	apiVersion: number;
	pluginId: string;
	enabled: boolean;
}

export interface AiDocumentSession {
	readonly path: string;
	describe(): Promise<DescribeResult>;
	apply(ops: DocumentOp[], options?: ApplyOptions): Promise<ApplyResult>;
	save(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }>;
	undo(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }>;
	redo(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }>;
	close(): Promise<void>;
}

export interface NpdeAiApi {
	isEnabled(): boolean;
	getInfo(): NpdeAiApiInfo;
	listCapabilities(): CapabilityManifest | { ok: false; errors: ApplyResult['errors'] };
	validateOps(ops: DocumentOp[]): { ok: boolean; errors: ApplyResult['errors'] };
	describe(path: string): Promise<DescribeResult>;
	apply(path: string, ops: DocumentOp[], options?: ApplyOptions): Promise<ApplyResult>;
	undo(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }>;
	redo(path: string): Promise<{ ok: boolean; errors: ApplyResult['errors'] }>;
	openSession(path: string): Promise<AiDocumentSession>;
}

class AiDocumentSessionImpl implements AiDocumentSession {
	constructor(
		readonly path: string,
		private readonly core: AiCore,
	) {}

	describe(): Promise<DescribeResult> {
		return this.core.describe(this.path);
	}

	apply(ops: DocumentOp[], options?: ApplyOptions): Promise<ApplyResult> {
		return this.core.apply(this.path, ops, options);
	}

	save(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		return this.core.saveSession(this.path);
	}

	undo(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		return this.core.undoAgentEdit(this.path);
	}

	redo(): Promise<{ ok: boolean; errors: ApplyResult['errors'] }> {
		return this.core.redoAgentEdit(this.path);
	}

	close(): Promise<void> {
		return this.core.closeSession(this.path);
	}
}

export function createNpdeAiApi(core: AiCore): NpdeAiApi {
	return {
		isEnabled: () => core.isEnabled(),
		getInfo: () => ({
			apiVersion: AI_API_VERSION,
			pluginId: PLUGIN_ID,
			enabled: core.isEnabled(),
		}),
		listCapabilities: () => core.listCapabilities(),
		validateOps: (ops) => core.validateOps(ops),
		describe: (path) => core.describe(path),
		apply: (path, ops, options) => core.apply(path, ops, options),
		undo: (path) => core.undoAgentEdit(path),
		redo: (path) => core.redoAgentEdit(path),
		openSession: (path) => Promise.resolve(new AiDocumentSessionImpl(path, core)),
	};
}
