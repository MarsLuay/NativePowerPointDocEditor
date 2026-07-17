import {
	DocumentSaveCoordinator,
	type DocumentSaveCoordinatorOptions,
	type DocumentSaveState,
} from '../../save/DocumentSaveCoordinator';

export type DocxSaveSource = 'manual' | 'autosave';

export interface DocxSessionSnapshot {
	readonly dirty: boolean;
	readonly editVersion: number;
	readonly saveState: DocumentSaveState;
	readonly saveError?: unknown;
}

export type DocxSessionListener = (snapshot: DocxSessionSnapshot) => void;

export type DocxSessionOptions<Context, Serialized, Prepared, Validated> = Omit<
	DocumentSaveCoordinatorOptions<Context, Serialized, Prepared, Validated, DocxSaveSource>,
	'onStateChange'
>;

/**
 * Owns DOCX dirty/save state. The view supplies persistence details while
 * React subscribes to this stable state boundary.
 */
export class DocxSession<Context, Serialized, Prepared, Validated> {
	private readonly listeners = new Set<DocxSessionListener>();
	private saveError: unknown;
	readonly saveCoordinator: DocumentSaveCoordinator<Context, Serialized, Prepared, Validated, DocxSaveSource>;

	constructor(options: DocxSessionOptions<Context, Serialized, Prepared, Validated>) {
		this.saveCoordinator = new DocumentSaveCoordinator({
			...options,
			onStateChange: (_state, error) => {
				this.saveError = error;
				this.emit();
			},
		});
	}

	get dirty(): boolean {
		return this.saveCoordinator.dirty;
	}

	get editVersion(): number {
		return this.saveCoordinator.version;
	}

	get saveState(): DocumentSaveState {
		return this.saveCoordinator.state;
	}

	snapshot(): DocxSessionSnapshot {
		return {
			dirty: this.dirty,
			editVersion: this.editVersion,
			saveState: this.saveState,
			saveError: this.saveError,
		};
	}

	subscribe(listener: DocxSessionListener): () => void {
		this.listeners.add(listener);
		return () => this.listeners.delete(listener);
	}

	markDirty(): void {
		this.saveCoordinator.markDirty();
	}

	save(source: DocxSaveSource = 'manual'): Promise<boolean> {
		return this.saveCoordinator.save(source);
	}

	clearAutosave(): void {
		this.saveCoordinator.clearAutosave();
	}

	reset(): void {
		this.saveCoordinator.reset();
	}

	private emit(): void {
		const snapshot = this.snapshot();
		for (const listener of this.listeners) listener(snapshot);
	}
}
