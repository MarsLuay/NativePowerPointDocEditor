export type DocumentSaveState = 'clean' | 'dirty' | 'saving' | 'failed';

export interface DocumentSaveRequest<Source extends string> {
	source: Source;
	targetVersion: number;
}

export interface DocumentSaveAdapter<Context, Serialized, Prepared, Validated, Source extends string> {
	serialize(context: Context, request: DocumentSaveRequest<Source>): Promise<Serialized>;
	prepareForWrite(serialized: Serialized, context: Context, request: DocumentSaveRequest<Source>): Promise<Prepared>;
	validate(prepared: Prepared, context: Context, request: DocumentSaveRequest<Source>): Promise<Validated>;
	persist(prepared: Prepared, validated: Validated, context: Context, request: DocumentSaveRequest<Source>): Promise<void>;
}

export type DocumentSavePhaseName = 'serialize' | 'prepare' | 'validate' | 'persist';

export interface DocumentSavePhaseMeasurement {
	name: DocumentSavePhaseName;
	durationMs: number;
	synchronousWorkMs: number;
}

export interface DocumentSaveMetrics {
	documentBytes?: number | null;
	changedContentBytes?: number | null;
	outputBytes?: number | null;
}

export interface DocumentSaveMeasurement<Source extends string> {
	source: Source;
	targetVersion: number;
	startedAt: number;
	endedAt: number;
	phases: DocumentSavePhaseMeasurement[];
	documentBytes: number | null;
	changedContentBytes: number | null;
	outputBytes: number | null;
}

export interface DocumentSaveCoordinatorOptions<Context, Serialized, Prepared, Validated, Source extends string> {
	adapter: DocumentSaveAdapter<Context, Serialized, Prepared, Validated, Source>;
	getContext(): Context | null;
	autosave: {
		enabled(): boolean;
		delayMs(): number;
		source: Source;
	};
	onStateChange?(state: DocumentSaveState, error?: unknown): void;
	onAutosaveScheduled?(delayMs: number, version: number): void;
	onAutosaveStarted?(version: number): void;
	onSaveStarted?(request: DocumentSaveRequest<Source>, context: Context, startedAt: number): void;
	onSaveCompleted?(measurement: DocumentSaveMeasurement<Source>): void;
	getSaveMetrics?(context: Context, serialized?: Serialized): DocumentSaveMetrics;
	now?: () => number;
	/** If set, autosave timers call this instead of coordinator.save(). */
	runAutosave?(version: number): void;
	setTimeout?(this: void, callback: () => void, delayMs: number): number;
	clearTimeout?(this: void, timer: number): void;
}

interface PendingSave<Source extends string> extends DocumentSaveRequest<Source> {
	resolve: (saved: boolean) => void;
}

/**
 * Format-neutral save lifecycle. It serializes, prepares, validates, and
 * persists one save at a time; concurrent requests collapse to the newest.
 */
export class DocumentSaveCoordinator<Context, Serialized, Prepared, Validated, Source extends string> {
	private readonly setTimer: (callback: () => void, delayMs: number) => number;
	private readonly clearTimer: (timer: number) => void;
	private readonly now: () => number;
	private stateValue: DocumentSaveState = 'clean';
	private dirtyVersion = 0;
	private autosaveTimer: number | null = null;
	private activeSave: Promise<void> | null = null;
	private pendingSave: PendingSave<Source> | null = null;

	constructor(private readonly options: DocumentSaveCoordinatorOptions<Context, Serialized, Prepared, Validated, Source>) {
		this.setTimer = options.setTimeout ?? ((callback, delayMs) => window.setTimeout(callback, delayMs));
		this.clearTimer = options.clearTimeout ?? ((timer) => window.clearTimeout(timer));
		this.now = options.now ?? (() => typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now());
	}

	get state(): DocumentSaveState {
		return this.stateValue;
	}

	get dirty(): boolean {
		return this.stateValue !== 'clean';
	}

	get version(): number {
		return this.dirtyVersion;
	}

	markDirty(): void {
		this.dirtyVersion++;
		this.setState('dirty');
		this.scheduleAutosave();
	}

	clearAutosave(): void {
		if (this.autosaveTimer !== null) {
			this.clearTimer(this.autosaveTimer);
			this.autosaveTimer = null;
		}
	}

	save(source: Source): Promise<boolean> {
		this.clearAutosave();
		const request: DocumentSaveRequest<Source> = { source, targetVersion: this.dirtyVersion };

		return new Promise<boolean>((resolve) => {
			const next: PendingSave<Source> = { ...request, resolve };
			if (this.pendingSave) {
				const previous = this.pendingSave.resolve;
				this.pendingSave = {
					...next,
					resolve: (saved) => {
						previous(saved);
						resolve(saved);
					},
				};
			} else {
				this.pendingSave = next;
			}

			if (!this.activeSave) {
				this.activeSave = this.drainSaves().finally(() => {
					this.activeSave = null;
				});
			}
		});
	}

	async waitForIdle(): Promise<void> {
		while (this.activeSave) {
			await this.activeSave;
		}
	}

	/** Test-only: inject an in-flight save promise for teardown waits. */
	setActiveSaveForTests(promise: Promise<void>): void {
		const tracked = Promise.resolve(promise).finally(() => {
			if (this.activeSave === tracked) this.activeSave = null;
		});
		this.activeSave = tracked;
	}

	reset(): void {
		this.clearAutosave();
		this.dirtyVersion = 0;
		this.pendingSave = null;
		this.setState('clean');
	}

	private async drainSaves(): Promise<void> {
		while (this.pendingSave) {
			const request = this.pendingSave;
			this.pendingSave = null;
			const saved = await this.execute(request);
			request.resolve(saved);
		}
	}

	private async execute(request: DocumentSaveRequest<Source>): Promise<boolean> {
		const context = this.options.getContext();
		if (!context) return false;

		const startedAt = this.now();
		const phases: DocumentSavePhaseMeasurement[] = [];
		let metrics = this.options.getSaveMetrics?.(context) ?? {};
		let outputBytes = metrics.outputBytes ?? null;
		this.options.onSaveStarted?.(request, context, startedAt);
		this.setState('saving');
		const runPhase = async <Result>(name: DocumentSavePhaseName, operation: () => Promise<Result>): Promise<Result> => {
			const invocationStartedAt = this.now();
			const resultPromise = operation();
			const invocationEndedAt = this.now();
			const result = await resultPromise;
			const endedAt = this.now();
			phases.push({
				name,
				durationMs: Math.max(0, endedAt - invocationStartedAt),
				synchronousWorkMs: Math.max(0, invocationEndedAt - invocationStartedAt),
			});
			return result;
		};

		try {
			// Edits can land while serialize/prepare runs. Persist only a snapshot
			// that still matches live dirtyVersion; otherwise retry so disk does
			// not lag behind the editor (and host buffer swaps cannot revive it).
			const maxAttempts = 8;
			let liveRequest: DocumentSaveRequest<Source> = request;
			for (let attempt = 0; attempt < maxAttempts; attempt++) {
				liveRequest = {
					source: request.source,
					targetVersion: this.dirtyVersion,
				};
				const serialized = await runPhase('serialize', () => this.options.adapter.serialize(context, liveRequest));
				metrics = this.options.getSaveMetrics?.(context, serialized) ?? metrics;
				outputBytes = metrics.outputBytes ?? outputBytes;
				if (this.dirtyVersion !== liveRequest.targetVersion) {
					continue;
				}
				const prepared = await runPhase('prepare', () => this.options.adapter.prepareForWrite(serialized, context, liveRequest));
				if (this.dirtyVersion !== liveRequest.targetVersion) {
					continue;
				}
				const validated = await runPhase('validate', () => this.options.adapter.validate(prepared, context, liveRequest));
				if (this.dirtyVersion !== liveRequest.targetVersion) {
					continue;
				}
				await runPhase('persist', () => this.options.adapter.persist(prepared, validated, context, liveRequest));

				if (this.dirtyVersion === liveRequest.targetVersion) {
					this.setState('clean');
				} else {
					this.setState('dirty');
					this.scheduleAutosave();
				}
				return true;
			}

			this.setState('dirty');
			this.scheduleAutosave();
			return false;
		} catch (error) {
			this.setState('failed', error);
			if (this.options.autosave.enabled()) this.scheduleAutosave(5000);
			return false;
		} finally {
			const endedAt = this.now();
			this.options.onSaveCompleted?.({
				source: request.source,
				targetVersion: request.targetVersion,
				startedAt,
				endedAt,
				phases,
				documentBytes: metrics.documentBytes ?? null,
				changedContentBytes: metrics.changedContentBytes ?? null,
				outputBytes,
			});
		}
	}

	private scheduleAutosave(delayMs = this.options.autosave.delayMs()): void {
		this.clearAutosave();
		if (!this.options.autosave.enabled() || this.stateValue === 'clean') return;

		this.options.onAutosaveScheduled?.(delayMs, this.dirtyVersion);
		this.autosaveTimer = this.setTimer(() => {
			this.autosaveTimer = null;
			this.options.onAutosaveStarted?.(this.dirtyVersion);
			if (this.options.runAutosave) {
				this.options.runAutosave(this.dirtyVersion);
			} else {
				void this.save(this.options.autosave.source);
			}
		}, delayMs);
	}

	private setState(state: DocumentSaveState, error?: unknown): void {
		this.stateValue = state;
		this.options.onStateChange?.(state, error);
	}
}
