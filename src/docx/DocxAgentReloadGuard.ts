export interface DocxReloadIdentity {
	documentSession: number;
	filePath: string;
}

interface AgentReloadBuffer extends DocxReloadIdentity {
	version: number;
	buffer: ArrayBuffer;
}

interface ReadyWaiter {
	identity: DocxReloadIdentity;
	promise: Promise<void>;
	resolve: () => void;
	reject: (reason: Error) => void;
	timeout: number;
}

function matchesIdentity(left: DocxReloadIdentity | null, right: DocxReloadIdentity): boolean {
	return Boolean(
		left
		&& left.documentSession === right.documentSession
		&& left.filePath === right.filePath,
	);
}

function reloadError(message: string): Error {
	return new Error(`DOCX agent reload ${message}`);
}

/**
 * Keeps an agent-patched package authoritative while React swaps the live
 * editor to a new document epoch. The epoch and file path are both required:
 * a stale editor for another file must never receive or persist this package.
 */
export class DocxAgentReloadGuard {
	private version = 0;
	private pending: AgentReloadBuffer | null = null;
	private latest: AgentReloadBuffer | null = null;
	private completed: DocxReloadIdentity | null = null;
	private waiter: ReadyWaiter | null = null;

	begin(identity: DocxReloadIdentity, buffer: ArrayBuffer): void {
		this.rejectWaiter(reloadError('was superseded by a newer reload.'));
		this.completed = null;
		const entry: AgentReloadBuffer = {
			...identity,
			version: ++this.version,
			buffer: buffer.slice(0),
		};
		this.pending = entry;
		this.latest = entry;
	}

	/** Replaces a staged package after asynchronous preprocessing completes. */
	stage(identity: DocxReloadIdentity, buffer: ArrayBuffer): boolean {
		if (!matchesIdentity(this.pending, identity)) {
			return false;
		}

		const entry: AgentReloadBuffer = {
			...identity,
			version: ++this.version,
			buffer: buffer.slice(0),
		};
		this.pending = entry;
		this.latest = entry;
		return true;
	}

	getVersion(): number {
		return this.version;
	}

	getPendingBuffer(identity: DocxReloadIdentity): ArrayBuffer | null {
		const pending = this.pending;
		return matchesIdentity(pending, identity) && pending
			? pending.buffer.slice(0)
			: null;
	}

	getLatestBufferAfter(
		version: number,
		identity: DocxReloadIdentity,
	): { version: number; buffer: ArrayBuffer } | null {
		if (!matchesIdentity(this.latest, identity) || !this.latest || this.latest.version <= version) {
			return null;
		}

		return {
			version: this.latest.version,
			buffer: this.latest.buffer.slice(0),
		};
	}

	async waitForReady(identity: DocxReloadIdentity, timeoutMs: number): Promise<void> {
		if (matchesIdentity(this.completed, identity)) {
			return;
		}
		if (!matchesIdentity(this.pending, identity)) {
			throw reloadError('is no longer current.');
		}
		if (this.waiter && matchesIdentity(this.waiter.identity, identity)) {
			return this.waiter.promise;
		}

		let resolve!: () => void;
		let reject!: (reason: Error) => void;
		const promise = new Promise<void>((resolvePromise, rejectPromise) => {
			resolve = resolvePromise;
			reject = rejectPromise;
		});
		const timeout = window.setTimeout(() => {
			this.fail(identity, reloadError(`did not become ready within ${timeoutMs}ms.`));
		}, timeoutMs);
		this.waiter = { identity: { ...identity }, promise, resolve, reject, timeout };
		return promise;
	}

	complete(identity: DocxReloadIdentity): boolean {
		if (!matchesIdentity(this.pending, identity)) {
			return false;
		}

		this.pending = null;
		this.completed = { ...identity };
		this.resolveWaiter(identity);
		return true;
	}

	/** Rejects only the specified current reload without disturbing a successor. */
	fail(identity: DocxReloadIdentity, reason: Error): boolean {
		const isCurrent = matchesIdentity(this.pending, identity) || matchesIdentity(this.completed, identity);
		if (!isCurrent) {
			return false;
		}

		this.pending = null;
		this.latest = null;
		this.completed = null;
		this.rejectWaiter(reason, identity);
		return true;
	}

	clear(reason = reloadError('was canceled because the document lifecycle changed.')): void {
		this.version += 1;
		this.pending = null;
		this.latest = null;
		this.completed = null;
		this.rejectWaiter(reason);
	}

	private resolveWaiter(identity: DocxReloadIdentity): void {
		if (!this.waiter || !matchesIdentity(this.waiter.identity, identity)) {
			return;
		}
		window.clearTimeout(this.waiter.timeout);
		const waiter = this.waiter;
		this.waiter = null;
		waiter.resolve();
	}

	private rejectWaiter(reason: Error, identity?: DocxReloadIdentity): void {
		if (!this.waiter || (identity && !matchesIdentity(this.waiter.identity, identity))) {
			return;
		}
		window.clearTimeout(this.waiter.timeout);
		const waiter = this.waiter;
		this.waiter = null;
		waiter.reject(reason);
	}
}
