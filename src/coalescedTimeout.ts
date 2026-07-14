interface TimeoutView {
	setTimeout(handler: TimerHandler, timeout?: number): number;
	clearTimeout(id: number): void;
}

export class CoalescedTimeout {
	private timeoutId: number | null = null;

	constructor(
		private readonly view: TimeoutView,
		private readonly task: () => void,
	) {}

	schedule(delay: number): void {
		if (this.timeoutId !== null) return;
		this.timeoutId = this.view.setTimeout(() => {
			this.timeoutId = null;
			this.task();
		}, delay);
	}

	cancel(): void {
		if (this.timeoutId === null) return;
		this.view.clearTimeout(this.timeoutId);
		this.timeoutId = null;
	}
}
