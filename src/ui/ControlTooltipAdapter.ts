import { Component } from 'obsidian';

import { TooltipController } from './TooltipController';

/**
 * Shared lifecycle for delayed control tooltips. Each surface (DOCX editor
 * chrome, PowerPoint toolbars) supplies a target resolver plus optional
 * surface-specific hooks; the shared {@link TooltipController} owns hover
 * scheduling, positioning, and accessibility wiring.
 *
 * Subclasses implement {@link resolveTarget} and may override
 * {@link onTargetRecognized} (per-hover side effects) and {@link onAttach}
 * (extra listeners/observers, returning their own cleanup).
 */
export abstract class ControlTooltipAdapter extends Component {
	private tooltipController: TooltipController | null = null;
	private surfaceCleanup: (() => void) | null = null;
	private lifecycleRegistered = false;

	attach(root: HTMLElement): void {
		this.teardown();
		this.surfaceCleanup = this.onAttach(root) ?? null;
		this.tooltipController = new TooltipController({
			root,
			getTarget: (target) => this.resolveTarget(target, root),
			onTargetRecognized: (target) => this.onTargetRecognized(target, root),
		});

		if (!this.lifecycleRegistered) {
			this.lifecycleRegistered = true;
			this.register(() => this.teardown());
		}
	}

	private teardown(): void {
		this.tooltipController?.detach();
		this.tooltipController = null;
		this.surfaceCleanup?.();
		this.surfaceCleanup = null;
	}

	protected abstract resolveTarget(target: EventTarget | null, root: HTMLElement): HTMLElement | null;

	protected onTargetRecognized(_target: HTMLElement, _root: HTMLElement): void {}

	protected onAttach(_root: HTMLElement): (() => void) | void {}
}
