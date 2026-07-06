import { Component } from 'obsidian';

import { TooltipController } from '../ui/TooltipController';
import { resolvePowerPointTooltipTarget } from './toolbarTooltipTarget';

/**
 * PowerPoint adapter for shared delayed control tooltips.
 */
export class ToolbarTooltipController extends Component {
	private tooltipController: TooltipController | null = null;
	private cleanupRegistered = false;

	attach(root: HTMLElement): void {
		this.tooltipController?.detach();
		this.tooltipController = new TooltipController({
			root,
			getTarget: (target) => resolvePowerPointTooltipTarget(target, root),
		});

		if (!this.cleanupRegistered) {
			this.cleanupRegistered = true;
			this.register(() => {
				this.tooltipController?.detach();
				this.tooltipController = null;
			});
		}
	}
}
