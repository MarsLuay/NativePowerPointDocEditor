import { ControlTooltipAdapter } from '../ui/ControlTooltipAdapter';
import { resolvePowerPointTooltipTarget } from './toolbarTooltipTarget';

/**
 * PowerPoint adapter for shared delayed control tooltips.
 */
export class ToolbarTooltipController extends ControlTooltipAdapter {
	protected resolveTarget(target: EventTarget | null, root: HTMLElement): HTMLElement | null {
		return resolvePowerPointTooltipTarget(target, root);
	}
}
