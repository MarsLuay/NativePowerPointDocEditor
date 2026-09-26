import { containsEventTarget, isElementLike } from '../ui/tooltipUtils';
import {
	EDITOR_CHROME_REGIONS,
	PPTX_EDITOR_CHROME_TOOLBAR_POPOVER_CLASS,
} from '../editorChromeRegions';

export const POWERPOINT_TOOLTIP_TARGET_SELECTOR = [
	`${EDITOR_CHROME_REGIONS.toolbar.pptx.selector} button`,
	`${EDITOR_CHROME_REGIONS.contextToolbar.pptx.selector} button`,
	'.native-powerpoint-find-panel button',
	'.native-powerpoint-rotate-handle',
].join(', ');

const POWERPOINT_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	`.${PPTX_EDITOR_CHROME_TOOLBAR_POPOVER_CLASS}`,
	'[role="dialog"]',
	'[role="menu"]',
].join(', ');

export function resolvePowerPointTooltipTarget(target: EventTarget | null, root: HTMLElement): HTMLElement | null {
	if (!isElementLike(target)) {
		return null;
	}

	const candidate = target.closest<HTMLElement>(POWERPOINT_TOOLTIP_TARGET_SELECTOR);
	if (!candidate || !containsEventTarget(root, candidate)) {
		return null;
	}

	if (candidate.closest(POWERPOINT_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR)) {
		return null;
	}

	return candidate;
}
