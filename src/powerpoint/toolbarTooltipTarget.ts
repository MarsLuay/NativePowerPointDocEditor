import { containsEventTarget, isElementLike } from '../ui/tooltipUtils';

export const POWERPOINT_TOOLTIP_TARGET_SELECTOR = [
	'.native-powerpoint-toolbar button',
	'.native-powerpoint-text-toolbar button',
	'.native-powerpoint-find-panel button',
	'.native-powerpoint-rotate-handle',
].join(', ');

const POWERPOINT_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	'.native-powerpoint-toolbar-popover',
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
