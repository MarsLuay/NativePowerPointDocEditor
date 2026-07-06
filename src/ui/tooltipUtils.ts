export const TOOLBAR_TOOLTIP_CLASS = 'native-powerpoint-doc-editor-toolbar-tooltip';
export const TOOLBAR_TOOLTIP_DELAY_MS = 450;
export const TOOLBAR_TOOLTIP_GAP_PX = 8;
export const TOOLBAR_TOOLTIP_VIEWPORT_PADDING_PX = 8;
export const SUSPENDED_TITLE_DATA_KEY = 'nativePowerPointDocEditorTooltipTitle';

export interface TooltipPositionOptions {
	gapPx?: number;
	viewportPaddingPx?: number;
	viewportWidth?: number;
}

export function normalizeTooltipText(value: string | null | undefined): string {
	return (value ?? '').replace(/\s+/g, ' ').trim();
}

export function getToolbarTooltipText(target: HTMLElement): string {
	const explicitTooltip = normalizeTooltipText(target.dataset.tooltip ?? target.getAttribute('data-tooltip'));
	if (explicitTooltip) {
		return explicitTooltip;
	}

	const ariaLabel = normalizeTooltipText(target.getAttribute('aria-label'));
	if (ariaLabel) {
		return ariaLabel;
	}

	return normalizeTooltipText(target.getAttribute('title') ?? target.dataset[SUSPENDED_TITLE_DATA_KEY]);
}

export function suspendNativeTitle(target: HTMLElement | null): void {
	if (!target) {
		return;
	}

	const title = target.getAttribute('title');
	if (title === null) {
		return;
	}

	if (target.dataset[SUSPENDED_TITLE_DATA_KEY] === undefined) {
		target.dataset[SUSPENDED_TITLE_DATA_KEY] = title;
	}
	target.removeAttribute('title');
}

export function restoreNativeTitle(target: HTMLElement | null): void {
	if (!target) {
		return;
	}

	const title = target.dataset[SUSPENDED_TITLE_DATA_KEY];
	if (title !== undefined) {
		target.setAttribute('title', title);
	}
	delete target.dataset[SUSPENDED_TITLE_DATA_KEY];
}

export function isElementLike(value: unknown): value is HTMLElement {
	return (
		(typeof value === 'object' || typeof value === 'function')
		&& value !== null
		&& typeof (value as Element).closest === 'function'
		&& typeof (value as Element).matches === 'function'
	);
}

export function containsEventTarget(container: HTMLElement | Document, target: EventTarget | null): boolean {
	return Boolean(
		target
		&& typeof (container as ParentNode).contains === 'function'
		&& (container as ParentNode).contains(target as Node),
	);
}

export function positionToolbarTooltip(
	target: HTMLElement,
	tooltip: HTMLDivElement,
	options: TooltipPositionOptions = {},
): void {
	const gapPx = options.gapPx ?? TOOLBAR_TOOLTIP_GAP_PX;
	const viewportPaddingPx = options.viewportPaddingPx ?? TOOLBAR_TOOLTIP_VIEWPORT_PADDING_PX;
	const viewportWidth = options.viewportWidth
		?? target.ownerDocument.defaultView?.innerWidth
		?? window.innerWidth;

	const rect = target.getBoundingClientRect();
	tooltip.style.setProperty('--native-powerpoint-doc-editor-toolbar-tooltip-left', `${Math.round(rect.left + rect.width / 2)}px`);
	tooltip.style.setProperty('--native-powerpoint-doc-editor-toolbar-tooltip-top', `${Math.round(rect.bottom + gapPx)}px`);
	tooltip.classList.remove('is-left-aligned', 'is-right-aligned');

	const tooltipRect = tooltip.getBoundingClientRect();
	if (tooltipRect.left < viewportPaddingPx) {
		tooltip.style.setProperty('--native-powerpoint-doc-editor-toolbar-tooltip-left', `${viewportPaddingPx}px`);
		tooltip.classList.add('is-left-aligned');
	} else if (tooltipRect.right > viewportWidth - viewportPaddingPx) {
		tooltip.style.setProperty('--native-powerpoint-doc-editor-toolbar-tooltip-left', `${viewportWidth - viewportPaddingPx}px`);
		tooltip.classList.add('is-right-aligned');
	}
}
