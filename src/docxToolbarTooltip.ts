import { EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE } from './docxEditorChromeMarkers';
import { TooltipController } from './ui/TooltipController';
import {
	containsEventTarget,
	getToolbarTooltipText,
	isElementLike,
	normalizeTooltipText,
	positionToolbarTooltip,
	restoreNativeTitle,
	suspendNativeTitle,
	TOOLBAR_TOOLTIP_CLASS,
	TOOLBAR_TOOLTIP_DELAY_MS,
} from './ui/tooltipUtils';

export {
	getToolbarTooltipText,
	positionToolbarTooltip,
	restoreNativeTitle,
	suspendNativeTitle,
	TOOLBAR_TOOLTIP_DELAY_MS,
};

export const DOCX_EDITOR_ROOT_CLASS_PREFIX = 'native-powerpoint-doc-editor-editor-';
export const DOCX_TOOLBAR_TOOLTIP_CLASS = TOOLBAR_TOOLTIP_CLASS;
export const EIGENPAL_TOOLTIP_ATTRIBUTE = 'data-native-powerpoint-doc-editor-eigenpal-tooltip';
export const EIGENPAL_TOOLBAR_TOOLTIP_SELECTOR = [
	`.ep-root [${EIGENPAL_TOOLTIP_ATTRIBUTE}]`,
	'.ep-root [role="tooltip"]',
	'.ep-root [data-radix-tooltip-content]',
	'.ep-root .fixed.z-50.px-2.py-1.rounded-md.shadow-lg:not([role])',
].join(', ');

const DOCX_FORMATTING_TOOLBAR_SELECTOR = '[data-testid="formatting-bar"]';

export const DOCX_TOOLTIP_TOOLBAR_SELECTOR = [
	'[data-testid="editor-toolbar"]',
	'.docx-table-toolbar',
].join(', ');

const DOCX_TOOLTIP_BUTTON_SELECTOR = [
	'[data-testid="editor-toolbar"] button',
	'.docx-table-toolbar button',
].join(', ');

const TOOLBAR_BUTTON_SELECTOR = [
	'.ep-root [data-testid="editor-toolbar"] button',
	'.ep-root .docx-table-toolbar button',
].join(', ');

const DOCUMENT_LINK_TITLE_CONTAINER_SELECTOR = [
	'.layout-page',
	'.layout-page-content',
	'.ProseMirror',
	'.ep-hyperlink-popup',
].join(', ');

const TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	`[${EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE}]`,
	'[data-testid="title-bar"]',
	DOCX_FORMATTING_TOOLBAR_SELECTOR,
	'[role="menubar"]',
	'[role="menu"]',
	'[role="dialog"]',
	'.ep-hyperlink-popup',
	'[data-radix-popper-content-wrapper]',
].join(', ');

const EIGENPAL_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	'.ep-hyperlink-popup',
	'[role="dialog"]',
	'[role="menu"]',
	'[role="listbox"]',
	'[data-radix-select-content]',
	'[data-radix-popper-content-wrapper]',
].join(', ');

function getToolbarButtonMetadata(button: HTMLElement): string {
	return normalizeTooltipText(button.dataset.tooltip)
		|| normalizeTooltipText(button.getAttribute('aria-label'))
		|| normalizeTooltipText(button.dataset.nativePowerPointDocEditorNativeTitle)
		|| normalizeTooltipText(button.getAttribute('title'));
}

function standardizeToolbarButtonTooltipMetadata(button: HTMLElement): void {
	const label = getToolbarButtonMetadata(button);
	if (!label) {
		return;
	}

	if (!normalizeTooltipText(button.dataset.tooltip)) {
		button.dataset.tooltip = label;
	}
	if (!normalizeTooltipText(button.getAttribute('aria-label'))) {
		button.setAttribute('aria-label', label);
	}
}

export function neutralizeToolbarButtonTooltipSources(host: ParentNode): void {
	host.querySelectorAll<HTMLButtonElement>(TOOLBAR_BUTTON_SELECTOR).forEach((button) => {
		if (button.closest(`[${EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE}]`)) {
			return;
		}
		standardizeToolbarButtonTooltipMetadata(button);
	});
}

export function stripFormattingDropdownButtonTitles(formattingBar: HTMLElement): void {
	formattingBar.querySelectorAll<HTMLElement>('button[title]').forEach((button) => {
		const title = button.getAttribute('title');
		if (title !== null && button.dataset.nativePowerPointDocEditorNativeTitle === undefined) {
			button.dataset.nativePowerPointDocEditorNativeTitle = title;
		}
		delete button.dataset.tooltip;
		button.removeAttribute('title');
	});
}

export function getDocxEditorRootFromTarget(target: EventTarget | null): HTMLElement | null {
	if (!isElementLike(target)) {
		return null;
	}

	return target.closest<HTMLElement>(`[class*="${DOCX_EDITOR_ROOT_CLASS_PREFIX}"]`);
}

export function getDocumentLinkTitleTarget(
	target: EventTarget | null,
	editorRoot: HTMLElement | null,
): HTMLAnchorElement | null {
	if (!isElementLike(target) || !editorRoot) {
		return null;
	}

	const link = target.closest<HTMLAnchorElement>('a[title]');
	if (!link || !containsEventTarget(editorRoot, link)) {
		return null;
	}

	return link.closest(DOCUMENT_LINK_TITLE_CONTAINER_SELECTOR) ? link : null;
}

export function getToolbarTooltipTarget(target: EventTarget | null, editorRoot: HTMLElement | null): HTMLElement | null {
	if (!isElementLike(target) || !editorRoot) {
		return null;
	}

	const candidate = target.closest<HTMLElement>(DOCX_TOOLTIP_BUTTON_SELECTOR);
	if (!candidate || !containsEventTarget(editorRoot, candidate)) {
		return null;
	}

	if (candidate.closest(DOCX_FORMATTING_TOOLBAR_SELECTOR)) {
		return null;
	}

	const excludedAncestor = candidate.closest(TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR);
	if (excludedAncestor && containsEventTarget(editorRoot, excludedAncestor)) {
		return null;
	}

	const toolbar = candidate.closest(DOCX_TOOLTIP_TOOLBAR_SELECTOR);
	if (!toolbar || !containsEventTarget(editorRoot, toolbar)) {
		return null;
	}

	if (candidate.getAttribute('aria-expanded') === 'true') {
		return null;
	}

	return candidate;
}

export function resolveDocxTooltipTarget(
	target: EventTarget | null,
	editorRoot: HTMLElement,
): HTMLElement | null {
	return getToolbarTooltipTarget(target, editorRoot);
}

function isEigenpalToolbarTooltip(element: Element): element is HTMLElement {
	return isElementLike(element)
		&& !element.classList.contains(DOCX_TOOLBAR_TOOLTIP_CLASS)
		&& !element.closest(EIGENPAL_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR)
		&& element.matches(EIGENPAL_TOOLBAR_TOOLTIP_SELECTOR);
}

export function suppressEigenpalToolbarTooltips(editorRoot: HTMLElement | null): void {
	if (!editorRoot) {
		return;
	}

	editorRoot.querySelectorAll<HTMLElement>(EIGENPAL_TOOLBAR_TOOLTIP_SELECTOR).forEach((element) => {
		if (!isEigenpalToolbarTooltip(element)) {
			return;
		}

		if (element.getAttribute(EIGENPAL_TOOLTIP_ATTRIBUTE) === 'true' && element.hidden) {
			return;
		}

		element.setAttribute(EIGENPAL_TOOLTIP_ATTRIBUTE, 'true');
		element.hidden = true;
	});
}

const tooltipManagers = new WeakMap<HTMLElement, () => void>();

function installDocxToolbarTooltipManager(editorRoot: HTMLElement): () => void {
	let activeDocumentLink: HTMLAnchorElement | null = null;
	let eigenpalSuppressionTimer: number | null = null;
	const view = editorRoot.ownerDocument.defaultView ?? window;

	const clearDocumentLinkTitle = (): void => {
		restoreNativeTitle(activeDocumentLink);
		activeDocumentLink = null;
	};

	const scheduleEigenpalTooltipSuppression = (): void => {
		if (eigenpalSuppressionTimer !== null) {
			view.clearTimeout(eigenpalSuppressionTimer);
		}
		eigenpalSuppressionTimer = view.setTimeout(() => {
			eigenpalSuppressionTimer = null;
			suppressEigenpalToolbarTooltips(editorRoot);
		}, 0);
	};

	const tooltipController = new TooltipController({
		root: editorRoot,
		getTarget: (target) => resolveDocxTooltipTarget(target, editorRoot),
		onTargetRecognized: () => suppressEigenpalToolbarTooltips(editorRoot),
	});

	const handlePointerOver = (evt: PointerEvent): void => {
		const link = getDocumentLinkTitleTarget(evt.target, editorRoot);
		if (link && link !== activeDocumentLink) {
			clearDocumentLinkTitle();
			activeDocumentLink = link;
			suspendNativeTitle(link);
		}
	};

	const handlePointerOut = (evt: PointerEvent): void => {
		if (
			activeDocumentLink
			&& !containsEventTarget(activeDocumentLink, evt.relatedTarget)
		) {
			clearDocumentLinkTitle();
		}
	};

	const handleBlur = (): void => {
		clearDocumentLinkTitle();
	};

	editorRoot.addEventListener('pointerover', handlePointerOver, true);
	editorRoot.addEventListener('pointerout', handlePointerOut, true);
	view.addEventListener('blur', handleBlur, false);

	const eigenpalTooltipObserver = new MutationObserver(scheduleEigenpalTooltipSuppression);
	eigenpalTooltipObserver.observe(editorRoot, { childList: true, subtree: true });
	scheduleEigenpalTooltipSuppression();

	return () => {
		eigenpalTooltipObserver.disconnect();
		if (eigenpalSuppressionTimer !== null) {
			view.clearTimeout(eigenpalSuppressionTimer);
			eigenpalSuppressionTimer = null;
		}
		editorRoot.removeEventListener('pointerover', handlePointerOver, true);
		editorRoot.removeEventListener('pointerout', handlePointerOut, true);
		view.removeEventListener('blur', handleBlur, false);
		tooltipController.detach();
		clearDocumentLinkTitle();
	};
}

export function attachDocxToolbarTooltipManager(editorRoot: HTMLElement): () => void {
	tooltipManagers.get(editorRoot)?.();

	const detachManager = installDocxToolbarTooltipManager(editorRoot);
	let detached = false;
	const detach = (): void => {
		if (detached) {
			return;
		}

		detached = true;
		detachManager();
	};

	tooltipManagers.set(editorRoot, detach);
	return () => {
		if (tooltipManagers.get(editorRoot) === detach) {
			tooltipManagers.delete(editorRoot);
		}
		detach();
	};
}
