import {
	DOCX_EDITOR_FORMATTING_BAR_SELECTOR,
	DOCX_EDITOR_MENUBAR_SELECTOR,
	DOCX_EDITOR_ROOT_SELECTOR,
	DOCX_EDITOR_TITLE_BAR_SELECTOR,
	DOCX_EDITOR_TOOLBAR_SELECTOR,
	DOCX_VENDOR_TOOLTIP_ATTRIBUTE,
	DOCX_VENDOR_TOOLTIP_SELECTOR,
	DOCX_HYPERLINK_POPUP_SELECTOR,
	DOCX_RENDERED_PAGE_CONTENT_SELECTOR,
	DOCX_RENDERED_PAGE_SELECTOR,
	DOCX_TABLE_TOOLBAR_SELECTOR,
	EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE,
} from './docxEditorChromeMarkers';
import { ControlTooltipAdapter } from './ui/ControlTooltipAdapter';
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
export const VENDOR_TOOLTIP_ATTRIBUTE = DOCX_VENDOR_TOOLTIP_ATTRIBUTE;
export const VENDOR_TOOLBAR_TOOLTIP_SELECTOR = DOCX_VENDOR_TOOLTIP_SELECTOR;

const DOCX_FORMATTING_TOOLBAR_SELECTOR = DOCX_EDITOR_FORMATTING_BAR_SELECTOR;

export const DOCX_TOOLTIP_TOOLBAR_SELECTOR = [
	DOCX_EDITOR_TOOLBAR_SELECTOR,
	DOCX_TABLE_TOOLBAR_SELECTOR,
].join(', ');

const DOCX_TOOLTIP_BUTTON_SELECTOR = [
	`${DOCX_EDITOR_TOOLBAR_SELECTOR} button`,
	`${DOCX_TABLE_TOOLBAR_SELECTOR} button`,
].join(', ');

const TOOLBAR_BUTTON_SELECTOR = [
	`${DOCX_EDITOR_ROOT_SELECTOR} ${DOCX_EDITOR_TOOLBAR_SELECTOR} button`,
	`${DOCX_EDITOR_ROOT_SELECTOR} ${DOCX_TABLE_TOOLBAR_SELECTOR} button`,
].join(', ');

const DOCUMENT_LINK_TITLE_CONTAINER_SELECTOR = [
	DOCX_RENDERED_PAGE_SELECTOR,
	DOCX_RENDERED_PAGE_CONTENT_SELECTOR,
	'.ProseMirror',
	DOCX_HYPERLINK_POPUP_SELECTOR,
].join(', ');

const TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	`[${EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE}]`,
	DOCX_EDITOR_TITLE_BAR_SELECTOR,
	DOCX_FORMATTING_TOOLBAR_SELECTOR,
	DOCX_EDITOR_MENUBAR_SELECTOR,
	'[role="menu"]',
	'[role="dialog"]',
	DOCX_HYPERLINK_POPUP_SELECTOR,
].join(', ');

const VENDOR_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR = [
	DOCX_HYPERLINK_POPUP_SELECTOR,
	'[role="dialog"]',
	'[role="menu"]',
	'[role="listbox"]',
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

function isVendorToolbarTooltip(element: Element): element is HTMLElement {
	return isElementLike(element)
		&& !element.classList.contains(DOCX_TOOLBAR_TOOLTIP_CLASS)
		&& !element.closest(VENDOR_TOOLTIP_EXCLUDED_ANCESTOR_SELECTOR)
		&& element.matches(VENDOR_TOOLBAR_TOOLTIP_SELECTOR);
}

export function suppressVendorToolbarTooltips(editorRoot: HTMLElement | null): void {
	if (!editorRoot) {
		return;
	}

	editorRoot.querySelectorAll<HTMLElement>(VENDOR_TOOLBAR_TOOLTIP_SELECTOR).forEach((element) => {
		if (!isVendorToolbarTooltip(element)) {
			return;
		}

		if (element.getAttribute(VENDOR_TOOLTIP_ATTRIBUTE) === 'true' && element.hidden) {
			return;
		}

		element.setAttribute(VENDOR_TOOLTIP_ATTRIBUTE, 'true');
		element.hidden = true;
	});
}

/**
 * DOCX adapter for shared delayed control tooltips. Beyond the shared hover
 * lifecycle it suspends native `title` tooltips on document hyperlinks and
 * keeps the vendored toolbar tooltips suppressed as the editor DOM
 * mutates.
 */
class DocxToolbarTooltipController extends ControlTooltipAdapter {
	protected resolveTarget(target: EventTarget | null, root: HTMLElement): HTMLElement | null {
		return resolveDocxTooltipTarget(target, root);
	}

	protected onTargetRecognized(_target: HTMLElement, root: HTMLElement): void {
		suppressVendorToolbarTooltips(root);
	}

	protected onAttach(editorRoot: HTMLElement): () => void {
		let activeDocumentLink: HTMLAnchorElement | null = null;
		let vendorSuppressionTimer: number | null = null;
		const view = editorRoot.ownerDocument.defaultView ?? window;

		const clearDocumentLinkTitle = (): void => {
			restoreNativeTitle(activeDocumentLink);
			activeDocumentLink = null;
		};

		const scheduleVendorTooltipSuppression = (): void => {
			if (vendorSuppressionTimer !== null) {
				view.clearTimeout(vendorSuppressionTimer);
			}
			vendorSuppressionTimer = view.setTimeout(() => {
				vendorSuppressionTimer = null;
				suppressVendorToolbarTooltips(editorRoot);
			}, 0);
		};

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

		const vendorTooltipObserver = new MutationObserver(scheduleVendorTooltipSuppression);
		vendorTooltipObserver.observe(editorRoot, { childList: true, subtree: true });
		scheduleVendorTooltipSuppression();

		return () => {
			vendorTooltipObserver.disconnect();
			if (vendorSuppressionTimer !== null) {
				view.clearTimeout(vendorSuppressionTimer);
				vendorSuppressionTimer = null;
			}
			editorRoot.removeEventListener('pointerover', handlePointerOver, true);
			editorRoot.removeEventListener('pointerout', handlePointerOut, true);
			view.removeEventListener('blur', handleBlur, false);
			clearDocumentLinkTitle();
		};
	}
}

const tooltipManagers = new WeakMap<HTMLElement, DocxToolbarTooltipController>();

export function attachDocxToolbarTooltipManager(editorRoot: HTMLElement): () => void {
	tooltipManagers.get(editorRoot)?.unload();

	const controller = new DocxToolbarTooltipController();
	controller.load();
	controller.attach(editorRoot);
	tooltipManagers.set(editorRoot, controller);

	return () => {
		if (tooltipManagers.get(editorRoot) === controller) {
			tooltipManagers.delete(editorRoot);
		}
		controller.unload();
	};
}
