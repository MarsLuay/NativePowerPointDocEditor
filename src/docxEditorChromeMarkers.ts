export type DocxEditorChromeMenuItemKind = 'edit' | 'search' | 'settings';

const MENU_ITEM_CLASSES: Record<DocxEditorChromeMenuItemKind, string> = {
	edit: 'native-powerpoint-doc-editor-edit-menu-item',
	search: 'native-powerpoint-doc-editor-search-menu-item',
	settings: 'native-powerpoint-doc-editor-settings-menu-item',
};

export const EDITOR_CHROME_MENU_ITEMS = Object.fromEntries(
	Object.entries(MENU_ITEM_CLASSES).map(([kind, className]) => {
		const attribute = `data-native-powerpoint-doc-editor-${kind}-menu-item`;
		return [kind, {
			attribute,
			className,
			selector: `[${attribute}], .${className}`,
		}];
	}),
) as Record<DocxEditorChromeMenuItemKind, {
	attribute: string;
	className: string;
	selector: string;
}>;

export const EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE = 'data-native-powerpoint-doc-editor-no-toolbar-tooltip';
export const DOCX_EDITOR_ROOT_ATTRIBUTE = 'data-native-powerpoint-doc-editor-root';
export const DOCX_EDITOR_ROOT_SELECTOR = `[${DOCX_EDITOR_ROOT_ATTRIBUTE}]`;
export const DOCX_EDITOR_TITLE_BAR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-title-bar';
export const DOCX_EDITOR_TITLE_BAR_SELECTOR = `[${DOCX_EDITOR_TITLE_BAR_ATTRIBUTE}]`;
export const DOCX_EDITOR_TOOLBAR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-toolbar';
export const DOCX_EDITOR_TOOLBAR_SELECTOR = `[${DOCX_EDITOR_TOOLBAR_ATTRIBUTE}]`;
export const DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-formatting-bar';
export const DOCX_EDITOR_FORMATTING_BAR_SELECTOR = `[${DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENUBAR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menubar';
export const DOCX_EDITOR_MENUBAR_SELECTOR = `[${DOCX_EDITOR_MENUBAR_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENU_ROOT_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menu-root';
export const DOCX_EDITOR_MENU_ROOT_SELECTOR = `[${DOCX_EDITOR_MENU_ROOT_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menu-button';
export const DOCX_EDITOR_MENU_BUTTON_SELECTOR = `[${DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENU_DROPDOWN_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menu-dropdown';
export const DOCX_EDITOR_MENU_DROPDOWN_SELECTOR = `[${DOCX_EDITOR_MENU_DROPDOWN_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENU_ITEM_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menu-item';
export const DOCX_EDITOR_MENU_ITEM_SELECTOR = `[${DOCX_EDITOR_MENU_ITEM_ATTRIBUTE}]`;
export const DOCX_EDITOR_MENU_ITEM_BUTTON_ATTRIBUTE = 'data-native-powerpoint-doc-editor-menu-item-button';
export const DOCX_EDITOR_MENU_ITEM_BUTTON_SELECTOR = `[${DOCX_EDITOR_MENU_ITEM_BUTTON_ATTRIBUTE}]`;
export const DOCX_EDITOR_SCROLL_CONTAINER_ATTRIBUTE = 'data-native-powerpoint-doc-editor-scroll-container';
export const DOCX_EDITOR_SCROLL_CONTAINER_SELECTOR = `[${DOCX_EDITOR_SCROLL_CONTAINER_ATTRIBUTE}]`;
export const DOCX_EDITOR_PAGES_ATTRIBUTE = 'data-native-powerpoint-doc-editor-pages';
export const DOCX_EDITOR_PAGES_SELECTOR = `[${DOCX_EDITOR_PAGES_ATTRIBUTE}]`;
export const DOCX_RENDERED_PAGE_ATTRIBUTE = 'data-native-powerpoint-doc-editor-page';
export const DOCX_RENDERED_PAGE_SELECTOR = `[${DOCX_RENDERED_PAGE_ATTRIBUTE}]`;
export const DOCX_RENDERED_PAGE_CONTENT_ATTRIBUTE = 'data-native-powerpoint-doc-editor-page-content';
export const DOCX_RENDERED_PAGE_CONTENT_SELECTOR = `[${DOCX_RENDERED_PAGE_CONTENT_ATTRIBUTE}]`;
export const DOCX_RENDERED_PARAGRAPH_ATTRIBUTE = 'data-native-powerpoint-doc-editor-paragraph';
export const DOCX_RENDERED_PARAGRAPH_SELECTOR = `[${DOCX_RENDERED_PARAGRAPH_ATTRIBUTE}]`;
export const DOCX_RENDERED_LIST_MARKER_ATTRIBUTE = 'data-native-powerpoint-doc-editor-list-marker';
export const DOCX_RENDERED_LIST_MARKER_SELECTOR = `[${DOCX_RENDERED_LIST_MARKER_ATTRIBUTE}]`;
export const DOCX_RENDERED_RUN_TAB_ATTRIBUTE = 'data-native-powerpoint-doc-editor-run-tab';
export const DOCX_RENDERED_RUN_TAB_SELECTOR = `[${DOCX_RENDERED_RUN_TAB_ATTRIBUTE}]`;
export const DOCX_HIDDEN_PROSEMIRROR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-hidden-prosemirror';
export const DOCX_HIDDEN_PROSEMIRROR_SELECTOR = `[${DOCX_HIDDEN_PROSEMIRROR_ATTRIBUTE}]`;
export const DOCX_CARET_ATTRIBUTE = 'data-native-powerpoint-doc-editor-caret';
export const DOCX_CARET_SELECTOR = `[${DOCX_CARET_ATTRIBUTE}]`;
export const DOCX_FONT_SIZE_DECREASE_ATTRIBUTE = 'data-native-powerpoint-doc-editor-font-size-decrease';
export const DOCX_FONT_SIZE_DECREASE_SELECTOR = `[${DOCX_FONT_SIZE_DECREASE_ATTRIBUTE}]`;
export const DOCX_FONT_SIZE_INCREASE_ATTRIBUTE = 'data-native-powerpoint-doc-editor-font-size-increase';
export const DOCX_FONT_SIZE_INCREASE_SELECTOR = `[${DOCX_FONT_SIZE_INCREASE_ATTRIBUTE}]`;
export const DOCX_FONT_SIZE_INPUT_ATTRIBUTE = 'data-native-powerpoint-doc-editor-font-size-input';
export const DOCX_FONT_SIZE_INPUT_SELECTOR = `[${DOCX_FONT_SIZE_INPUT_ATTRIBUTE}]`;
export const DOCX_FONT_SIZE_DISPLAY_ATTRIBUTE = 'data-native-powerpoint-doc-editor-font-size-display';
export const DOCX_FONT_SIZE_DISPLAY_SELECTOR = `[${DOCX_FONT_SIZE_DISPLAY_ATTRIBUTE}]`;
export const DOCX_TABLE_TOOLBAR_ATTRIBUTE = 'data-native-powerpoint-doc-editor-table-toolbar';
export const DOCX_TABLE_TOOLBAR_SELECTOR = `[${DOCX_TABLE_TOOLBAR_ATTRIBUTE}]`;
export const DOCX_HYPERLINK_POPUP_ATTRIBUTE = 'data-native-powerpoint-doc-editor-hyperlink-popup';
export const DOCX_HYPERLINK_POPUP_SELECTOR = `[${DOCX_HYPERLINK_POPUP_ATTRIBUTE}]`;
export const DOCX_VENDOR_TOOLTIP_ATTRIBUTE = 'data-native-powerpoint-doc-editor-vendor-tooltip';
export const DOCX_VENDOR_TOOLTIP_SELECTOR = `[${DOCX_VENDOR_TOOLTIP_ATTRIBUTE}]`;

type AttributeTarget = Pick<Element, 'setAttribute'>;

function isElementLike(node: unknown): node is Element {
	return Boolean(
		node
		&& typeof node === 'object'
		&& 'setAttribute' in node
		&& typeof (node as Element).setAttribute === 'function'
		&& 'querySelector' in node
		&& 'children' in node,
	);
}

export function markEditorChromeMenuItem(
	element: AttributeTarget,
	kind: DocxEditorChromeMenuItemKind,
) {
	element.setAttribute(EDITOR_CHROME_MENU_ITEMS[kind].attribute, 'true');
	element.setAttribute(EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE, 'true');
}

export function markEditorChromeMenuButton(element: AttributeTarget) {
	element.setAttribute(DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE, 'true');
}

export function markEditorChromeNoToolbarTooltip(element: AttributeTarget) {
	element.setAttribute(EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE, 'true');
}

export const DOCX_VENDOR_TABLE_TOOLBAR_CLASS = 'docx-table-toolbar';

/** Vendor DOM hooks used to locate chrome before our attrs exist. */
export const DOCX_VENDOR_ROOT_SELECTOR = '.docx-editor-root.docx-editor, [data-testid="docx-editor"]';
export const DOCX_VENDOR_TOOLBAR_SELECTOR = '[data-testid="editor-toolbar"]';
export const DOCX_VENDOR_TITLE_BAR_SELECTOR = '[data-testid="title-bar"]';
export const DOCX_VENDOR_FORMATTING_BAR_SELECTOR = '[data-testid="formatting-bar"]';
export const DOCX_VENDOR_MENUBAR_SELECTOR = '[role="menubar"]';
export const DOCX_VENDOR_SCROLL_CONTAINER_SELECTOR = '.docx-editor__scroll-container';
export const DOCX_VENDOR_RENDERED_PARAGRAPH_SELECTOR = '.layout-paragraph';

function stampAttribute(element: Element, attribute: string): void {
	if (!element.hasAttribute(attribute)) {
		element.setAttribute(attribute, 'true');
	}
}

function stampFirstMatch(scope: ParentNode, selector: string, attribute: string): void {
	const element = scope.querySelector(selector);
	if (element) {
		stampAttribute(element, attribute);
	}
}

function stampAllMatches(scope: ParentNode, selector: string, attribute: string): void {
	scope.querySelectorAll(selector).forEach((element) => {
		stampAttribute(element, attribute);
	});
}

/**
 * Stamps plugin-owned region markers onto vendored DOCX / vendor DOM.
 * Maps vendor hooks (`.docx-editor-root`, `data-testid`, `[role="menubar"]`) onto
 * `data-native-powerpoint-doc-editor-*` attrs that `styles.css` and chrome sync use.
 * Safe to call repeatedly. Pass the host element (or any ancestor) — does not
 * require the root attr to already exist.
 */
export function stampDocxEditorChromeRegions(scope: ParentNode): void {
	const roots = Array.from(
		scope.querySelectorAll<HTMLElement>(DOCX_VENDOR_ROOT_SELECTOR),
	);

	if (roots.length === 0) {
		// Host may already be the root, or only partial chrome exists.
		stampChromeWithin(scope);
		return;
	}

	for (const root of roots) {
		stampAttribute(root, DOCX_EDITOR_ROOT_ATTRIBUTE);
		stampChromeWithin(root);
	}
}

function stampChromeWithin(scope: ParentNode): void {
	stampFirstMatch(scope, DOCX_VENDOR_TOOLBAR_SELECTOR, DOCX_EDITOR_TOOLBAR_ATTRIBUTE);
	stampFirstMatch(scope, DOCX_VENDOR_TITLE_BAR_SELECTOR, DOCX_EDITOR_TITLE_BAR_ATTRIBUTE);
	stampFirstMatch(scope, DOCX_VENDOR_FORMATTING_BAR_SELECTOR, DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE);
	stampFirstMatch(scope, DOCX_VENDOR_SCROLL_CONTAINER_SELECTOR, DOCX_EDITOR_SCROLL_CONTAINER_ATTRIBUTE);
	stampAllMatches(scope, DOCX_VENDOR_MENUBAR_SELECTOR, DOCX_EDITOR_MENUBAR_ATTRIBUTE);

	scope.querySelectorAll<HTMLElement>(`${DOCX_VENDOR_MENUBAR_SELECTOR} > div`).forEach((menuRoot) => {
		stampAttribute(menuRoot, DOCX_EDITOR_MENU_ROOT_ATTRIBUTE);
		const menuButton = menuRoot.querySelector(':scope > button');
		if (menuButton) {
			stampAttribute(menuButton, DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE);
		}

		// MenuDropdown panel is a sibling of the trigger (only present while open).
		Array.from(menuRoot.children).forEach((child) => {
			if (!isElementLike(child) || child === menuButton) {
				return;
			}
			stampAttribute(child, DOCX_EDITOR_MENU_DROPDOWN_ATTRIBUTE);
			Array.from(child.children).forEach((itemWrapper) => {
				if (!isElementLike(itemWrapper)) {
					return;
				}
				const itemButton = itemWrapper.querySelector(':scope > button');
				if (!itemButton) {
					return;
				}
				stampAttribute(itemWrapper, DOCX_EDITOR_MENU_ITEM_ATTRIBUTE);
				stampAttribute(itemButton, DOCX_EDITOR_MENU_ITEM_BUTTON_ATTRIBUTE);
			});
		});
	});

	stampFirstMatch(scope, '[data-testid="font-size-decrease"]', DOCX_FONT_SIZE_DECREASE_ATTRIBUTE);
	stampFirstMatch(scope, '[data-testid="font-size-increase"]', DOCX_FONT_SIZE_INCREASE_ATTRIBUTE);
	stampFirstMatch(scope, '[data-testid="font-size-input"]', DOCX_FONT_SIZE_INPUT_ATTRIBUTE);
	stampFirstMatch(scope, '[data-testid="font-size-display"]', DOCX_FONT_SIZE_DISPLAY_ATTRIBUTE);

	stampAllMatches(scope, '.paged-editor__pages', DOCX_EDITOR_PAGES_ATTRIBUTE);
	stampAllMatches(scope, '.paged-editor__hidden-pm', DOCX_HIDDEN_PROSEMIRROR_ATTRIBUTE);
	stampAllMatches(scope, '[data-testid="caret"]', DOCX_CARET_ATTRIBUTE);
	stampAllMatches(scope, '.layout-page', DOCX_RENDERED_PAGE_ATTRIBUTE);
	stampAllMatches(scope, '.layout-page-content', DOCX_RENDERED_PAGE_CONTENT_ATTRIBUTE);
	stampAllMatches(scope, DOCX_VENDOR_RENDERED_PARAGRAPH_SELECTOR, DOCX_RENDERED_PARAGRAPH_ATTRIBUTE);
	stampAllMatches(
		scope,
		`.${DOCX_VENDOR_TABLE_TOOLBAR_CLASS}[role="toolbar"]`,
		DOCX_TABLE_TOOLBAR_ATTRIBUTE,
	);
}
