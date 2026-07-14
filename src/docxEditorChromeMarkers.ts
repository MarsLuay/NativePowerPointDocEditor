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
export const DOCX_EIGENPAL_TOOLTIP_ATTRIBUTE = 'data-native-powerpoint-doc-editor-eigenpal-tooltip';
export const DOCX_EIGENPAL_TOOLTIP_SELECTOR = `[${DOCX_EIGENPAL_TOOLTIP_ATTRIBUTE}]`;

type AttributeTarget = Pick<Element, 'setAttribute'>;

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

function stampAttribute(element: Element, attribute: string): void {
	if (!element.hasAttribute(attribute)) {
		element.setAttribute(attribute, 'true');
	}
}

/**
 * Stamps plugin-owned region markers onto vendored DOCX DOM that Eigenpal does
 * not mark yet (table context toolbar). Safe to call repeatedly from chrome sync.
 */
export function stampDocxEditorChromeRegions(root: ParentNode): void {
	root.querySelectorAll<HTMLElement>(`.${DOCX_VENDOR_TABLE_TOOLBAR_CLASS}[role="toolbar"]`).forEach((toolbar) => {
		stampAttribute(toolbar, DOCX_TABLE_TOOLBAR_ATTRIBUTE);
	});
}
