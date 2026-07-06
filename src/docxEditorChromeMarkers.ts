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

type AttributeTarget = Pick<Element, 'setAttribute'>;

export function markEditorChromeMenuItem(
	element: AttributeTarget,
	kind: DocxEditorChromeMenuItemKind,
) {
	element.setAttribute(EDITOR_CHROME_MENU_ITEMS[kind].attribute, 'true');
	element.setAttribute(EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE, 'true');
}

export function markEditorChromeNoToolbarTooltip(element: AttributeTarget) {
	element.setAttribute(EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE, 'true');
}
