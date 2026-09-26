import {
	DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE,
	DOCX_EDITOR_FORMATTING_BAR_SELECTOR,
	DOCX_EDITOR_MENUBAR_ATTRIBUTE,
	DOCX_EDITOR_MENUBAR_SELECTOR,
	DOCX_EDITOR_PAGES_ATTRIBUTE,
	DOCX_EDITOR_PAGES_SELECTOR,
	DOCX_EDITOR_SCROLL_CONTAINER_SELECTOR,
	DOCX_EDITOR_TITLE_BAR_ATTRIBUTE,
	DOCX_EDITOR_TITLE_BAR_SELECTOR,
	DOCX_EDITOR_TOOLBAR_ATTRIBUTE,
	DOCX_EDITOR_TOOLBAR_SELECTOR,
	DOCX_VENDOR_TABLE_TOOLBAR_CLASS,
} from './docxEditorChromeMarkers';

/** Canonical editor chrome regions shared by DOCX and PPTX surfaces. */
export type EditorChromeRegionId =
	| 'header'
	| 'menubar'
	| 'toolbar'
	| 'contextToolbar'
	| 'documentSurface';

export interface EditorChromeRegionDocxBinding {
	attribute: string;
	selector: string;
}

export interface EditorChromeRegionPptxBinding {
	className: string;
	selector: string;
}

export interface EditorChromeRegionBinding {
	id: EditorChromeRegionId;
	docx: EditorChromeRegionDocxBinding;
	pptx: EditorChromeRegionPptxBinding;
}

export const PPTX_EDITOR_CHROME_HEADER_CLASS = 'native-powerpoint-headerbar';
export const PPTX_EDITOR_CHROME_MENUBAR_CLASS = 'native-powerpoint-menubar';
export const PPTX_EDITOR_CHROME_TOOLBAR_CLASS = 'native-powerpoint-toolbar';
export const PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS = 'native-powerpoint-text-toolbar';
export const PPTX_EDITOR_CHROME_DOCUMENT_SURFACE_CLASS = 'native-powerpoint-canvas-pane';
export const PPTX_EDITOR_CHROME_DOCUMENT_CONTENT_CLASS = 'native-powerpoint-slide-surface';
export const PPTX_EDITOR_CHROME_TOOLBAR_POPOVER_CLASS = 'native-powerpoint-toolbar-popover';

export { DOCX_VENDOR_TABLE_TOOLBAR_CLASS };

/**
 * Shared chrome-region contract.
 * DOCX resolves through `data-native-powerpoint-doc-editor-*` markers (vendored
 * Vendor DOM plus plugin stamping). PPTX resolves through owned classes.
 */
export const EDITOR_CHROME_REGIONS: Record<EditorChromeRegionId, EditorChromeRegionBinding> = {
	header: {
		id: 'header',
		docx: {
			attribute: DOCX_EDITOR_TITLE_BAR_ATTRIBUTE,
			selector: DOCX_EDITOR_TITLE_BAR_SELECTOR,
		},
		pptx: {
			className: PPTX_EDITOR_CHROME_HEADER_CLASS,
			selector: `.${PPTX_EDITOR_CHROME_HEADER_CLASS}`,
		},
	},
	menubar: {
		id: 'menubar',
		docx: {
			attribute: DOCX_EDITOR_MENUBAR_ATTRIBUTE,
			selector: DOCX_EDITOR_MENUBAR_SELECTOR,
		},
		pptx: {
			className: PPTX_EDITOR_CHROME_MENUBAR_CLASS,
			selector: `.${PPTX_EDITOR_CHROME_MENUBAR_CLASS}`,
		},
	},
	toolbar: {
		id: 'toolbar',
		docx: {
			attribute: DOCX_EDITOR_TOOLBAR_ATTRIBUTE,
			selector: DOCX_EDITOR_TOOLBAR_SELECTOR,
		},
		pptx: {
			className: PPTX_EDITOR_CHROME_TOOLBAR_CLASS,
			selector: `.${PPTX_EDITOR_CHROME_TOOLBAR_CLASS}`,
		},
	},
	contextToolbar: {
		id: 'contextToolbar',
		docx: {
			attribute: DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE,
			selector: `${DOCX_EDITOR_FORMATTING_BAR_SELECTOR}, [data-native-powerpoint-doc-editor-table-toolbar]`,
		},
		pptx: {
			className: PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS,
			selector: `.${PPTX_EDITOR_CHROME_CONTEXT_TOOLBAR_CLASS}`,
		},
	},
	documentSurface: {
		id: 'documentSurface',
		docx: {
			attribute: DOCX_EDITOR_PAGES_ATTRIBUTE,
			selector: `${DOCX_EDITOR_SCROLL_CONTAINER_SELECTOR}, ${DOCX_EDITOR_PAGES_SELECTOR}`,
		},
		pptx: {
			className: PPTX_EDITOR_CHROME_DOCUMENT_SURFACE_CLASS,
			selector: `.${PPTX_EDITOR_CHROME_DOCUMENT_SURFACE_CLASS}`,
		},
	},
};

export const PPTX_EDITOR_FORMATTING_SURFACE_SELECTOR = [
	EDITOR_CHROME_REGIONS.menubar.pptx.selector,
	'.native-powerpoint-menubar-dropdown',
	EDITOR_CHROME_REGIONS.toolbar.pptx.selector,
	EDITOR_CHROME_REGIONS.contextToolbar.pptx.selector,
	`.${PPTX_EDITOR_CHROME_TOOLBAR_POPOVER_CLASS}`,
].join(', ');

export function getEditorChromeRegion(region: EditorChromeRegionId): EditorChromeRegionBinding {
	return EDITOR_CHROME_REGIONS[region];
}

export function getDocxEditorChromeRegionSelector(region: EditorChromeRegionId): string {
	return EDITOR_CHROME_REGIONS[region].docx.selector;
}

export function getPptxEditorChromeRegionSelector(region: EditorChromeRegionId): string {
	return EDITOR_CHROME_REGIONS[region].pptx.selector;
}
