import type { Command, Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { ForwardRefExoticComponent, ReactNode, RefAttributes } from 'react';

export const DOCX_PACKAGE_DOCUMENT_KEY = 'document';

/** Editing states exposed by the vendored DOCX editor. */
export type EditorMode = 'editing' | 'suggesting' | 'viewing';

/** Rendered page DOM made available to plugin integrations. */
export interface RenderedDomContext {
	pagesContainer: HTMLElement;
	zoom: number;
	getCoordinatesForPosition(position: number): RenderedDomCoordinates | null;
}

export interface RenderedDomCoordinates {
	x: number;
	y: number;
	height: number;
}

/** Minimal ProseMirror-backed editor surface NPDE uses through the editor ref. */
export interface DocxEditorCoreRef {
	getView(): EditorView | null;
	relayout(): void;
}

/** Comment fields used by the plugin's review-sidebar controls. */
export interface DocxComment {
	id: number;
	author?: string;
	parentId?: number | null;
	done?: boolean;
}

/** DOCX document fields read by NPDE's pagination and ruler integrations. */
export interface DocxDocument {
	package?: DocxDocumentPackage;
}

export interface DocxDocumentPackage {
	[DOCX_PACKAGE_DOCUMENT_KEY]?: DocxDocumentPart;
	[key: string]: unknown;
}

export interface DocxDocumentPart {
	finalSectionProperties?: DocxSectionProperties;
	sections?: DocxSection[];
}

export interface DocxSection {
	properties?: DocxSectionProperties;
}

export interface DocxSectionProperties {
	pageHeight?: number;
	marginTop?: number;
	marginBottom?: number;
}

/** Imperative DOCX editor surface consumed by NPDE. */
export interface DocxEditorRef {
	getDocument(): DocxDocument | null;
	getEditorRef(): DocxEditorCoreRef | null;
	save(options?: { selective?: boolean }): Promise<ArrayBuffer | null>;
	setZoom(zoom: number): void;
	getZoom(): number;
	getTotalPages(): number;
	scrollToPosition(position: number): void;
	getSelectionInfo(): { selectedText?: string } | null;
	getComments(): readonly DocxComment[];
}

export type FontCategory = 'sans-serif' | 'serif' | 'monospace' | 'other';

export interface FontOption {
	name: string;
	fontFamily: string;
	category?: FontCategory;
}

/** Locale values are intentionally structural: NPDE only reads the language tag. */
export interface Translations {
	_lang?: string;
	[key: string]: string | null | Translations | undefined;
}

export type TranslationKey = string;
export type TranslationVars = Record<string, string | number>;
export type TranslationFunction = (key: TranslationKey, vars?: TranslationVars) => string;

export interface DocxEditorSidebarItem {
	id: string;
	anchorPos: number;
	estimatedHeight?: number;
	priority?: number;
	render: () => ReactNode;
}

/** The DOCX editor JSX props used by this plugin. */
export interface DocxEditorProps {
	className?: string;
	documentBuffer?: ArrayBuffer | Blob | File | Uint8Array | null;
	mode?: EditorMode;
	onModeChange?: (mode: EditorMode) => void;
	author?: string;
	i18n?: Translations;
	commentsSidebarOpen?: boolean;
	onCommentsSidebarOpenChange?: (open: boolean) => void;
	initialZoom?: number;
	colorMode?: 'light' | 'dark' | 'system';
	showRuler?: boolean;
	showToolbar?: boolean;
	showZoomControl?: boolean;
	showOutlineButton?: boolean;
	readOnly?: boolean;
	disableFindReplaceShortcuts?: boolean;
	externalPlugins?: Plugin[];
	fontFamilies?: readonly (FontOption | string)[];
	documentName?: string;
	documentNameEditable?: boolean;
	pluginSidebarItems?: readonly DocxEditorSidebarItem[];
	onRenderedDomContextReady?: (context: RenderedDomContext) => void;
	onEditorViewReady?: (view: EditorView) => void;
	onSelectionChange?: (selection: unknown) => void;
	onDocumentNameChange?: (name: string) => void;
	renderLogo?: () => ReactNode;
	renderTitleBarRight?: () => ReactNode;
	onChange?: (doc: unknown) => void;
	onFontsLoaded?: () => void;
	onSave?: (buffer: ArrayBuffer) => void | Promise<void>;
	onError?: (error: Error) => void;
	onCommentsChange?: (comments: readonly DocxComment[]) => void;
	onCommentAdd?: (comment: DocxComment) => void;
	onCommentReply?: (reply: DocxComment, parent: DocxComment) => void;
	onCommentDelete?: (comment: DocxComment) => void;
	onCommentResolve?: (comment: DocxComment) => void;
	onCommentUnresolve?: (comment: DocxComment) => void;
}

export type DocxEditorComponent = ForwardRefExoticComponent<DocxEditorProps & RefAttributes<DocxEditorRef>>;
export type DocxCommand = Command;
export interface DocxTextRange {
	from: number;
	to: number;
}
