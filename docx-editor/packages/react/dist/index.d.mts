/**
 * @eigenpal/docx-editor-react
 *
 * Curated root entry for the documented React editor API. Advanced surfaces
 * stay public through explicit subpaths:
 * - `@eigenpal/docx-editor-react/ui`
 * - `@eigenpal/docx-editor-react/dialogs`
 * - `@eigenpal/docx-editor-react/hooks`
 * - `@eigenpal/docx-editor-react/plugin-api`
 *
 * Framework-agnostic document utilities live in `@eigenpal/docx-editor-core`.
 * Agent/MCP surfaces live in `@eigenpal/docx-editor-agents`.
 *
 * @packageDocumentation
 * @public
 */
import * as React from 'react';
import { ReactNode, CSSProperties } from 'react';
import * as prosemirror_view from 'prosemirror-view';
import { EditorView } from 'prosemirror-view';
import * as prosemirror_state from 'prosemirror-state';
import { EditorState, Transaction } from 'prosemirror-state';
import { Document, HeaderFooter, Theme } from '@eigenpal/docx-editor-core/types/document';
import { FontOption } from '@eigenpal/docx-editor-core/utils/fontOptions';
import { R as ReactSidebarItem } from './types-D35gNE-_.mjs';
import { Comment } from '@eigenpal/docx-editor-core/types/content';
import { Translations, TFunction } from '@eigenpal/docx-editor-i18n';
import { PrintOptions, EditorHandle } from '@eigenpal/docx-editor-core';
export { CreateEmptyDocumentOptions, createDocumentWithText, createEmptyDocument } from '@eigenpal/docx-editor-core';
import { DocumentAgent, ContentControlFilter, ContentControlValue } from '@eigenpal/docx-editor-core/agent';
import { ScrollToParaIdOptions, DocxInput, FontDefinition } from '@eigenpal/docx-editor-core/utils';
import { SelectionState, PMContentControl } from '@eigenpal/docx-editor-core/prosemirror';
import { Layout } from '@eigenpal/docx-editor-core/layout-engine';
import { RenderedDomContext } from '@eigenpal/docx-editor-core/plugin-api';

/**
 * Options for the agent panel mount on the right side of the editor.
 *
 * Three control patterns:
 *  - **Uncontrolled**: `agentPanel={{ render }}` — toolbar button + panel
 *    close button toggle the panel. Width persists to localStorage.
 *  - **Controlled**: `agentPanel={{ render, open, onOpenChange }}` — the
 *    consumer owns open state (e.g. tied to a global menu).
 *  - **Headless**: omit `agentPanel`, use the toolkit directly via
 *    `useDocxAgentTools` — render the panel anywhere you want.
 */
interface AgentPanelOptions {
    /** Render-prop returning the panel content. Called only when open. */
    render: (ctx: {
        close: () => void;
    }) => ReactNode;
    /** Controlled open state. Omit for uncontrolled. */
    open?: boolean;
    /** Fires when toolbar button or panel close button is clicked. */
    onOpenChange?: (open: boolean) => void;
    /** Show the toolbar toggle button. Default: true. */
    showToolbarButton?: boolean;
    /** Optional badge / dot on the toolbar button. */
    toolbarBadge?: ReactNode;
    /** Optional panel title. Default: t('agentPanel.defaultTitle'). */
    title?: string;
    /** Optional panel header icon. Default: sparkle. */
    icon?: ReactNode;
    /** Initial panel width in px (uncontrolled). Default: 360. */
    defaultWidth?: number;
    /** Min drag width. Default: 280. */
    minWidth?: number;
    /** Max drag width. Default: 600. */
    maxWidth?: number;
}

interface PagedEditorRef {
    /** Get the current document. */
    getDocument(): Document | null;
    /** Get the ProseMirror EditorState. */
    getState(): EditorState | null;
    /** Get the ProseMirror EditorView. */
    getView(): EditorView | null;
    /** Focus the editor. */
    focus(): void;
    /** Blur the editor. */
    blur(): void;
    /** Check if focused. */
    isFocused(): boolean;
    /** Dispatch a transaction. */
    dispatch(tr: Transaction): void;
    /** Undo. */
    undo(): boolean;
    /** Redo. */
    redo(): boolean;
    /** Set selection by PM position. */
    setSelection(anchor: number, head?: number): void;
    /** Get current layout. */
    getLayout(): Layout | null;
    /** Force re-layout. */
    relayout(): void;
    /** Scroll the visible pages to bring a PM position into view. */
    scrollToPosition(pmPos: number): void;
    /**
     * Scroll to the paragraph identified by Word `w14:paraId` / PM `paraId`.
     * Pass `options.highlight` to briefly flash rendered paragraph fragments.
     * @returns whether a matching paragraph was found
     */
    scrollToParaId(paraId: string, options?: ScrollToParaIdOptions): boolean;
    /**
     * Scroll the paginated view so `pageNumber` (1-indexed) is in view.
     * No-op if the layout isn't ready yet or pageNumber is out of range.
     */
    scrollToPage(pageNumber: number): void;
    /**
     * Scroll to the comment identified by `commentId` and select its range so
     * the selection overlay highlights it. Resolves the id → PM range via the
     * live comment marks; returns `false` (not a throw, not a silent no-op)
     * when the id no longer resolves so the caller can surface a "location no
     * longer exists" affordance.
     */
    scrollToCommentId(commentId: number): boolean;
    /**
     * Scroll to the tracked change identified by `revisionId` and select its
     * range so the selection overlay highlights it. Resolves the id → PM range
     * via the live tracked-change marks; returns `false` when the id no longer
     * resolves (the change was accepted/rejected/deleted).
     */
    scrollToChangeId(revisionId: number): boolean;
    /**
     * Select the PM position range `[from, to]` so the selection overlay
     * highlights it, and scroll its start into view. No-op for a malformed
     * range or a `from` past the document end; `to` is clamped to the document
     * size (raw caller positions, so out-of-range must not throw).
     */
    highlightRange(from: number, to: number): void;
    /**
     * Look up the persistent hidden HF PM EditorView for a given HeaderFooter
     * instance. Returns null when none is mounted (no document, or `hf` is not
     * present in `Document.package.headers/footers`). Phase 2 of the HF
     * unification: the inline overlay uses this to replicate edits into the
     * persistent PM so the painter — which reads from the persistent PM per
     * phase 1 — re-renders live during typing. Phase 5 deletes the inline
     * overlay's PM and this method's only remaining caller is the click /
     * focus router (phase 3).
     */
    getHfPmView(hf: HeaderFooter): EditorView | null;
    /** Get all active header/footer EditorViews mapped by rId. */
    getHfPmViews(): Map<string, EditorView>;
}

/**
 * EditorMode union + the catalog the editing-mode dropdown renders.
 * Lives next to DocxEditor.tsx so the dropdown component and the parent
 * forwardRef body share one source of truth.
 */

type EditorMode = 'editing' | 'suggesting' | 'viewing';

/**
 * DocxEditor props
 */
interface DocxEditorProps {
    /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
    documentBuffer?: DocxInput | null;
    /** Pre-parsed document (alternative to documentBuffer) */
    document?: Document | null;
    /** Callback when document is saved */
    onSave?: (buffer: ArrayBuffer) => void;
    /**
     * Callback when a DOCX file is selected through `File > Open` or Cmd/Ctrl+O.
     * Pass it to route the picked file through your own import pipeline. Omit it
     * to keep the built-in local document load behavior.
     */
    onOpen?: (file: File) => void | Promise<void>;
    /** Author name used for comments and track changes */
    author?: string;
    /** Callback when document changes */
    onChange?: (document: Document) => void;
    /** Callback when selection changes */
    onSelectionChange?: (state: SelectionState | null) => void;
    /** Callback on error */
    onError?: (error: Error) => void;
    /** Callback when fonts are loaded */
    onFontsLoaded?: () => void;
    /** External ProseMirror plugins (from PluginHost) */
    externalPlugins?: prosemirror_state.Plugin[];
    /**
     * When true, the editor treats the `document` prop as a schema seed only and
     * does not load it into ProseMirror on mount. Content is expected to come from
     * external sources — typically `externalPlugins` such as `ySyncPlugin` from
     * `y-prosemirror`, but also any code that dispatches transactions directly.
     *
     * You must still pass a `document` prop (e.g., `createEmptyDocument()`) so the
     * editor can build its schema and render the shell.
     */
    externalContent?: boolean;
    /** Callback when editor view is ready (for PluginHost) */
    onEditorViewReady?: (view: prosemirror_view.EditorView) => void;
    /** Color theme mode for UI styling. `'system'` follows the OS preference. */
    colorMode?: 'light' | 'dark' | 'system';
    /** Document theme schema object */
    theme?: Theme | null;
    /** Whether to show toolbar (default: true) */
    showToolbar?: boolean;
    /**
     * Whether to show `File > Open` and enable Cmd/Ctrl+O (default: true).
     * Set false when you provide your own open action elsewhere.
     */
    showFileOpen?: boolean;
    /** Whether to show the Help menu in the menu bar (default: true) */
    showHelpMenu?: boolean;
    /** Whether to show zoom control (default: true) */
    showZoomControl?: boolean;
    /** Whether to show page margin guides/boundaries (default: false) */
    showMarginGuides?: boolean;
    /** Color for margin guides (default: '#c0c0c0') */
    marginGuideColor?: string;
    /** Whether to show horizontal ruler (default: false) */
    showRuler?: boolean;
    /** Unit for ruler display (default: 'inch') */
    rulerUnit?: 'inch' | 'cm';
    /** Initial zoom level (default: 1.0) */
    initialZoom?: number;
    /** Whether the editor is read-only. When true, hides toolbar and rulers */
    readOnly?: boolean;
    /**
     * When true, the editor does not intercept Cmd/Ctrl+F or Cmd/Ctrl+H.
     * This lets the browser or host app handle native find/history shortcuts.
     */
    disableFindReplaceShortcuts?: boolean;
    /** Custom toolbar actions */
    toolbarExtra?: ReactNode;
    /** Additional CSS class name */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Placeholder when no document */
    placeholder?: ReactNode;
    /** Loading indicator */
    loadingIndicator?: ReactNode;
    /** Whether to show the document outline sidebar (default: false) */
    showOutline?: boolean;
    /** Whether to show the floating outline toggle button (default: true) */
    showOutlineButton?: boolean;
    /**
     * Custom list of fonts shown in the toolbar's font-family dropdown.
     * Strings render in the "Other" group; pass `FontOption[]` for category
     * grouping and CSS fallback chains. Omit to use the built-in 12-font
     * default. An empty array renders an empty (but enabled) dropdown.
     *
     * Pass a stable reference (memoized or module-level) — inline arrays
     * create a new identity per render and invalidate the picker's memo.
     *
     * @example fontFamilies={['Arial', 'Roboto']}
     * @example fontFamilies={[{ name: 'Roboto', fontFamily: 'Roboto, sans-serif', category: 'sans-serif' }]}
     */
    fontFamilies?: ReadonlyArray<string | FontOption>;
    /**
     * Custom font faces to register with the browser before the editor measures
     * text. Each entry injects an `@font-face` rule. Pass a URL (woff2/woff/
     * ttf/otf), an ArrayBuffer, or omit `src` to load by name from Google Fonts.
     * Multiple entries can share `family` to register different weights/styles.
     *
     * Pass a stable reference — inline arrays re-register faces on each render
     * (the loader dedupes by `family|weight|style`, so it's harmless but wastes
     * work).
     *
     * @example
     * fonts={[
     *   { family: 'Custom Sans', src: '/fonts/CustomSans-Regular.woff2' },
     *   { family: 'Custom Sans', src: '/fonts/CustomSans-Bold.woff2', weight: 700 },
     * ]}
     */
    fonts?: ReadonlyArray<FontDefinition>;
    /**
     * Text-watermark presets shown in the watermark dialog's preset dropdown.
     * Omit to use the built-in MS Word phrases (`DEFAULT_WATERMARK_PRESETS`:
     * CONFIDENTIAL, DRAFT, DO NOT COPY, SAMPLE, URGENT, ASAP). Pass an empty
     * array to hide the preset dropdown and require custom text.
     *
     * @example watermarkPresets={['INTERNAL', 'PROPRIETARY', 'COPY']}
     */
    watermarkPresets?: readonly string[];
    /** Print options for print preview */
    printOptions?: PrintOptions;
    /**
     * Callback when print is triggered. Pass it to enable the `File > Print`
     * menu entry; omit to hide. The imperative `ref.current.print()` also
     * invokes this callback.
     */
    onPrint?: () => void;
    /** Callback when content is copied */
    onCopy?: () => void;
    /** Callback when content is cut */
    onCut?: () => void;
    /** Callback when content is pasted */
    onPaste?: () => void;
    /** Editor mode: 'editing' (direct edits), 'suggesting' (track changes), or 'viewing' (read-only). Default: 'editing' */
    mode?: EditorMode;
    /** Callback when the editing mode changes */
    onModeChange?: (mode: EditorMode) => void;
    /** Callback when a comment is added via the UI */
    onCommentAdd?: (comment: Comment) => void;
    /** Callback when a comment is resolved via the UI */
    onCommentResolve?: (comment: Comment) => void;
    /** Callback when a comment is deleted via the UI */
    onCommentDelete?: (comment: Comment) => void;
    /** Callback when a reply is added to a comment via the UI */
    onCommentReply?: (reply: Comment, parent: Comment) => void;
    /**
     * Controlled comments array. When provided, the editor reads comment thread
     * metadata (text, author, replies, resolved status) from this prop instead
     * of internal state, and emits every change through `onCommentsChange`.
     *
     * Use this with collaboration backends (Yjs, Liveblocks, Automerge, …) so
     * comment threads sync across peers — the PM document only carries the
     * range markers; thread metadata lives outside the doc and needs its own
     * sync channel.
     *
     * If omitted, the editor falls back to internal state (current behavior).
     * The granular `onCommentAdd`/`onCommentResolve`/`onCommentDelete`/
     * `onCommentReply` callbacks fire in both modes.
     */
    comments?: Comment[];
    /** Fires whenever the comments array changes (controlled mode). */
    onCommentsChange?: (comments: Comment[]) => void;
    /** Controlled comments-sidebar visibility; source of truth when set. Pair with `onCommentsSidebarOpenChange`; omit for the default self-managed behavior. */
    commentsSidebarOpen?: boolean;
    /** Fires with the next open state whenever the editor wants to show or hide the comments sidebar. Fires in both controlled and uncontrolled modes. */
    onCommentsSidebarOpenChange?: (open: boolean) => void;
    /**
     * Callback when rendered DOM context is ready (for plugin overlays).
     * Used by PluginHost to get access to the rendered page DOM for positioning.
     */
    onRenderedDomContextReady?: (context: RenderedDomContext) => void;
    /**
     * Plugin overlays to render inside the editor viewport.
     * Passed from PluginHost to render plugin-specific overlays.
     */
    pluginOverlays?: ReactNode;
    /** Sidebar items from plugins (passed from PluginHost). */
    pluginSidebarItems?: ReactSidebarItem[];
    /** Rendered DOM context from PluginHost (for sidebar position resolution). */
    pluginRenderedDomContext?: RenderedDomContext | null;
    /** Custom logo/icon for the title bar */
    renderLogo?: () => ReactNode;
    /** Document name shown in the title bar */
    documentName?: string;
    /** Callback when document name changes */
    onDocumentNameChange?: (name: string) => void;
    /** Whether the document name is editable (default: true) */
    documentNameEditable?: boolean;
    /** Custom right-side actions for the title bar */
    renderTitleBarRight?: () => ReactNode;
    /** Translation overrides. Import a locale JSON file and pass it directly. */
    i18n?: Translations;
    /**
     * Mount a controllable agent panel on the right side of the editor. The
     * panel is the chrome (header, close button, drag-resize); the consumer
     * supplies whatever content goes inside via `render` — typically a chat
     * UI from `@ai-sdk/react`'s `useChat`, `assistant-ui`, or any other
     * framework. We do not ship message bubbles, a composer, or a chat engine.
     *
     * Three control patterns:
     *  - **Uncontrolled**: `agentPanel={{ render }}` — toolbar button + panel
     *    close button toggle the panel. Width persists to localStorage.
     *  - **Controlled**: `agentPanel={{ render, open, onOpenChange }}` — the
     *    consumer owns open state (e.g. tied to a global menu).
     *  - **Headless**: omit `agentPanel`, use the toolkit directly via
     *    `useDocxAgentTools` — render the panel anywhere you want.
     */
    agentPanel?: AgentPanelOptions;
}
/**
 * DocxEditor ref interface
 */
interface DocxEditorRef {
    /** Get the DocumentAgent for programmatic access */
    getAgent: () => DocumentAgent | null;
    /** Get the current document */
    getDocument: () => Document | null;
    /** Get the editor ref */
    getEditorRef: () => PagedEditorRef | null;
    /** Save the document to buffer. Pass { selective: false } to force full repack. */
    save: (options?: {
        selective?: boolean;
    }) => Promise<ArrayBuffer | null>;
    /** Set zoom level */
    setZoom: (zoom: number) => void;
    /** Get current zoom level */
    getZoom: () => number;
    /** Focus the editor */
    focus: () => void;
    /** Get current page number */
    getCurrentPage: () => number;
    /** Get total page count */
    getTotalPages: () => number;
    /**
     * Scroll the paginated view so the given page is in view.
     * Page numbers are 1-indexed (matches `getCurrentPage` / `getTotalPages`).
     * No-op for out-of-range or non-integer values.
     * @example ref.current?.scrollToPage(2)
     */
    scrollToPage: (pageNumber: number) => void;
    /**
     * Scroll the paginated view to the paragraph with the given Word `w14:paraId`.
     * Pass `options.highlight` to briefly flash it in a custom color.
     * @returns whether a matching paragraph exists in the ProseMirror document
     * @example ref.current?.scrollToParaId('1A2B3C4D', { highlight: { color: 'rgba(255, 235, 59, 0.55)' } })
     */
    scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
    /**
     * Scroll the paginated view to a specific ProseMirror document position.
     * Use this when you have a raw PM offset; for Word `w14:paraId` use
     * `scrollToParaId` instead.
     * @example ref.current?.scrollToPosition(42)
     */
    scrollToPosition: (pmPos: number) => void;
    /**
     * Scroll the paginated view to the comment with the given id and select its
     * anchored range so the selection overlay highlights it. Resolves the id
     * against the live comment marks at call time.
     * @returns `false` when the id no longer resolves (the comment was deleted
     *   or its anchored text removed between render and click), so the caller
     *   can surface a "location no longer exists" affordance rather than
     *   silently no-op'ing.
     * @example ref.current?.scrollToCommentId(3)
     */
    scrollToCommentId: (commentId: number) => boolean;
    /**
     * Scroll the paginated view to the tracked change with the given Word
     * revision `w:id` and select its range so the selection overlay highlights
     * it. Resolves the id against the live tracked-change marks at call time
     * (matching coalesced revisions the way the changes sidebar does).
     * @returns `false` when the id no longer resolves (the change was
     *   accepted, rejected, or deleted between render and click).
     * @example ref.current?.scrollToChangeId(42)
     */
    scrollToChangeId: (revisionId: number) => boolean;
    /**
     * Select the ProseMirror position range `[from, to]` so the selection
     * overlay highlights it, and scroll its start into view. The selection
     * persists until it next changes (there is no auto-clearing flash). No-op
     * for a malformed range or a `from` past the document end; `to` is clamped
     * to the document size.
     * @example ref.current?.highlightRange(10, 24)
     */
    highlightRange: (from: number, to: number) => void;
    /** Open print preview */
    openPrintPreview: () => void;
    /** Print the document directly */
    print: () => void;
    /** Load a pre-parsed document programmatically */
    loadDocument: (doc: Document) => void;
    /** Load a DOCX buffer programmatically (ArrayBuffer, Uint8Array, Blob, or File) */
    loadDocumentBuffer: (buffer: DocxInput) => Promise<void>;
    /** Add a comment programmatically. Anchored by Word `w14:paraId` so
     * it survives unrelated edits. Returns the comment ID, or null if
     * the paraId is unknown or the search text isn't found / is ambiguous. */
    addComment: (options: {
        paraId: string;
        text: string;
        author: string;
        /** Optional: anchor to a specific phrase within the paragraph (must be unique). */
        search?: string;
    }) => number | null;
    /** Reply to an existing comment. Returns the reply comment ID. */
    replyToComment: (commentId: number, text: string, author: string) => number | null;
    /** Resolve (mark as done) a comment. */
    resolveComment: (commentId: number) => void;
    /** Suggest a tracked change. Pass `replaceWith: ''` to delete the matched text;
     * pass `search: ''` to insert at paragraph end. Returns false on missing paraId,
     * missing/ambiguous search, or attempt to layer on an existing tracked change. */
    proposeChange: (options: {
        paraId: string;
        search: string;
        replaceWith: string;
        author: string;
    }) => boolean;
    /** Locate every paragraph containing `query` (case-insensitive substring).
     * Returns a stable handle (paraId + the matched phrase) the agent can pass
     * back to `addComment` / `proposeChange`. */
    findInDocument: (query: string, options?: {
        caseSensitive?: boolean;
        limit?: number;
    }) => Array<{
        paraId: string;
        match: string;
        before: string;
        after: string;
    }>;
    /**
     * Apply character formatting (bold / italic / color / size / font / etc.)
     * to a paragraph or to a unique phrase within it. This is a direct edit,
     * not a tracked change. Returns false on missing paraId or ambiguous search.
     */
    applyFormatting: (options: {
        paraId: string;
        search?: string;
        marks: {
            bold?: boolean;
            italic?: boolean;
            underline?: boolean | {
                style?: string;
            };
            strike?: boolean;
            color?: {
                rgb?: string;
                themeColor?: string;
            };
            highlight?: string;
            fontSize?: number;
            fontFamily?: {
                ascii?: string;
                hAnsi?: string;
            };
        };
    }) => boolean;
    /**
     * Apply a paragraph style by styleId (e.g. `'Heading1'`, `'Quote'`).
     * Direct edit, not a tracked change. Returns false if paraId is unknown.
     */
    setParagraphStyle: (options: {
        paraId: string;
        styleId: string;
    }) => boolean;
    /**
     * Insert a page or section break after the paragraph identified by `paraId`.
     * `'page'` adds a page break; `'sectionNextPage'` / `'sectionContinuous'`
     * start a new section on a new page / the same page. Direct edit, not a
     * tracked change. Returns false if paraId is unknown.
     */
    insertBreak: (options: {
        paraId: string;
        type: 'page' | 'sectionNextPage' | 'sectionContinuous';
    }) => boolean;
    /**
     * Read the contents of a single page. 1-indexed; returns null if the page
     * does not exist. Each paragraph is returned with its stable paraId so the
     * agent can comment on or modify it without an extra round-trip.
     */
    getPageContent: (pageNumber: number) => {
        pageNumber: number;
        text: string;
        paragraphs: Array<{
            paraId: string;
            text: string;
            styleId?: string;
        }>;
    } | null;
    /** Read the user's current cursor / selection — what's highlighted right now. */
    getSelectionInfo: () => {
        paraId: string | null;
        selectedText: string;
        paragraphText: string;
        before: string;
        after: string;
    } | null;
    /** Get all comments. */
    getComments: () => Comment[];
    /**
     * List block-level content controls (SDTs) in the live document, optionally
     * filtered by `tag`/`alias`/`id`/`type`. Each result includes the control's
     * text and PM position. Anchors for templates and document automation.
     */
    getContentControls: (filter?: ContentControlFilter) => PMContentControl[];
    /** Scroll the first content control matching `filter` into view. Returns false if none. */
    scrollToContentControl: (filter: ContentControlFilter) => boolean;
    /**
     * Replace the content of the first control matching `filter` with `text`
     * (newlines become paragraphs). Returns false if no match. Throws if the
     * control is content-locked unless `{ force: true }`.
     */
    setContentControlContent: (filter: ContentControlFilter, text: string, options?: {
        force?: boolean;
    }) => boolean;
    /**
     * Remove the first control matching `filter`. With `{ keepContent: true }`
     * the inner blocks are unwrapped in place. Returns false if no match. Throws
     * if the control is deletion-locked unless `{ force: true }`.
     */
    removeContentControl: (filter: ContentControlFilter, options?: {
        force?: boolean;
        keepContent?: boolean;
    }) => boolean;
    /**
     * Set a typed value on the first control matching `filter`: a dropdown
     * selection (`{ kind: 'dropdown', value }`), checkbox (`{ kind: 'checkbox',
     * checked }`), or date (`{ kind: 'date', date }`). Updates the visible
     * content and structured state. Returns false if no match; throws if
     * content-locked (unless `force`) or the value doesn't fit the control type.
     */
    setContentControlValue: (filter: ContentControlFilter, value: ContentControlValue, options?: {
        force?: boolean;
    }) => boolean;
    /** Subscribe to document changes. Fires after every committed edit. Returns unsubscribe. */
    onContentChange: (listener: (document: Document) => void) => () => void;
    /** Subscribe to selection changes (cursor moves / selection changes). Returns unsubscribe. */
    onSelectionChange: (listener: (selection: SelectionState | null) => void) => () => void;
}

/**
 * DocxEditor - Complete DOCX editor component
 */
declare const DocxEditor: React.ForwardRefExoticComponent<DocxEditorProps & React.RefAttributes<DocxEditorRef>>;

/**
 * Simple imperative API for rendering a DOCX editor into a DOM element.
 *
 * Returns an `EditorHandle` (from @eigenpal/docx-editor-core) that works with
 * any framework implementation.
 *
 * Usage:
 * ```ts
 * import { renderAsync } from '@eigenpal/docx-editor-react';
 *
 * const editor = await renderAsync(docxBlob, document.getElementById('container'), {
 *   readOnly: false,
 *   showToolbar: true,
 * });
 *
 * // Save the edited document
 * const blob = await editor.save();
 *
 * // Clean up
 * editor.destroy();
 * ```
 */

/**
 * Options for {@link renderAsync}. A subset of DocxEditorProps minus
 * `documentBuffer` / `document` (passed as the first argument instead).
 */
type RenderAsyncOptions = Omit<DocxEditorProps, 'documentBuffer' | 'document'>;
/**
 * React-specific handle that extends the framework-agnostic EditorHandle
 * with zoom control.
 */
interface DocxEditorHandle extends EditorHandle {
    /** Set zoom level (1.0 = 100%). */
    setZoom: (zoom: number) => void;
    /** Scroll to a body paragraph by Word `w14:paraId`. */
    scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
    /** Scroll to a raw ProseMirror document position. */
    scrollToPosition: (pmPos: number) => void;
}
/**
 * Render a DOCX editor into a container element.
 *
 * @param input - DOCX data as ArrayBuffer, Uint8Array, Blob, or File
 * @param container - DOM element to render into
 * @param options - Editor configuration (toolbar, readOnly, callbacks, etc.)
 * @returns A handle with save / destroy / getDocument methods
 */
declare function renderAsync(input: DocxInput, container: HTMLElement, options?: RenderAsyncOptions): Promise<DocxEditorHandle>;

interface LocaleProviderProps {
    i18n?: Translations;
    children: ReactNode;
}
declare function LocaleProvider({ i18n, children }: LocaleProviderProps): React.JSX.Element;
declare function useTranslation(): {
    t: TFunction;
};

/**
 * @eigenpal/docx-editor-react
 *
 * Curated root entry for the documented React editor API. Advanced surfaces
 * stay public through explicit subpaths:
 * - `@eigenpal/docx-editor-react/ui`
 * - `@eigenpal/docx-editor-react/dialogs`
 * - `@eigenpal/docx-editor-react/hooks`
 * - `@eigenpal/docx-editor-react/plugin-api`
 *
 * Framework-agnostic document utilities live in `@eigenpal/docx-editor-core`.
 * Agent/MCP surfaces live in `@eigenpal/docx-editor-agents`.
 *
 * @packageDocumentation
 * @public
 */
declare const VERSION = "0.0.2";

export { DocxEditor, type DocxEditorHandle, type DocxEditorProps, type DocxEditorRef, type EditorMode, LocaleProvider, type LocaleProviderProps, type RenderAsyncOptions, VERSION, renderAsync, useTranslation };
