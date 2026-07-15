/**
 * @eigenpal/docx-editor-core (default entry point)
 *
 * Fat barrel that re-exports the parser, serializer, agent, plugin
 * registry, and the most-used types. No React/DOM imports.
 *
 * **When to import from `.` vs `./headless`:** identical for Node.js
 * use; `.` is the convenient aggregate, `./headless` is its mirror with
 * a slightly different name suffix. Adapter authors who only need a
 * specific slice should prefer the smaller subpaths (`./docx`, `./agent`,
 * `./prosemirror`, `./layout-*`, `./utils`) — they tree-shake better.
 *
 * @example
 * ```ts
 * import { parseDocx, serializeDocx, resolveColor } from '@eigenpal/docx-editor-core';
 * ```
 * @packageDocumentation
 * @public
 */
export { parseDocx } from './docx/parser.mjs';
export { s as serializeDocumentBody, a as serializeDocx, b as serializeSectionProperties } from './sectionPropertiesSerializer-Bp3ooY2t.mjs';
export { createDocx, default as repackDocx, updateMultipleFiles } from './docx/rezip.mjs';
export { a as attemptSelectiveSave } from './selectiveSave-jinP_4xa.mjs';
export { b as buildPatchedDocumentXml, v as validatePatchSafety } from './selectiveXmlPatch-ypkxlTD_.mjs';
export { C as CreateEmptyDocumentOptions, P as ProcessTemplateOptions, a as ProcessTemplateResult, z as ThemeMatrixCell, b as blendColors, c as colorsEqual, d as createDocumentWithText, e as createEmptyDocument, f as createRgbColor, h as createThemeColor, i as darkenColor, A as ensureHexPrefix, B as generateThemeTintShadeMatrix, j as getContrastingColor, l as getTemplateTags, D as getThemeTintShadeHex, m as isBlack, n as isWhite, o as lightenColor, p as parseColorString, r as processTemplate, t as processTemplateAsBlob, u as processTemplateDetailed, v as resolveColor, w as resolveHighlightColor, E as resolveHighlightToCss, x as resolveShadingColor, y as validateTemplate } from './colorResolver-BGCmWhJZ.mjs';
export { A as AgentContextOptions, D as DocumentAgent, E as ExtendedSelectionContext, e as SelectionContextOptions, f as buildExtendedSelectionContext, g as buildSelectionContext, k as executeCommand, l as executeCommands, m as getAgentContext, n as getDocumentSummary } from './selectionContext-BjMyVQZO.mjs';
export { emuToPixels, emuToTwips, formatPx, halfPointsToPixels, pixelsToEmu, pixelsToTwips, pointsToPixels, twipsToEmu, twipsToPixels } from './utils/units.mjs';
export { I as InsertPosition, c as canRenderFont, a as countPageBreaks, b as createColumnBreak, d as createHorizontalRule, e as createLineBreak, f as createPageBreak, g as createPageBreakParagraph, h as createPageBreakRun, i as findPageBreaks, j as getLoadedFonts, k as hasPageBreakBefore, l as insertHorizontalRule, m as insertPageBreak, n as isBreakContent, o as isColumnBreak, p as isFontLoaded, q as isFontsLoading, r as isGoogleFontsEnabled, s as isLineBreak, t as isPageBreak, u as loadFont, v as loadFontFromBuffer, w as loadFonts, x as onFontsLoaded, y as preloadCommonFonts, z as removePageBreak, A as setGoogleFontsEnabled } from './fontLoader-ChYTPDqD.mjs';
export { D as DocxInput, t as toArrayBuffer } from './docxInput-DTbCa48g.mjs';
export { f as findParagraphByParaId, a as findStartPosForParaId } from './findParagraphByParaId-Maw_8M5D.mjs';
export { V as VariableDetectionResult, a as VariableOccurrence, d as detectVariables, b as detectVariablesDetailed, c as detectVariablesInBody, e as detectVariablesInParagraph, f as documentHasVariables, g as extractVariablesFromText, h as formatVariable, i as hasTemplateVariables, j as isValidVariableName, p as parseVariable, r as removeVariables, k as replaceVariables, s as sanitizeVariableName } from './variableDetector-C_OU2N8t.mjs';
import { Document } from './types/document.mjs';
export { DocxPackage } from './types/document.mjs';
export { AIAction, AIActionRequest, AgentCommand, AgentContext, AgentResponse, ApplyStyleCommand, DeleteTextCommand, FormatTextCommand, InsertHyperlinkCommand, InsertImageCommand, InsertTableCommand, InsertTextCommand, ParagraphContext, Position, Range, ReplaceTextCommand, SelectionContext, SetVariableCommand, SuggestedAction } from './types/agentApi.mjs';
export { EditorPluginCore, PanelConfig, PluginPanelProps, PositionCoordinates, RenderedDomContext } from './plugin-api/types.mjs';
export { C as CorePlugin, e as McpSession, M as McpToolDefinition, i as McpToolHandler, j as McpToolResult } from './types-DTF0N7UE.mjs';
export { P as PluginRegistry, p as pluginRegistry, r as registerPlugins } from './registry-NLnxX7pq.mjs';
export { docxtemplaterPlugin } from './core-plugins.mjs';
import { S as Subscribable } from './Subscribable-DOz6Ohoo.mjs';
import { ErrorManagerSnapshot, PluginLifecycleSnapshot, PluginLifecycleConfig } from './managers/types.mjs';
export { AutoSaveManagerOptions, AutoSaveSnapshot, AutoSaveStatus, CellCoordinates, EditorHandle, ErrorNotification, ErrorSeverity, SavedDocumentData, TableSelectionSnapshot } from './managers/types.mjs';
export { AutoSaveManager, formatLastSaveTime, formatStorageSize, getAutoSaveStatusLabel, getAutoSaveStorageSize, isAutoSaveSupported } from './managers/AutoSaveManager.mjs';
export { TABLE_DATA_ATTRIBUTES, TableSelectionManager, deleteTableFromDocument, findTableFromClick, getTableFromDocument, updateTableInDocument } from './managers/TableSelectionManager.mjs';
import { l as Run } from './content-BZ9rYecc.mjs';
export { B as BlockContent, z as BookmarkEnd, A as BookmarkStart, C as Comment, f as CommentRangeEnd, g as CommentRangeStart, h as Deletion, D as DocumentBody, E as Endnote, O as Field, V as FooterReference, F as Footnote, y as HeaderFooter, Z as HeaderReference, H as Hyperlink, I as Image, i as Insertion, M as MoveFrom, j as MoveTo, P as Paragraph, k as ParagraphContent, m as RunContent, S as SectionProperties, ak as Shape, T as Table, n as TableCell, o as TableRow, av as TextBox, p as TextContent, q as TrackedChangeInfo, r as TrackedRunChange } from './content-BZ9rYecc.mjs';
import { EditorView } from 'prosemirror-view';
export { FlowBlock, FootnoteContent, Layout, Measure, Page } from './layout-engine/types.mjs';
export { P as ParagraphFormatting, T as TextFormatting } from './formatting-DFtuRFQY.mjs';
export { C as ConvertFootnoteOptions, a as ConvertHeaderFooterOptions, F as FOOTNOTE_SEPARATOR_HEIGHT, b as FootnoteRefLocation, H as HeaderFooterMetrics, M as MAX_FOOTNOTE_LAYOUT_PASSES, c as MeasureBlocksFn, S as StabilizeFootnoteLayoutArgs, d as StabilizeFootnoteLayoutResult, e as buildFootnoteContentMap, f as buildFootnoteRenderItems, g as calculateFootnoteReservedHeights, h as collectFootnoteRefs, i as convertHeaderFooterToContent, j as footnoteReservedHeightsEqual, m as mapFootnotesToPages, s as stabilizeFootnoteLayout } from './headerFooterLayout-BGp1V3Jw.mjs';
export { L as ListLevel, a as NumberingDefinitions } from './lists-CyGxd5Y2.mjs';
export { R as Relationship, a as Style, S as StyleDefinitions, T as Theme, g as ThemeColorScheme, h as ThemeFont, i as ThemeFontScheme } from './styles-BNjUANte.mjs';
import './colors-C3vA7HUU.mjs';
import './docx/wrapTypes.mjs';
import './watermark-D90356ZM.mjs';
import 'jszip';
import 'prosemirror-model';
import 'prosemirror-state';
import './footnotes-BZ24OTAT.mjs';

/**
 * Framework-agnostic print helpers shared by the React and Vue
 * adapters. Lifted from packages/react/src/components/ui/PrintPreview.tsx
 * so both adapters use the same parsing / preview-window code path.
 *
 * The thin button component + the print-time CSS injection stay
 * adapter-local (they're framework-specific JSX/SFC bits); the data
 * helpers below are pure functions.
 */
interface PrintOptions {
    includeHeaders?: boolean;
    includeFooters?: boolean;
    includePageNumbers?: boolean;
    pageRange?: {
        start: number;
        end: number;
    } | null;
    scale?: number;
    printBackground?: boolean;
    margins?: 'default' | 'none' | 'minimum';
}
declare function getDefaultPrintOptions(): PrintOptions;
/** Trigger browser print dialog for the current document. */
declare function triggerPrint(): void;
/**
 * Open a new window with print-optimised body content.
 *
 * Built entirely via DOM APIs (no `document.write` of interpolated strings):
 * `title` is assigned as a property so a value like `</title><script>` is
 * treated as text and cannot break out, and `content` is parsed in an inert
 * document and imported rather than concatenated into markup. `content` is the
 * caller's already-rendered print HTML; provide trusted markup.
 * Sibling copy: packages/react/src/components/ui/PrintPreview.tsx.
 */
declare function openPrintWindow(title: string | undefined, content: string): Window | null;
/** Parse "1", "1-5", etc. into a page range, or null on invalid. */
declare function parsePageRange(input: string, maxPages: number): {
    start: number;
    end: number;
} | null;
declare function formatPageRange(range: {
    start: number;
    end: number;
} | null, totalPages: number): string;
declare function isPrintSupported(): boolean;

/**
 * ClipboardManager
 *
 * Framework-agnostic class for clipboard operations in the editor.
 * Extracted from the React `useClipboard` hook.
 *
 * Handles:
 * - DOM selection traversal and run extraction
 * - Formatting extraction from computed styles
 * - Clipboard read/write operations
 */

/** Selection data for clipboard operations */
interface ClipboardSelection {
    text: string;
    runs: Run[];
    startParagraphIndex: number;
    startRunIndex: number;
    startOffset: number;
    endParagraphIndex: number;
    endRunIndex: number;
    endOffset: number;
    isMultiParagraph: boolean;
}
/**
 * Convert a CSS color string (rgb/rgba/hex) to a 6-char uppercase hex string.
 *
 * NOTE: This differs from `colorResolver.rgbToHex(r, g, b)` which takes
 * numeric components. This function parses CSS color strings.
 */
declare function cssColorToHex(color: string): string | null;
/** Extract formatting from an HTML element's computed styles. */
declare function extractFormattingFromElement(element: HTMLElement): Run['formatting'];
/** Get selected runs from the current DOM selection. */
declare function getSelectionRuns(): Run[];
/** Create a ClipboardSelection from the current DOM selection. */
declare function createSelectionFromDOM(): ClipboardSelection | null;
declare const rgbToHex: typeof cssColorToHex;

/**
 * ErrorManager
 *
 * Framework-agnostic pub/sub error notification system.
 * Replaces React's `componentDidCatch` + context pattern for error notifications.
 *
 * Usage with React:
 * ```ts
 * const { notifications } = useSyncExternalStore(manager.subscribe, manager.getSnapshot);
 * ```
 */

declare class ErrorManager extends Subscribable<ErrorManagerSnapshot> {
    private notifications;
    private idCounter;
    private timers;
    constructor();
    /** Show an error notification (persistent, not auto-dismissed). */
    showError(message: string, details?: string): string;
    /** Show a warning notification (auto-dismissed after 5s). */
    showWarning(message: string, details?: string): string;
    /** Show an info notification (auto-dismissed after 5s). */
    showInfo(message: string, details?: string): string;
    /** Dismiss a notification by ID. */
    dismiss(id: string): void;
    /** Clear all notifications and cancel pending timers. */
    clearAll(): void;
    /** Destroy the manager and clean up all timers. */
    destroy(): void;
    private addNotification;
    private emitSnapshot;
}

/**
 * PluginLifecycleManager
 *
 * Framework-agnostic class for managing editor plugin lifecycle.
 * Extracted from React's `PluginHost.tsx`.
 *
 * Handles:
 * - Plugin initialization and state tracking
 * - Plugin state updates via `updateStates()`
 * - Plugin destroy/cleanup
 *
 * Does NOT handle (framework hosts are responsible for):
 * - CSS injection (use the exported `injectStyles` utility)
 * - DOM event listeners / dispatch wrapping
 */

/** Inject CSS styles into the document head. Returns a cleanup function. */
declare function injectStyles(pluginId: string, css: string): () => void;
declare class PluginLifecycleManager extends Subscribable<PluginLifecycleSnapshot> {
    private plugins;
    private pluginStates;
    private version;
    constructor();
    /**
     * Initialize plugins with an editor view.
     * Calls `plugin.initialize(editorView)` for each plugin.
     *
     * Note: CSS injection and DOM event listeners are the responsibility
     * of the framework-specific host (e.g. React PluginHost).
     */
    initialize(plugins: PluginLifecycleConfig[], editorView: EditorView): void;
    /**
     * Update all plugin states by calling `onStateChange` on each plugin.
     * Returns true if any plugin state changed.
     */
    updateStates(editorView: EditorView): boolean;
    /** Get plugin state by ID. */
    getPluginState<T>(pluginId: string): T | undefined;
    /** Set plugin state by ID. */
    setPluginState<T>(pluginId: string, state: T): void;
    /** Destroy all plugins and clean up. */
    destroy(): void;
    private destroyPlugins;
    private emitSnapshot;
}

/**
 * LayoutCoordinator
 *
 * Framework-agnostic class coordinating the PM state → layout engine →
 * layout painter → selection overlay pipeline.
 *
 * Extracted from PagedEditor.tsx. Manages:
 * - Layout pipeline state (blocks, measures, layout)
 * - Selection state (selectionRects, caretPosition)
 * - Drag selection state
 * - Column resize state
 * - Image interaction state
 *
 * Usage with React:
 * ```ts
 * const snapshot = useSyncExternalStore(coordinator.subscribe, coordinator.getSnapshot);
 * ```
 *
 * NOTE: This class defines the state shape and subscription pattern.
 * Full integration with PagedEditor is done incrementally.
 */

/** Selection rectangle for rendering selection overlays */
interface SelectionRect {
    x: number;
    y: number;
    width: number;
    height: number;
    pageIndex: number;
}
/** Caret position for rendering the blinking cursor */
interface CaretPosition {
    x: number;
    y: number;
    height: number;
    pageIndex: number;
}
/** Info about the currently selected/hovered image */
interface ImageSelectionInfo {
    pmPos: number;
    pageIndex: number;
    rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    widthEmu: number;
    heightEmu: number;
    isInline: boolean;
}
/** Column resize tracking state */
interface ColumnResizeState {
    isResizing: boolean;
    startX: number;
    columnIndex: number;
    tablePmStart: number;
    originalWidths: {
        left: number;
        right: number;
    };
}
/** The full snapshot exposed to UI frameworks */
interface LayoutCoordinatorSnapshot {
    /** Computed page layout, null until first computation */
    hasLayout: boolean;
    /** Selection rectangles for range selection overlay */
    selectionRects: SelectionRect[];
    /** Caret position for cursor overlay */
    caretPosition: CaretPosition | null;
    /** Currently selected/hovered image */
    selectedImageInfo: ImageSelectionInfo | null;
    /** Whether the editor is focused */
    isFocused: boolean;
    /** Whether a text drag is in progress */
    isDragging: boolean;
    /** Whether a column resize is in progress */
    isResizingColumn: boolean;
    /** Whether an image interaction is in progress */
    isImageInteracting: boolean;
    /** Version counter — incremented on every state change */
    version: number;
}
declare class LayoutCoordinator extends Subscribable<LayoutCoordinatorSnapshot> {
    private _hasLayout;
    private _selectionRects;
    private _caretPosition;
    private _isDragging;
    private _dragAnchor;
    private _columnResize;
    private _selectedImageInfo;
    private _isImageInteracting;
    private _isFocused;
    private _version;
    constructor();
    /** Notify that layout has been computed. */
    setLayoutReady(hasLayout: boolean): void;
    /** Update selection rectangles and caret position. */
    updateSelection(selectionRects: SelectionRect[], caretPosition: CaretPosition | null): void;
    /** Start a drag selection from the given PM anchor position. */
    startDrag(anchor: number): void;
    /** End drag selection. */
    endDrag(): void;
    /** Get the drag anchor position. */
    getDragAnchor(): number | null;
    /** Start resizing a table column. */
    startColumnResize(tablePmStart: number, columnIndex: number, startX: number, originalWidths: {
        left: number;
        right: number;
    }): void;
    /** End column resize. */
    endColumnResize(): void;
    /** Get current column resize state. */
    getColumnResize(): ColumnResizeState;
    /** Set the currently selected image. */
    setSelectedImage(imageInfo: ImageSelectionInfo | null): void;
    /** Clear the image selection. */
    clearSelectedImage(): void;
    /** Set whether an image interaction (resize/move) is in progress. */
    setImageInteracting(interacting: boolean): void;
    /** Update focus state. */
    setFocused(focused: boolean): void;
    private emitSnapshot;
}

/**
 * EditorCoordinator
 *
 * Framework-agnostic class managing the document editor lifecycle:
 * - Document parsing and loading
 * - Font loading coordination
 * - Zoom level management
 * - Extension manager initialization
 * - Agent command execution
 *
 * Extracted from DocxEditor.tsx.
 *
 * Usage with React:
 * ```ts
 * const snapshot = useSyncExternalStore(coordinator.subscribe, coordinator.getSnapshot);
 * ```
 *
 * NOTE: This class defines the state shape and coordination logic.
 * Full integration with DocxEditor is done incrementally.
 */

/** Editor loading state */
type EditorLoadingState = 'idle' | 'parsing' | 'loading-fonts' | 'ready' | 'error';
/** Configuration for EditorCoordinator */
interface EditorCoordinatorOptions {
    /** Initial zoom level (default: 1.0) */
    initialZoom?: number;
    /** Callback when the document changes */
    onChange?: (document: Document) => void;
    /** Callback when an error occurs */
    onError?: (error: Error) => void;
}
/** The full snapshot exposed to UI frameworks */
interface EditorCoordinatorSnapshot {
    /** Current loading state */
    loadingState: EditorLoadingState;
    /** Error message if loadingState is 'error' */
    parseError: string | null;
    /** Whether the editor is ready for interaction */
    isReady: boolean;
    /** Current zoom level (1.0 = 100%) */
    zoom: number;
    /** Whether fonts have been loaded */
    fontsLoaded: boolean;
    /** Version counter */
    version: number;
}
declare class EditorCoordinator extends Subscribable<EditorCoordinatorSnapshot> {
    private _loadingState;
    private _parseError;
    private _zoom;
    private _fontsLoaded;
    private _document;
    private _version;
    private onChangeCallback?;
    private onErrorCallback?;
    constructor(options?: EditorCoordinatorOptions);
    /** Signal that document parsing has started. */
    setParsingStarted(): void;
    /** Signal that document parsing completed successfully. */
    setDocumentLoaded(document: Document): void;
    /** Signal that font loading completed. */
    setFontsLoaded(): void;
    /** Signal that an error occurred during loading. */
    setLoadError(error: Error): void;
    /** Get the current document. */
    getDocument(): Document | null;
    /** Update the document (after edits). */
    updateDocument(document: Document): void;
    /** Set the zoom level (1.0 = 100%). */
    setZoom(zoom: number): void;
    /** Get the current zoom level. */
    getZoom(): number;
    private emitSnapshot;
}

/**
 * @eigenpal/docx-editor-core (default entry point)
 *
 * Fat barrel that re-exports the parser, serializer, agent, plugin
 * registry, and the most-used types. No React/DOM imports.
 *
 * **When to import from `.` vs `./headless`:** identical for Node.js
 * use; `.` is the convenient aggregate, `./headless` is its mirror with
 * a slightly different name suffix. Adapter authors who only need a
 * specific slice should prefer the smaller subpaths (`./docx`, `./agent`,
 * `./prosemirror`, `./layout-*`, `./utils`) — they tree-shake better.
 *
 * @example
 * ```ts
 * import { parseDocx, serializeDocx, resolveColor } from '@eigenpal/docx-editor-core';
 * ```
 * @packageDocumentation
 * @public
 */
declare const VERSION = "0.0.2";

export { type CaretPosition, type ClipboardSelection, type ColumnResizeState, Document, EditorCoordinator, type EditorCoordinatorOptions, type EditorCoordinatorSnapshot, type EditorLoadingState, ErrorManager, ErrorManagerSnapshot, type ImageSelectionInfo, LayoutCoordinator, type LayoutCoordinatorSnapshot, PluginLifecycleConfig, PluginLifecycleManager, PluginLifecycleSnapshot, type PrintOptions, Run, type SelectionRect, Subscribable, VERSION, createSelectionFromDOM, extractFormattingFromElement, formatPageRange, getDefaultPrintOptions, getSelectionRuns, injectStyles, isPrintSupported, openPrintWindow, parsePageRange, rgbToHex, triggerPrint };
