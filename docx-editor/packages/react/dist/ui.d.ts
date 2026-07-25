/**
 * @npde/docx-editor-react/ui
 *
 * UI entry point — Toolbar, pickers, dialogs, and UI components.
 *
 * @example
 * ```tsx
 * import { Toolbar, FontPicker, ColorPicker } from '@npde/docx-editor-react/ui';
 * ```
 *
 * @packageDocumentation
 * @public
 */
import * as React$1 from 'react';
import React__default, { CSSProperties, ReactNode } from 'react';
import { Table, TableCell, ParagraphAlignment, ColorValue, Style, Theme, StyleType, SectionProperties, TabStop, Document } from '@npde/docx-editor-core/types/document';
import { FontOption } from '@npde/docx-editor-core/utils/fontOptions';
export { FontOption } from '@npde/docx-editor-core/utils/fontOptions';
import { ListState } from '@npde/docx-editor-core/utils/listState';
export { ListState, createDefaultListState } from '@npde/docx-editor-core/utils/listState';
import { T as TableSelection, a as TableContext, b as TableSplitConfig, c as TableAction } from './useFindReplace-BEVd1x38.js';
export { F as FindReplaceOptions, d as FindReplaceState, e as TableToolbar, f as TableToolbarProps, U as UseFindReplaceReturn, u as useFindReplace } from './useFindReplace-BEVd1x38.js';
import { SelectionContext, AIAction, AgentResponse } from '@npde/docx-editor-core/types/agentApi';
import { TranslationKey } from '@npde/docx-editor-i18n';
export { K as DialogKeyboardShortcut, F as FindReplaceDialog, a as FindReplaceDialogProps, H as HyperlinkData, b as HyperlinkDialog, c as HyperlinkDialogProps, d as KeyboardShortcutsDialog, e as KeyboardShortcutsDialogProps, P as PasteOption, f as PasteSpecialDialog, g as PasteSpecialDialogProps, S as ShortcutCategory, U as UseKeyboardShortcutsDialogOptions, h as UseKeyboardShortcutsDialogReturn, i as UsePasteSpecialOptions, j as UsePasteSpecialReturn, k as formatShortcutKeys, l as getAllCategories, m as getAllPasteOptions, n as getCategoryLabel, o as getCommonShortcuts, p as getDefaultPasteOption, q as getDefaultShortcuts, r as getPasteOption, s as getShortcutsByCategory, t as isPasteSpecialShortcut, u as useHyperlinkDialog, v as useKeyboardShortcutsDialog, w as usePasteSpecial } from './KeyboardShortcutsDialog-DDN-wlZe.js';
export { FindMatch, FindOptions, FindResult, HighlightOptions, createDefaultFindOptions, createSearchPattern, escapeRegexString, findAllMatches, findInDocument, findInParagraph, getDefaultHighlightOptions, getMatchCountText, isEmptySearch, replaceAllInContent, replaceFirstInContent, scrollToMatch } from '@npde/docx-editor-core/utils/findReplace';
export { PrintOptions, formatPageRange as formatPrintPageRange, getDefaultPrintOptions, isPrintSupported, openPrintWindow, parsePageRange, triggerPrint } from '@npde/docx-editor-core';
import '@npde/docx-editor-core/utils';

/**
 * Font Picker Component (Radix UI)
 *
 * A dropdown selector for choosing font families using Radix Select.
 * Row chrome matches StylePicker (padding, truncation, panel size).
 */

interface FontPickerProps {
    value?: string;
    onChange?: (fontFamily: string) => void;
    fonts?: FontOption[];
    /**
     * Fonts the loaded document references that the browser can render. Shown in
     * a "Document fonts" group above the built-in list, deduped against `fonts`.
     */
    documentFonts?: readonly FontOption[];
    disabled?: boolean;
    className?: string;
    placeholder?: string;
    width?: number | string;
    showPreview?: boolean;
}
declare function FontPicker({ value, onChange, fonts, documentFonts, disabled, className, placeholder, width, showPreview, }: FontPickerProps): React$1.JSX.Element;

/**
 * Print Utilities
 *
 * Provides print functionality with:
 * - Print button component for toolbar
 * - Print-specific CSS styles
 * - Browser print dialog trigger
 * - Page range utilities
 */

/**
 * PrintButton props
 */
interface PrintButtonProps {
    /** Callback when print is triggered */
    onPrint: () => void;
    /** Whether the button is disabled */
    disabled?: boolean;
    /** Button label */
    label?: string;
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Show icon */
    showIcon?: boolean;
    /** Compact mode */
    compact?: boolean;
}
/**
 * PrintButton - Standalone print button for toolbar
 */
declare function PrintButton({ onPrint, disabled, label: labelProp, className, style, showIcon, compact, }: PrintButtonProps): React__default.ReactElement;
/**
 * PrintStyles - Injects print-specific CSS
 */
declare function PrintStyles(): React__default.ReactElement;

/**
 * List Buttons Component
 *
 * A component for list formatting controls in the DOCX editor:
 * - Bullet list button
 * - Numbered list button
 * - Toggles list on/off for selection
 * - Indent/outdent for list levels
 */

/**
 * Props for the ListButtons component
 */
interface ListButtonsProps {
    /** Current list state of the selection */
    listState?: ListState;
    /** Callback when bullet list is toggled */
    onBulletList?: () => void;
    /** Callback when numbered list is toggled */
    onNumberedList?: () => void;
    /** Callback to increase list indent */
    onIndent?: () => void;
    /** Callback to decrease list indent */
    onOutdent?: () => void;
    /** Whether the buttons are disabled */
    disabled?: boolean;
    /** Additional CSS class name */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Show indent/outdent buttons */
    showIndentButtons?: boolean;
    /** Compact mode (smaller buttons) */
    compact?: boolean;
    /** Whether the current paragraph has left indentation (for enabling outdent) */
    hasIndent?: boolean;
}
/**
 * List buttons component for bullet/numbered list controls
 */
declare function ListButtons({ listState, onBulletList, onNumberedList, onIndent, onOutdent, disabled, className, style, showIndentButtons, compact, hasIndent, }: ListButtonsProps): React__default.JSX.Element;

/**
 * Document-model table operations — pure functions over the Table type:
 * row/column add+delete, cell merge/split, selection bounds + lookups,
 * grid-anchor calculations for irregular merged tables. Used by
 * TableToolbar.tsx and the React table-selection hook; published from
 * `@npde/docx-editor-react`.
 */

/**
 * Create a table context from a table and selection
 */
declare function createTableContext(table: Table, selection: TableSelection): TableContext;
/**
 * Get column count from a table
 */
declare function getColumnCount(table: Table): number;
/**
 * Get cell at specific row and column index
 */
declare function getCellAt(table: Table, rowIndex: number, columnIndex: number): TableCell | null;
declare function getTableSplitCellDialogConfig(table: Table, rowIndex: number, columnIndex: number): TableSplitConfig | null;
declare function splitTableCell(table: Table, rowIndex: number, columnIndex: number, rows: number, cols: number): Table;
/**
 * Add a row to a table at the specified index
 */
declare function addRow(table: Table, atIndex: number, position?: 'before' | 'after'): Table;
/**
 * Delete a row from a table
 */
declare function deleteRow(table: Table, rowIndex: number): Table;
/**
 * Add a column to a table at the specified index
 */
declare function addColumn(table: Table, atIndex: number, position?: 'before' | 'after'): Table;
/**
 * Delete a column from a table
 */
declare function deleteColumn(table: Table, columnIndex: number): Table;
/**
 * Merge cells in a selection
 */
declare function mergeCells(table: Table, selection: TableSelection): Table;
/**
 * Backward-compatible helper for callers that still use the older merged-cell
 * split behavior directly.
 *
 * User-facing Split cell is now dialog-driven. For document-model tables, use
 * `getTableSplitCellDialogConfig()` and `splitTableCell()` instead.
 */
declare function splitCell(table: Table, rowIndex: number, columnIndex: number): Table;

/**
 * Toolbar Component
 *
 * The customizable formatting rail — undo/redo, zoom, styles, fonts,
 * bold/italic/underline, colors, alignment, lists, table/image context,
 * clear formatting. Used standalone (`<Toolbar ...props>`), inside
 * `<EditorToolbar>` (reads from context via `EditorToolbar.Toolbar`), or
 * embedded inline. Also the home of the `ToolbarButton` / `ToolbarGroup` /
 * `ToolbarSeparator` primitives and the shared `FormattingAction` /
 * `SelectionFormatting` / `ToolbarProps` types.
 */

/**
 * Current formatting state of the selection
 */
interface SelectionFormatting {
    /** Whether selected text is bold */
    bold?: boolean;
    /** Whether selected text is italic */
    italic?: boolean;
    /** Whether selected text is underlined */
    underline?: boolean;
    /** Whether selected text has strikethrough */
    strike?: boolean;
    /** Whether selected text is superscript */
    superscript?: boolean;
    /** Whether selected text is subscript */
    subscript?: boolean;
    /** Font family of selected text */
    fontFamily?: string;
    /** Font size of selected text (in half-points) */
    fontSize?: number;
    /** Text color */
    color?: string;
    /** Highlight color */
    highlight?: string;
    /** Paragraph alignment */
    alignment?: ParagraphAlignment;
    /** List state of the current paragraph */
    listState?: ListState;
    /** Line spacing in twips (OOXML value, 240 = single spacing) */
    lineSpacing?: number;
    /** Paragraph style ID */
    styleId?: string;
    /** Paragraph left indentation in twips */
    indentLeft?: number;
    /** Whether the paragraph is RTL (bidi) */
    bidi?: boolean;
}
/**
 * Formatting action types
 */
type FormattingAction = 'bold' | 'italic' | 'underline' | 'strikethrough' | 'superscript' | 'subscript' | 'clearFormatting' | 'addParagraphBottomBorder' | 'clearParagraphBorders' | 'bulletList' | 'numberedList' | 'indent' | 'outdent' | 'insertLink' | 'setRtl' | 'setLtr' | {
    type: 'fontFamily';
    value: string;
} | {
    type: 'fontSize';
    value: number;
} | {
    type: 'textColor';
    value: ColorValue | string;
} | {
    type: 'highlightColor';
    value: string;
} | {
    type: 'alignment';
    value: ParagraphAlignment;
} | {
    type: 'lineSpacing';
    value: number;
} | {
    type: 'applyStyle';
    value: string;
};
/**
 * Props for the Toolbar (formatting rail) component
 */
interface ToolbarProps {
    /** Current formatting of the selection */
    currentFormatting?: SelectionFormatting;
    /** Callback when a formatting action is triggered */
    onFormat?: (action: FormattingAction) => void;
    /** Callback for undo action */
    onUndo?: () => void;
    /** Callback for redo action */
    onRedo?: () => void;
    /** Whether undo is available */
    canUndo?: boolean;
    /** Whether redo is available */
    canRedo?: boolean;
    /** Whether the toolbar is disabled */
    disabled?: boolean;
    /** Additional CSS class name */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Whether to enable keyboard shortcuts (default: true) */
    enableShortcuts?: boolean;
    /** Ref to the editor container for keyboard events */
    editorRef?: React__default.RefObject<HTMLElement>;
    /** Custom toolbar items to render at the end */
    children?: ReactNode;
    /** When true, renders with display:contents so children flow in the parent flex container */
    inline?: boolean;
    /** Whether to show font family picker (default: true) */
    showFontPicker?: boolean;
    /**
     * Custom list of fonts in the toolbar dropdown. When omitted, the built-in
     * 12-font default is used. Strings render in the "Other" group; pass
     * `FontOption[]` for category grouping and CSS fallback chains.
     * An empty array renders an empty (but enabled) dropdown.
     */
    fontFamilies?: ReadonlyArray<string | FontOption>;
    /**
     * Fonts the loaded document references that the browser can render (embedded
     * faces + system-resolved). Rendered in a "Document fonts" group, deduped
     * against `fontFamilies`. Managed by the editor, not a consumer prop.
     */
    documentFonts?: readonly FontOption[];
    /** Whether to show font size picker (default: true) */
    showFontSizePicker?: boolean;
    /** Whether to show text color picker (default: true) */
    showTextColorPicker?: boolean;
    /** Whether to show highlight color picker (default: true) */
    showHighlightColorPicker?: boolean;
    /** Whether to show alignment buttons (default: true) */
    showAlignmentButtons?: boolean;
    /** Whether to show list buttons (default: true) */
    showListButtons?: boolean;
    /** Whether to show line spacing picker (default: true) */
    showLineSpacingPicker?: boolean;
    /** Whether to show style picker (default: true) */
    showStylePicker?: boolean;
    /** Document styles for the style picker */
    documentStyles?: Style[];
    /** Theme for the style picker / color picker theme matrix */
    theme?: Theme | null;
    /** Callback for print action. Set to enable the File > Print menu entry. */
    onPrint?: () => void;
    /** Callback to open/import a DOCX file (File → Open) */
    onOpen?: () => void;
    /** Callback to save/download the current DOCX (File → Save) */
    onSave?: () => void;
    /** Whether to show zoom control (default: true) */
    showZoomControl?: boolean;
    /** Current zoom level (1.0 = 100%) */
    zoom?: number;
    /** Callback when zoom changes */
    onZoomChange?: (zoom: number) => void;
    /** Callback to refocus the editor after toolbar interactions */
    onRefocusEditor?: () => void;
    /** Callback when a table should be inserted */
    onInsertTable?: (rows: number, columns: number) => void;
    /** Whether to show table insert button (default: true) */
    showTableInsert?: boolean;
    /** Whether to show the Help menu in the menu bar (default: true) */
    showHelpMenu?: boolean;
    /** Callback when user wants to insert an image */
    onInsertImage?: () => void;
    /** Callback when user wants to insert a page break */
    onInsertPageBreak?: () => void;
    /** Callback when user wants to insert a "next page" section break */
    onInsertSectionBreakNextPage?: () => void;
    /** Callback when user wants to insert a "continuous" section break */
    onInsertSectionBreakContinuous?: () => void;
    /** Callback when user wants to insert a table of contents */
    onInsertTOC?: () => void;
    /** Callback when user wants to insert a shape */
    onInsertShape?: (data: {
        shapeType: string;
        width: number;
        height: number;
        fillColor?: string;
        fillType?: string;
        outlineWidth?: number;
        outlineColor?: string;
    }) => void;
    /** Image context when an image is selected */
    imageContext?: {
        wrapType: string;
        displayMode: string;
        cssFloat: string | null;
    } | null;
    /** Callback when image wrap type changes */
    onImageWrapType?: (wrapType: string) => void;
    /** Callback for image transform (rotate/flip) */
    onImageTransform?: (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => void;
    /** Callback to open image properties dialog (alt text + border) */
    onOpenImageProperties?: () => void;
    /** Callback to open page setup dialog */
    onPageSetup?: () => void;
    /** Callback to open the watermark dialog */
    onWatermark?: () => void;
    /** Table context when cursor is in a table */
    tableContext?: {
        isInTable: boolean;
        rowCount?: number;
        columnCount?: number;
        canSplitCell?: boolean;
        hasMultiCellSelection?: boolean;
        cellBorderColor?: ColorValue;
        cellBackgroundColor?: string;
    } | null;
    /** Callback when a table action is triggered */
    onTableAction?: (action: TableAction) => void;
}
/**
 * Props for individual toolbar buttons
 */
interface ToolbarButtonProps {
    /** Whether the button is in active/pressed state */
    active?: boolean;
    /** Whether the button is disabled */
    disabled?: boolean;
    /** Optional label seed for test ids; not shown as a hover tooltip. */
    title?: string;
    /** Click handler */
    onClick?: () => void;
    /** Button content */
    children: ReactNode;
    /** Additional CSS class name */
    className?: string;
    /** ARIA label for accessibility */
    ariaLabel?: string;
}
/**
 * Props for toolbar button groups
 */
interface ToolbarGroupProps$1 {
    /** Group label for accessibility */
    label?: string;
    /** Group content */
    children: ReactNode;
    /** Additional CSS class name */
    className?: string;
}
/**
 * Individual toolbar button with shadcn styling
 */
declare function ToolbarButton({ active, disabled, title, onClick, children, className, ariaLabel, }: ToolbarButtonProps): React__default.JSX.Element;
/**
 * Toolbar button group with modern styling
 */
declare function ToolbarGroup$1({ label, children, className }: ToolbarGroupProps$1): React__default.JSX.Element;
/**
 * Toolbar separator
 */
declare function ToolbarSeparator(): React__default.JSX.Element;
/**
 * Icon-based formatting toolbar — undo/redo, zoom, styles, fonts,
 * bold/italic/underline, colors, alignment, lists, table/image context, clear formatting.
 */
declare function Toolbar(explicitProps: ToolbarProps): React__default.JSX.Element;

/**
 * Props for the EditorToolbar compound component.
 * Extends ToolbarProps with title bar-specific fields.
 */
interface EditorToolbarProps extends ToolbarProps {
}

/**
 * TitleBar and sub-components for the Google Docs-style 2-level toolbar.
 *
 * - TitleBar: two-row layout (row 1: logo + doc name + right actions, row 2: menu bar)
 * - Logo: renders custom logo content left-aligned
 * - DocumentName: editable document name input
 * - MenuBar: File/Format/Insert menus (auto-wired from EditorToolbarContext)
 * - TitleBarRight: right-aligned actions slot
 */

interface LogoProps {
    children: ReactNode;
}
declare function Logo({ children }: LogoProps): React__default.JSX.Element;
interface DocumentNameProps {
    value: string;
    onChange?: (value: string) => void;
    placeholder?: string;
    editable?: boolean;
}
declare function DocumentName({ value, onChange, placeholder, editable }: DocumentNameProps): React__default.JSX.Element;
interface TitleBarRightProps {
    children: ReactNode;
}
declare function TitleBarRight({ children }: TitleBarRightProps): React__default.JSX.Element;
declare function MenuBar(): React__default.JSX.Element;
interface TitleBarProps {
    children: ReactNode;
}
/**
 * TitleBar layout (Google Docs style):
 *
 *   ┌──────────┬────────────────────────────┬──────────────────┐
 *   │          │ Document Name              │                  │
 *   │  Logo    │                            │  Right Actions   │
 *   │          │ File  Format  Insert       │                  │
 *   └──────────┴────────────────────────────┴──────────────────┘
 *
 * Logo and TitleBarRight span full height. DocumentName + MenuBar
 * stack vertically in the center column.
 */
declare function TitleBar({ children }: TitleBarProps): React__default.JSX.Element;

/**
 * EditorToolbar — Google Docs-style 2-level compound component.
 *
 * Usage:
 *   <EditorToolbar {...toolbarProps}>
 *     <EditorToolbar.TitleBar>
 *       <EditorToolbar.Logo><MyIcon /></EditorToolbar.Logo>
 *       <EditorToolbar.DocumentName value={name} onChange={setName} />
 *       <EditorToolbar.MenuBar />
 *       <EditorToolbar.TitleBarRight>
 *         <button>Save</button>
 *       </EditorToolbar.TitleBarRight>
 *     </EditorToolbar.TitleBar>
 *     <EditorToolbar.Toolbar />
 *   </EditorToolbar>
 */

interface EditorToolbarComponent {
    (props: EditorToolbarProps & {
        children: ReactNode;
    }): React.JSX.Element;
    TitleBar: typeof TitleBar;
    Logo: typeof Logo;
    DocumentName: typeof DocumentName;
    MenuBar: typeof MenuBar;
    TitleBarRight: typeof TitleBarRight;
    Toolbar: typeof Toolbar;
}
declare const EditorToolbar: EditorToolbarComponent;

/**
 * Context Menu Component
 *
 * Right-click context menu for AI actions on selected text.
 * Shows AI options like rewrite, expand, summarize, translate, etc.
 */

/**
 * Context menu props
 */
interface ContextMenuProps {
    /** Whether the menu is visible */
    isOpen: boolean;
    /** Menu position */
    position: {
        x: number;
        y: number;
    };
    /** Selected text */
    selectedText: string;
    /** Selection context for AI operations */
    selectionContext?: SelectionContext;
    /** Callback when an action is selected */
    onAction: (action: AIAction, customPrompt?: string) => void;
    /** Callback when menu is closed */
    onClose: () => void;
    /** Available actions (defaults to DEFAULT_AI_ACTIONS) */
    actions?: AIAction[];
    /** Whether to show custom prompt option */
    showCustomPrompt?: boolean;
    /** Additional className */
    className?: string;
}
declare const ContextMenu: React__default.FC<ContextMenuProps>;
/**
 * Hook to manage context menu state
 */
declare function useContextMenu(): {
    isOpen: boolean;
    position: {
        x: number;
        y: number;
    };
    selectedText: string;
    selectionContext: SelectionContext | undefined;
    openMenu: (e: React__default.MouseEvent | MouseEvent, text: string, context?: SelectionContext) => void;
    closeMenu: () => void;
};
/**
 * Get action shortcuts
 */
declare function getActionShortcut(action: AIAction): string | undefined;
/**
 * Check if action is available for selection
 */
declare function isActionAvailable(_action: AIAction, selectedText: string, _selectionContext?: SelectionContext): boolean;
/**
 * Get default actions for selection
 */
declare function getDefaultActions(): AIAction[];
/**
 * Get all available actions
 */
declare function getAllActions(): AIAction[];

/**
 * Response Preview Component
 *
 * Shows AI response preview with diff view before applying changes.
 * Allows user to accept, reject, or edit the response.
 */

/**
 * Response preview props
 */
interface ResponsePreviewProps {
    /** Original selected text */
    originalText: string;
    /** AI response (or null if loading/error) */
    response: AgentResponse | null;
    /** Action that was performed */
    action: AIAction;
    /** Whether the response is loading */
    isLoading: boolean;
    /** Error message if request failed */
    error?: string;
    /** Callback when user accepts the change */
    onAccept: (newText: string) => void;
    /** Callback when user rejects the change */
    onReject: () => void;
    /** Callback when user wants to retry */
    onRetry?: () => void;
    /** Allow editing before accepting */
    allowEdit?: boolean;
    /** Show diff view */
    showDiff?: boolean;
    /** Additional className */
    className?: string;
    /** Position for the preview */
    position?: {
        x: number;
        y: number;
    };
}
declare const ResponsePreview: React__default.FC<ResponsePreviewProps>;
/**
 * State for response preview
 */
interface ResponsePreviewState {
    isVisible: boolean;
    originalText: string;
    response: AgentResponse | null;
    action: AIAction;
    isLoading: boolean;
    error?: string;
    position?: {
        x: number;
        y: number;
    };
}
/**
 * Hook to manage response preview state
 */
declare function useResponsePreview(): {
    state: ResponsePreviewState;
    showPreview: (originalText: string, action: AIAction, position?: {
        x: number;
        y: number;
    }) => void;
    setResponse: (response: AgentResponse) => void;
    setError: (error: string) => void;
    hidePreview: () => void;
};
/**
 * Create a mock response for testing
 */
declare function createMockResponse(newText: string, warnings?: string[]): AgentResponse;
/**
 * Create an error response
 */
declare function createErrorResponse(error: string): AgentResponse;

/**
 * Text Context Menu Component
 *
 * Right-click context menu for text editing operations.
 * Shows Cut, Copy, Paste, and other text editing options.
 */

/**
 * Context menu action types
 */
type TextContextAction = 'cut' | 'copy' | 'paste' | 'pasteAsPlainText' | 'selectAll' | 'delete' | 'separator' | 'addRowAbove' | 'addRowBelow' | 'deleteRow' | 'addColumnLeft' | 'addColumnRight' | 'deleteColumn' | 'mergeCells' | 'splitCell' | 'selectTable' | 'deleteTable' | 'addComment';
/**
 * Menu item configuration
 */
interface TextContextMenuItem {
    /** Action type */
    action: TextContextAction;
    /** Display label */
    label: string;
    /** Keyboard shortcut hint */
    shortcut?: string;
    /** Whether the item is disabled */
    disabled?: boolean;
    /** Whether to show divider after this item */
    dividerAfter?: boolean;
}
/**
 * Context menu props
 */
interface TextContextMenuProps {
    /** Whether the menu is visible */
    isOpen: boolean;
    /** Menu position */
    position: {
        x: number;
        y: number;
    };
    /** Whether there's a selection (enables copy/cut) */
    hasSelection: boolean;
    /** Whether the editor is editable (enables paste/cut/delete) */
    isEditable: boolean;
    /** Whether clipboard has content (enables paste) */
    hasClipboardContent?: boolean;
    /** Callback when an action is selected */
    onAction: (action: TextContextAction) => void;
    /** Callback when menu is closed */
    onClose: () => void;
    /** Custom menu items (overrides default) */
    items?: TextContextMenuItem[];
    /** Additional className */
    className?: string;
}
/**
 * Hook options for text context menu
 */
interface UseTextContextMenuOptions {
    /** Whether the context menu is enabled */
    enabled?: boolean;
    /** Whether the editor is editable */
    isEditable?: boolean;
    /** Container element ref */
    containerRef?: React__default.RefObject<HTMLElement>;
    /** Callback when an action is triggered */
    onAction?: (action: TextContextAction) => void;
}
/**
 * Hook return value
 */
interface UseTextContextMenuReturn {
    /** Whether the menu is open */
    isOpen: boolean;
    /** Menu position */
    position: {
        x: number;
        y: number;
    };
    /** Whether there's a text selection */
    hasSelection: boolean;
    /** Open the context menu */
    openMenu: (event: React__default.MouseEvent | MouseEvent) => void;
    /** Close the context menu */
    closeMenu: () => void;
    /** Handle action selection */
    handleAction: (action: TextContextAction) => void;
    /** Context menu event handler for onContextMenu prop */
    onContextMenu: (event: React__default.MouseEvent) => void;
}
declare const TextContextMenu: React__default.FC<TextContextMenuProps>;
/**
 * Hook to manage text context menu state
 */
declare function useTextContextMenu(options?: UseTextContextMenuOptions): UseTextContextMenuReturn;
/**
 * Get action label
 */
declare function getTextActionLabel(action: TextContextAction): string;
/**
 * Get action shortcut
 */
declare function getTextActionShortcut(action: TextContextAction): string;
/**
 * Get default menu item definitions (untranslated, use translation keys)
 */
declare function getDefaultTextContextMenuItems(): TextContextMenuItem[];
/**
 * Check if action is available
 */
declare function isTextActionAvailable(action: TextContextAction, hasSelection: boolean, isEditable: boolean): boolean;

/**
 * Zoom Control Component
 *
 * A − / + stepper around a Radix Select: the buttons step through the zoom
 * levels and the middle % opens a dropdown for a direct pick. Matches the Vue
 * toolbar's zoom stepper and the font-size stepper.
 */

interface ZoomLevel {
    value: number;
    label: string;
}
interface ZoomControlProps {
    value?: number;
    onChange?: (zoom: number) => void;
    levels?: ZoomLevel[];
    disabled?: boolean;
    className?: string;
    compact?: boolean;
}
declare function ZoomControl({ value, onChange, levels, disabled, className, compact, }: ZoomControlProps): React$1.JSX.Element;

/**
 * Font Size Picker Component (Google Docs Style)
 *
 * A font size control with minus/plus buttons and editable input.
 * Features:
 * - Minus button to decrease font size
 * - Plus button to increase font size
 * - Editable input for custom sizes
 * - Click input to show dropdown with preset sizes
 */

interface FontSizePickerProps {
    value?: number;
    onChange?: (size: number) => void;
    sizes?: number[];
    disabled?: boolean;
    className?: string;
    placeholder?: string;
    width?: number | string;
    minSize?: number;
    maxSize?: number;
}

declare function FontSizePicker({ value, onChange, sizes, disabled, className, placeholder, minSize, maxSize, }: FontSizePickerProps): React$1.JSX.Element;

/**
 * Line Spacing Picker Component (Radix UI)
 *
 * A dropdown selector for choosing line spacing values using Radix Select.
 * Styled like Google Docs with options: Single, 1.15, 1.5, Double
 */

interface LineSpacingOption {
    label: string;
    labelKey?: TranslationKey;
    value: number;
    twipsValue: number;
}
interface LineSpacingPickerProps {
    value?: number;
    onChange?: (twipsValue: number) => void;
    options?: LineSpacingOption[];
    disabled?: boolean;
    className?: string;
    width?: number | string;
}
declare function LineSpacingPicker({ value, onChange, options, disabled, className, }: LineSpacingPickerProps): React$1.JSX.Element;

type ColorPickerMode = 'text' | 'highlight' | 'border';
interface ColorPickerProps {
    mode: ColorPickerMode;
    value?: ColorValue | string;
    onChange?: (color: ColorValue | string) => void;
    theme?: Theme | null;
    disabled?: boolean;
    className?: string;
    style?: CSSProperties;
    title?: string;
    /** Override the default icon for the mode */
    icon?: string;
    /** Override the auto/no-color button label */
    autoLabel?: string;
    /**
     * Word-style split button. When true (default), renders two halves:
     *  - left (apply): re-applies the last picked color directly
     *  - right (arrow): opens the color picker
     * When false, the legacy single button is rendered (one click toggles dropdown).
     */
    splitButton?: boolean;
    /**
     * Initial "last picked" color used by the apply half before the user picks
     * anything. Defaults: text → red, highlight → yellow, border → black.
     */
    defaultColor?: ColorValue | string;
}
declare function ColorPicker({ mode, value, onChange, theme, disabled, className, style, title, icon: iconOverride, autoLabel, splitButton, defaultColor, }: ColorPickerProps): React$1.JSX.Element;

/**
 * Style Picker Component (Radix UI)
 *
 * A dropdown selector for applying named paragraph styles using Radix Select.
 * Shows each style with its visual appearance (font size, bold, color).
 */

interface StyleOption {
    styleId: string;
    name: string;
    nameKey?: TranslationKey;
    type: StyleType;
    isDefault?: boolean;
    qFormat?: boolean;
    priority?: number;
    /** Font size in half-points for visual preview */
    fontSize?: number;
    /** Bold styling */
    bold?: boolean;
    /** Italic styling */
    italic?: boolean;
    /** Text color (RGB hex) */
    color?: string;
}
interface StylePickerProps {
    value?: string;
    onChange?: (styleId: string) => void;
    styles?: Style[];
    theme?: Theme | null;
    disabled?: boolean;
    className?: string;
    width?: number | string;
}
declare function StylePicker({ value, onChange, styles, disabled, className, width, }: StylePickerProps): React$1.JSX.Element;

/**
 * Alignment Dropdown Component (Google Docs style)
 *
 * A single dropdown button for paragraph alignment controls:
 * - Shows current alignment icon + chevron
 * - Opens a floating panel with Left, Center, Right, Justify options
 * - Active option is highlighted
 */

/**
 * Props for the AlignmentButtons component
 */
interface AlignmentButtonsProps {
    /** Current alignment value */
    value?: ParagraphAlignment;
    /** Callback when alignment is changed */
    onChange?: (alignment: ParagraphAlignment) => void;
    /** Whether the buttons are disabled */
    disabled?: boolean;
    /** Additional CSS class name */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Show labels next to icons */
    showLabels?: boolean;
    /** Compact mode (smaller buttons) */
    compact?: boolean;
}
/**
 * Alignment dropdown component — single button with popover panel
 */
declare function AlignmentButtons({ value, onChange, disabled, }: AlignmentButtonsProps): React__default.JSX.Element;

/**
 * HorizontalRuler Component — Google Docs style
 *
 * 3 handles only:
 * - Left side: first-line indent (▼ down at top) + left indent (▲ up at bottom)
 * - Right side: right indent (▼ down at top)
 *
 * Margins shown as gray zones on the ruler edges.
 * Drag the boundary between gray/white to adjust page margins.
 * Drag tooltip shows value during any drag.
 */

interface HorizontalRulerProps {
    sectionProps?: SectionProperties | null;
    zoom?: number;
    editable?: boolean;
    onLeftMarginChange?: (marginTwips: number) => void;
    onRightMarginChange?: (marginTwips: number) => void;
    onFirstLineIndentChange?: (indentTwips: number) => void;
    showFirstLineIndent?: boolean;
    firstLineIndent?: number;
    hangingIndent?: boolean;
    indentLeft?: number;
    indentRight?: number;
    onIndentLeftChange?: (indentTwips: number) => void;
    onIndentRightChange?: (indentTwips: number) => void;
    unit?: 'inch' | 'cm';
    className?: string;
    style?: CSSProperties;
    tabStops?: TabStop[] | null;
    onTabStopRemove?: (positionTwips: number) => void;
}
declare function HorizontalRuler({ sectionProps, zoom, editable, onLeftMarginChange, onRightMarginChange, onFirstLineIndentChange, showFirstLineIndent, firstLineIndent, hangingIndent, indentLeft, indentRight, onIndentLeftChange, onIndentRightChange, unit, className, style, tabStops, onTabStopRemove, }: HorizontalRulerProps): React__default.ReactElement;
declare function positionToMargin(positionPx: number, side: 'left' | 'right', pageWidthPx: number, zoom: number): number;
declare function getRulerDimensions(sectionProps?: SectionProperties | null, zoom?: number): {
    width: number;
    leftMargin: number;
    rightMargin: number;
    contentWidth: number;
};
declare function getMarginInUnits(marginTwips: number, unit: 'inch' | 'cm'): string;
declare function parseMarginFromUnits(value: string, unit: 'inch' | 'cm'): number | null;

interface TableBorderPickerProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
}
declare function TableBorderPicker({ onAction, disabled }: TableBorderPickerProps): React$1.JSX.Element;

interface TableBorderColorPickerProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
    theme?: Theme | null;
    /** Current border color (RGB hex without #) */
    value?: string;
}
declare function TableBorderColorPicker({ onAction, disabled, theme, value, }: TableBorderColorPickerProps): React$1.JSX.Element;

interface TableBorderWidthPickerProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
}
declare function TableBorderWidthPicker({ onAction, disabled, }: TableBorderWidthPickerProps): React$1.JSX.Element;

interface TableCellFillPickerProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
    theme?: Theme | null;
    /** Current fill color (RGB hex without #) */
    value?: string;
}
declare function TableCellFillPicker({ onAction, disabled, theme, value, }: TableCellFillPickerProps): React$1.JSX.Element;

/**
 * TableMergeButton - Toggle button for merge/split cells
 */

interface TableMergeButtonProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
    canMerge?: boolean;
    canSplit?: boolean;
}
declare function TableMergeButton({ onAction, disabled, canMerge, canSplit, }: TableMergeButtonProps): React__default.JSX.Element;

/**
 * TableInsertButtons - 4 icon buttons for row/column insertion
 *
 * Insert row above, insert row below, insert column left, insert column right.
 */

interface TableInsertButtonsProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
}
declare function TableInsertButtons({ onAction, disabled }: TableInsertButtonsProps): React__default.JSX.Element;

interface TableMoreDropdownProps {
    onAction: (action: TableAction) => void;
    disabled?: boolean;
    tableContext?: {
        isInTable: boolean;
        rowCount?: number;
        columnCount?: number;
        canSplitCell?: boolean;
        hasMultiCellSelection?: boolean;
        table?: {
            attrs?: {
                justification?: string;
            };
        };
    } | null;
}
declare function TableMoreDropdown({ onAction, disabled, tableContext, }: TableMoreDropdownProps): React$1.JSX.Element;

/**
 * Unsaved Changes Indicator Component
 *
 * Visual indicator that shows when document has unsaved changes.
 * Features:
 * - Configurable appearance (dot, badge, text)
 * - Pulse animation option for visibility
 * - Hook for tracking changes
 * - Browser beforeunload warning
 */

/**
 * Indicator variant type
 */
type IndicatorVariant = 'dot' | 'badge' | 'text' | 'icon';
/**
 * Indicator position type
 */
type IndicatorPosition = 'inline' | 'absolute-top-right' | 'absolute-top-left';
/**
 * Unsaved indicator props
 */
interface UnsavedIndicatorProps {
    /** Whether there are unsaved changes */
    hasUnsavedChanges: boolean;
    /** Variant of the indicator */
    variant?: IndicatorVariant;
    /** Position of the indicator */
    position?: IndicatorPosition;
    /** Whether to show pulse animation */
    showPulse?: boolean;
    /** Custom label for text variant */
    label?: string;
    /** Custom saved label for text variant */
    savedLabel?: string;
    /** Whether to show indicator when saved (always show) */
    showWhenSaved?: boolean;
    /** Custom color for unsaved state */
    unsavedColor?: string;
    /** Custom color for saved state */
    savedColor?: string;
    /** Size in pixels (for dot/icon) */
    size?: number;
    /** Additional className */
    className?: string;
    /** Additional inline styles */
    style?: React__default.CSSProperties;
    /** Click handler */
    onClick?: () => void;
    /** Title/tooltip text */
    title?: string;
}
/**
 * Hook options for tracking unsaved changes
 */
interface UseUnsavedChangesOptions {
    /** The document to track */
    document?: Document | null;
    /** Whether to warn before leaving page */
    warnBeforeLeave?: boolean;
    /** Custom warning message */
    warningMessage?: string;
    /** Whether tracking is enabled */
    enabled?: boolean;
    /** Callback when changes status changes */
    onChangeStatusChange?: (hasChanges: boolean) => void;
}
/**
 * Hook return value
 */
interface UseUnsavedChangesReturn {
    /** Whether there are unsaved changes */
    hasUnsavedChanges: boolean;
    /** Mark the document as saved (resets change tracking) */
    markAsSaved: () => void;
    /** Mark the document as changed */
    markAsChanged: () => void;
    /** Reset tracking (resets baseline) */
    resetTracking: (newDocument?: Document | null) => void;
    /** The last saved document snapshot */
    lastSavedDocument: Document | null;
    /** Number of changes since last save */
    changeCount: number;
}
declare const UnsavedIndicator: React__default.FC<UnsavedIndicatorProps>;
/**
 * Hook to track unsaved changes in a document
 */
declare function useUnsavedChanges(options?: UseUnsavedChangesOptions): UseUnsavedChangesReturn;
/**
 * Get indicator variant label
 */
declare function getVariantLabel(variant: IndicatorVariant): string;
/**
 * Get all indicator variants
 */
declare function getAllVariants(): IndicatorVariant[];
/**
 * Get all indicator positions
 */
declare function getAllPositions(): IndicatorPosition[];
/**
 * Create a document change tracker
 * Simple utility for external change tracking
 */
declare function createChangeTracker(): {
    markChanged: () => void;
    markSaved: () => void;
    getState: () => {
        changeCount: number;
        lastSaveTime: Date | null;
        hasUnsavedChanges: boolean;
    };
    reset: () => void;
};

/**
 * Loading Indicator Component
 *
 * Displays loading states for operations with configurable appearance.
 * Features:
 * - Multiple spinner styles (spinner, dots, bar, pulse)
 * - Overlay mode for blocking UI during operations
 * - Inline mode for subtle loading indication
 * - Progress bar variant
 * - Hook for managing loading states
 */

/**
 * Loading indicator variant
 */
type LoadingVariant = 'spinner' | 'dots' | 'bar' | 'pulse' | 'progress';
/**
 * Loading indicator size
 */
type LoadingSize = 'small' | 'medium' | 'large';
/**
 * Loading indicator props
 */
interface LoadingIndicatorProps {
    /** Whether loading is active */
    isLoading: boolean;
    /** Variant of the loading indicator */
    variant?: LoadingVariant;
    /** Size of the indicator */
    size?: LoadingSize;
    /** Loading message to display */
    message?: string;
    /** Whether to show as full-screen overlay */
    overlay?: boolean;
    /** Overlay background opacity (0-1) */
    overlayOpacity?: number;
    /** Progress percentage (0-100) for progress variant */
    progress?: number;
    /** Show progress percentage text */
    showProgressText?: boolean;
    /** Custom color */
    color?: string;
    /** Additional className */
    className?: string;
    /** Additional inline styles */
    style?: React__default.CSSProperties;
}
/**
 * Options for useLoading hook
 */
interface UseLoadingOptions {
    /** Initial loading state */
    initialLoading?: boolean;
    /** Minimum loading duration in ms (prevents flash) */
    minDuration?: number;
    /** Callback when loading starts */
    onStart?: () => void;
    /** Callback when loading ends */
    onEnd?: () => void;
}
/**
 * Return value of useLoading hook
 */
interface UseLoadingReturn {
    /** Current loading state */
    isLoading: boolean;
    /** Current message */
    message: string | null;
    /** Current progress (0-100) */
    progress: number;
    /** Start loading with optional message */
    startLoading: (message?: string) => void;
    /** Stop loading */
    stopLoading: () => void;
    /** Update progress (0-100) */
    setProgress: (progress: number) => void;
    /** Update message */
    setMessage: (message: string | null) => void;
    /** Wrap an async operation with loading state */
    withLoading: <T>(operation: () => Promise<T>, message?: string) => Promise<T>;
}
/**
 * Loading operation state
 */
interface LoadingOperation {
    id: string;
    message?: string;
    progress?: number;
    startTime: number;
}
declare const LoadingIndicator: React__default.FC<LoadingIndicatorProps>;
/**
 * Hook to manage loading states
 */
declare function useLoading(options?: UseLoadingOptions): UseLoadingReturn;
/**
 * Hook to manage multiple concurrent loading operations
 */
declare function useLoadingOperations(): {
    operations: LoadingOperation[];
    isAnyLoading: boolean;
    startOperation: (id: string, message?: string) => void;
    updateOperation: (id: string, updates: Partial<LoadingOperation>) => void;
    endOperation: (id: string) => void;
    getOperation: (id: string) => LoadingOperation | undefined;
};
/**
 * Get loading variant label
 */
declare function getLoadingVariantLabel(variant: LoadingVariant): string;
/**
 * Get all loading variants
 */
declare function getAllLoadingVariants(): LoadingVariant[];
/**
 * Get all loading sizes
 */
declare function getAllLoadingSizes(): LoadingSize[];
/**
 * Create a delay promise for testing loading states
 */
declare function delay(ms: number): Promise<void>;

/**
 * Responsive Toolbar Component
 *
 * A responsive toolbar wrapper that collapses items into an overflow menu
 * when the screen is narrow.
 *
 * Features:
 * - Automatically measures available space
 * - Moves items to overflow menu when they don't fit
 * - Priority-based item ordering
 * - Configurable breakpoints
 * - ResizeObserver for dynamic resizing
 */

/**
 * Priority level for toolbar items
 * Lower numbers = higher priority (shown first, hidden last)
 */
type ToolbarItemPriority = 1 | 2 | 3 | 4 | 5;
/**
 * Toolbar item configuration
 */
interface ToolbarItem {
    /** Unique identifier */
    id: string;
    /** The content to render */
    content: ReactNode;
    /** Priority level (1 = highest, 5 = lowest) */
    priority?: ToolbarItemPriority;
    /** Minimum width in pixels (for measuring) */
    minWidth?: number;
    /** Whether this item should never be hidden */
    alwaysVisible?: boolean;
    /** Whether to show separator after this item */
    separatorAfter?: boolean;
    /** Group name for keeping items together */
    group?: string;
}
/**
 * Props for ResponsiveToolbar component
 */
interface ResponsiveToolbarProps {
    /** Toolbar items */
    items: ToolbarItem[];
    /** Additional items for overflow menu only */
    overflowItems?: ToolbarItem[];
    /** Whether to show overflow button even when all items fit */
    alwaysShowOverflow?: boolean;
    /** Custom overflow button renderer */
    renderOverflowButton?: (itemCount: number, isOpen: boolean, onClick: () => void) => ReactNode;
    /** Custom overflow menu renderer */
    renderOverflowMenu?: (items: ToolbarItem[], onClose: () => void) => ReactNode;
    /** Gap between items in pixels */
    itemGap?: number;
    /** Padding for the toolbar */
    padding?: number | string;
    /** Minimum width for overflow button */
    overflowButtonWidth?: number;
    /** Additional className */
    className?: string;
    /** Additional styles */
    style?: CSSProperties;
    /** Height of the toolbar */
    height?: number | string;
    /** Background color */
    backgroundColor?: string;
    /** Border styles */
    borderBottom?: string;
}
/**
 * Options for useResponsiveToolbar hook
 */
interface UseResponsiveToolbarOptions {
    /** Container ref */
    containerRef: React__default.RefObject<HTMLElement | null>;
    /** Total items */
    items: ToolbarItem[];
    /** Gap between items */
    itemGap?: number;
    /** Width reserved for overflow button */
    overflowButtonWidth?: number;
}
/**
 * Return value of useResponsiveToolbar hook
 */
interface UseResponsiveToolbarReturn {
    /** Items that fit in visible area */
    visibleItems: ToolbarItem[];
    /** Items in overflow menu */
    overflowItems: ToolbarItem[];
    /** Whether overflow menu is needed */
    hasOverflow: boolean;
    /** Force a recalculation */
    recalculate: () => void;
}
/**
 * Hook to calculate which items fit in the toolbar
 */
declare function useResponsiveToolbar(options: UseResponsiveToolbarOptions): UseResponsiveToolbarReturn;
declare const ResponsiveToolbar: React__default.FC<ResponsiveToolbarProps>;
interface ToolbarGroupProps {
    /** Group items */
    children: ReactNode;
    /** Gap between items */
    gap?: number;
    /** Whether to show separator after group */
    separatorAfter?: boolean;
    /** Additional className */
    className?: string;
    /** Additional styles */
    style?: CSSProperties;
}
declare const ToolbarGroup: React__default.FC<ToolbarGroupProps>;
/**
 * Create a toolbar item
 */
declare function createToolbarItem(id: string, content: ReactNode, options?: Partial<Omit<ToolbarItem, 'id' | 'content'>>): ToolbarItem;
/**
 * Create toolbar items from an array of configs
 */
declare function createToolbarItems(configs: Array<{
    id: string;
    content: ReactNode;
    priority?: ToolbarItemPriority;
    minWidth?: number;
    alwaysVisible?: boolean;
    separatorAfter?: boolean;
    group?: string;
}>): ToolbarItem[];
/**
 * Get recommended priority for common toolbar items
 */
declare function getRecommendedPriority(itemType: string): ToolbarItemPriority;

/**
 * Insert Table Dialog Component
 *
 * Modal dialog for inserting a new table into the document.
 * Provides a visual grid selector for choosing rows and columns.
 *
 * Features:
 * - Visual grid selector (hover to select dimensions)
 * - Manual row/column input
 * - Preview of table dimensions
 * - Quick insert with default sizes
 */

/**
 * Table configuration for insertion
 */
interface TableConfig {
    /** Number of rows */
    rows: number;
    /** Number of columns */
    columns: number;
}
/**
 * Props for InsertTableDialog
 */
interface InsertTableDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when table is inserted */
    onInsert: (config: TableConfig) => void;
    /** Maximum rows in grid selector (default: 8) */
    maxGridRows?: number;
    /** Maximum columns in grid selector (default: 10) */
    maxGridColumns?: number;
    /** Maximum allowed rows (default: 100) */
    maxRows?: number;
    /** Maximum allowed columns (default: 20) */
    maxColumns?: number;
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
}
/**
 * InsertTableDialog - Modal for inserting tables with visual grid selector
 */
declare function InsertTableDialog({ isOpen, onClose, onInsert, maxGridRows, maxGridColumns, maxRows, maxColumns, className, style, }: InsertTableDialogProps): React__default.ReactElement | null;
/**
 * Hook for managing Insert Table dialog state
 */
declare function useInsertTableDialog(): {
    isOpen: boolean;
    open: () => void;
    close: () => void;
    toggle: () => void;
};
/**
 * Create a default TableConfig
 */
declare function createDefaultTableConfig(rows?: number, columns?: number): TableConfig;
/**
 * Validate TableConfig
 */
declare function isValidTableConfig(config: TableConfig, maxRows?: number, maxColumns?: number): boolean;
/**
 * Clamp TableConfig to valid range
 */
declare function clampTableConfig(config: TableConfig, maxRows?: number, maxColumns?: number): TableConfig;
/**
 * Format table dimensions for display
 */
declare function formatTableDimensions(config: TableConfig): string;
/**
 * Get common table presets
 */
declare function getTablePresets(): {
    label: string;
    config: TableConfig;
}[];

/**
 * Insert Image Dialog Component
 *
 * Modal dialog for inserting images into the document.
 * Supports file upload with preview and basic sizing options.
 *
 * Features:
 * - File input for image selection
 * - Drag and drop support
 * - Image preview
 * - Width/height controls with aspect ratio lock
 * - Alt text input
 */

/**
 * Image data for insertion
 */
interface ImageData {
    /** Base64 data URL or external URL */
    src: string;
    /** Image width in pixels */
    width: number;
    /** Image height in pixels */
    height: number;
    /** Alt text for accessibility */
    alt?: string;
    /** Original file name */
    fileName?: string;
    /** MIME type */
    mimeType?: string;
}
/**
 * Props for InsertImageDialog
 */
interface InsertImageDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when image is inserted */
    onInsert: (data: ImageData) => void;
    /** Maximum width in pixels (default: 800) */
    maxWidth?: number;
    /** Maximum height in pixels (default: 600) */
    maxHeight?: number;
    /** Accepted file types (default: image/*) */
    accept?: string;
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
}
/**
 * InsertImageDialog - Modal for inserting images with preview and sizing
 */
declare function InsertImageDialog({ isOpen, onClose, onInsert, maxWidth, maxHeight, accept, className, style, }: InsertImageDialogProps): React__default.ReactElement | null;
/**
 * Hook for managing Insert Image dialog state
 */
declare function useInsertImageDialog(): {
    isOpen: boolean;
    open: () => void;
    close: () => void;
    toggle: () => void;
};
/**
 * Check if a file is a valid image
 */
declare function isValidImageFile(file: File): boolean;
/**
 * Get supported image extensions
 */
declare function getSupportedImageExtensions(): string[];
/**
 * Get accept string for file input
 */
declare function getImageAcceptString(): string;
/**
 * Calculate scaled dimensions to fit within bounds
 */
declare function calculateFitDimensions(originalWidth: number, originalHeight: number, maxWidth: number, maxHeight: number): {
    width: number;
    height: number;
};
/**
 * Convert data URL to Blob
 */
declare function dataUrlToBlob(dataUrl: string): Blob;
/**
 * Get image dimensions from a data URL
 */
declare function getImageDimensions(src: string): Promise<{
    width: number;
    height: number;
}>;
/**
 * Format file size for display
 */
declare function formatFileSize(bytes: number): string;

/**
 * Insert Symbol Dialog Component
 *
 * Modal dialog for inserting special characters and symbols into the document.
 * Provides categorized symbol picker with search functionality.
 *
 * Features:
 * - Categorized symbol groups
 * - Recent symbols
 * - Search functionality
 * - Unicode character display
 */

/**
 * Symbol category
 */
interface SymbolCategory {
    /** Category name */
    name: string;
    /** Display label */
    label: string;
    /** Symbols in this category */
    symbols: string[];
}
/**
 * Props for InsertSymbolDialog
 */
interface InsertSymbolDialogProps {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Callback when dialog is closed */
    onClose: () => void;
    /** Callback when symbol is inserted */
    onInsert: (symbol: string) => void;
    /** Recently used symbols */
    recentSymbols?: string[];
    /** Additional CSS class */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
}
/**
 * Default symbol categories
 */
declare const SYMBOL_CATEGORIES: SymbolCategory[];
/**
 * InsertSymbolDialog - Modal for inserting special characters
 */
declare function InsertSymbolDialog({ isOpen, onClose, onInsert, recentSymbols, className, style, }: InsertSymbolDialogProps): React__default.ReactElement | null;
/**
 * Hook for managing Insert Symbol dialog state with recent symbols
 */
declare function useInsertSymbolDialog(maxRecent?: number): {
    isOpen: boolean;
    recentSymbols: string[];
    open: () => void;
    close: () => void;
    toggle: () => void;
    addRecent: (symbol: string) => void;
};
/**
 * Get all symbol categories
 */
declare function getSymbolCategories(): SymbolCategory[];
/**
 * Get symbols by category name
 */
declare function getSymbolsByCategory(categoryName: string): string[];
/**
 * Get symbol Unicode info
 */
declare function getSymbolInfo(symbol: string): {
    character: string;
    codePoint: string;
    decimal: number;
    hex: string;
};
/**
 * Search symbols by query
 */
declare function searchSymbols(query: string): string[];
/**
 * Get symbol from Unicode code point string
 */
declare function symbolFromCodePoint(codePointStr: string): string | null;

export { AlignmentButtons, type AlignmentButtonsProps, ColorPicker, type ColorPickerMode, type ColorPickerProps, ContextMenu, type ContextMenuProps, type DocumentNameProps, EditorToolbar, type EditorToolbarProps, FontPicker, type FontPickerProps, FontSizePicker, type FontSizePickerProps, HorizontalRuler, type HorizontalRulerProps, type ImageData, type IndicatorPosition, type IndicatorVariant, InsertImageDialog, type InsertImageDialogProps, InsertSymbolDialog, type InsertSymbolDialogProps, InsertTableDialog, type InsertTableDialogProps, type LineSpacingOption, LineSpacingPicker, type LineSpacingPickerProps, ListButtons, type ListButtonsProps, LoadingIndicator, type LoadingIndicatorProps, type LoadingOperation, type LoadingSize, type LoadingVariant, type LogoProps, PrintButton, type PrintButtonProps, PrintStyles, ResponsePreview, type ResponsePreviewProps, type ResponsePreviewState, ResponsiveToolbar, ToolbarGroup as ResponsiveToolbarGroup, type ToolbarGroupProps as ResponsiveToolbarGroupProps, type ResponsiveToolbarProps, SYMBOL_CATEGORIES, type StyleOption, StylePicker, type StylePickerProps, type SymbolCategory, TableAction, TableBorderColorPicker, type TableBorderColorPickerProps, TableBorderPicker, type TableBorderPickerProps, TableBorderWidthPicker, type TableBorderWidthPickerProps, TableCellFillPicker, type TableCellFillPickerProps, type TableConfig, TableContext, TableInsertButtons, type TableInsertButtonsProps, TableMergeButton, type TableMergeButtonProps, TableMoreDropdown, type TableMoreDropdownProps, TableSelection, TableSplitConfig, type TextContextAction, TextContextMenu, type TextContextMenuItem, type TextContextMenuProps, type TitleBarProps, type TitleBarRightProps, Toolbar, ToolbarButton, ToolbarGroup$1 as ToolbarGroup, type ToolbarItem, type ToolbarItemPriority, type ToolbarProps, ToolbarSeparator, UnsavedIndicator, type UnsavedIndicatorProps, type UseLoadingOptions, type UseLoadingReturn, type UseResponsiveToolbarOptions, type UseResponsiveToolbarReturn, type UseTextContextMenuOptions, type UseTextContextMenuReturn, type UseUnsavedChangesOptions, type UseUnsavedChangesReturn, ZoomControl, type ZoomControlProps, addColumn, addRow, calculateFitDimensions, clampTableConfig, createChangeTracker, createDefaultTableConfig, createErrorResponse, createMockResponse, createTableContext, createToolbarItem, createToolbarItems, dataUrlToBlob, delay, deleteColumn, deleteRow, formatFileSize, formatTableDimensions, getActionShortcut, getAllActions, getAllPositions as getAllIndicatorPositions, getAllVariants as getAllIndicatorVariants, getAllLoadingSizes, getAllLoadingVariants, getCellAt, getColumnCount, getDefaultActions, getDefaultTextContextMenuItems, getImageAcceptString, getImageDimensions, getLoadingVariantLabel, getMarginInUnits, getRecommendedPriority, getRulerDimensions, getSupportedImageExtensions, getSymbolCategories, getSymbolInfo as getSymbolUnicodeInfo, getSymbolsByCategory, getTablePresets, getTableSplitCellDialogConfig, getTextActionLabel, getTextActionShortcut, getVariantLabel, isActionAvailable, isTextActionAvailable, isValidImageFile, isValidTableConfig, mergeCells, parseMarginFromUnits, positionToMargin, searchSymbols, splitCell, splitTableCell, symbolFromCodePoint, useContextMenu, useInsertImageDialog, useInsertSymbolDialog, useInsertTableDialog, useLoading, useLoadingOperations, useResponsePreview, useResponsiveToolbar, useTextContextMenu, useUnsavedChanges };
