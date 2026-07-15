/**
 * @eigenpal/docx-editor-react/hooks
 *
 * React hooks for editor history, table selection, find/replace, autosave,
 * clipboard, and zoom. Use alongside the main `DocxEditor` component.
 *
 * @example
 * ```tsx
 * import { useAutoSave, useFindReplace } from '@eigenpal/docx-editor-react/hooks';
 * ```
 *
 * @packageDocumentation
 * @public
 */
import { Table, Document } from '@eigenpal/docx-editor-core/types/document';
import { a as TableContext, c as TableAction, b as TableSplitConfig } from './useFindReplace-Bc2ubEeV.js';
export { F as FindReplaceOptions, d as FindReplaceState, U as UseFindReplaceReturn, u as useFindReplace } from './useFindReplace-Bc2ubEeV.js';
import * as React$1 from 'react';
import React__default, { CSSProperties, RefObject } from 'react';
import { HighlightRect, SelectionHighlightConfig, ParsedClipboardContent } from '@eigenpal/docx-editor-core/utils';
import { ClipboardSelection, Theme, SavedDocumentData, AutoSaveStatus } from '@eigenpal/docx-editor-core';
export { AutoSaveStatus, ClipboardSelection, SavedDocumentData, TABLE_DATA_ATTRIBUTES, createSelectionFromDOM, formatLastSaveTime, formatStorageSize, getAutoSaveStatusLabel, getAutoSaveStorageSize, getSelectionRuns, isAutoSaveSupported } from '@eigenpal/docx-editor-core';
import { EditorView } from 'prosemirror-view';
import { EditorState } from 'prosemirror-state';
import { TrackedChangesResult } from '@eigenpal/docx-editor-core/prosemirror/utils/extractTrackedChanges';
export { TrackedChangesResult, extractTrackedChanges } from '@eigenpal/docx-editor-core/prosemirror/utils/extractTrackedChanges';
import '@eigenpal/docx-editor-core/utils/findReplace';

/**
 * History hook for undo/redo functionality
 *
 * Maintains undo/redo stacks with support for:
 * - undo() and redo() operations
 * - canUndo and canRedo state
 * - Keyboard shortcuts (Ctrl+Z, Ctrl+Y, Ctrl+Shift+Z)
 * - Grouping rapid changes to avoid cluttering history
 */
/**
 * History entry containing state and metadata
 */
interface HistoryEntry<T> {
    /** The state at this point */
    state: T;
    /** Timestamp when this entry was created */
    timestamp: number;
    /** Optional description of what changed */
    description?: string;
}
/**
 * Options for the useHistory hook
 */
interface UseHistoryOptions<T> {
    /** Maximum number of entries in history (default: 100) */
    maxEntries?: number;
    /** Time in ms to group rapid changes (default: 500) */
    groupingInterval?: number;
    /** Whether to enable keyboard shortcuts (default: true) */
    enableKeyboardShortcuts?: boolean;
    /** Custom comparison function for detecting changes */
    isEqual?: (a: T, b: T) => boolean;
    /** Callback when undo is triggered */
    onUndo?: (state: T) => void;
    /** Callback when redo is triggered */
    onRedo?: (state: T) => void;
    /** Ref to the container element for keyboard events */
    containerRef?: React.RefObject<HTMLElement>;
}
/**
 * Return type of the useHistory hook
 */
interface UseHistoryReturn<T> {
    /** Current state */
    state: T;
    /** Whether undo is available */
    canUndo: boolean;
    /** Whether redo is available */
    canRedo: boolean;
    /** Number of entries in undo stack */
    undoCount: number;
    /** Number of entries in redo stack */
    redoCount: number;
    /** Push a new state to history */
    push: (newState: T, description?: string) => void;
    /** Undo to previous state */
    undo: () => T | undefined;
    /** Redo to next state */
    redo: () => T | undefined;
    /** Clear all history */
    clear: () => void;
    /** Reset to initial state and clear history */
    reset: (newInitialState?: T) => void;
    /** Get all undo entries (for debugging/display) */
    getUndoStack: () => HistoryEntry<T>[];
    /** Get all redo entries (for debugging/display) */
    getRedoStack: () => HistoryEntry<T>[];
    /** Transform all stored states (current + undo/redo stacks) */
    transformAll: (fn: (state: T) => T) => void;
}
/**
 * Custom hook for managing undo/redo history
 */
declare function useHistory<T>(initialState: T, options?: UseHistoryOptions<T>): UseHistoryReturn<T>;
/**
 * Simplified hook that just tracks state changes automatically
 */
declare function useAutoHistory<T>(value: T, options?: UseHistoryOptions<T>): Omit<UseHistoryReturn<T>, 'push'>;
/**
 * Hook for document history with specialized comparison
 */
declare function useDocumentHistory<T extends {
    package?: {
        ['document']?: unknown;
        headers?: unknown;
        footers?: unknown;
    } | null;
} | null>(docxDocument: T, options?: Omit<UseHistoryOptions<T>, 'isEqual'>): UseHistoryReturn<T>;
/**
 * Create a history manager for non-React usage
 */
declare class HistoryManager<T> {
    private undoStack;
    private redoStack;
    private currentState;
    private maxEntries;
    private groupingInterval;
    private lastPushTime;
    private isEqual;
    constructor(initialState: T, options?: {
        maxEntries?: number;
        groupingInterval?: number;
        isEqual?: (a: T, b: T) => boolean;
    });
    get state(): T;
    get canUndo(): boolean;
    get canRedo(): boolean;
    push(newState: T, description?: string): void;
    undo(): T | undefined;
    redo(): T | undefined;
    clear(): void;
    reset(newInitialState?: T): void;
}

/**
 * useTableSelection Hook
 *
 * Thin React wrapper around the framework-agnostic TableSelectionManager.
 * Provides table selection tracking and table operation dispatch.
 */

interface TableSelectionState {
    context: TableContext | null;
    table: Table | null;
    tableIndex: number | null;
    rowIndex: number | null;
    columnIndex: number | null;
}
interface UseTableSelectionReturn {
    state: TableSelectionState;
    handleCellClick: (tableIndex: number, rowIndex: number, columnIndex: number) => void;
    handleAction: (action: TableAction) => void;
    getSplitCellConfig: () => TableSplitConfig | null;
    applySplitCell: (rows: number, cols: number) => void;
    clearSelection: () => void;
    isCellSelected: (tableIndex: number, rowIndex: number, columnIndex: number) => boolean;
    tableContext: TableContext | null;
}
interface UseTableSelectionOptions {
    ['document']: Document | null;
    onChange?: (docxDocument: Document) => void;
    onSelectionChange?: (context: TableContext | null) => void;
}
declare function useTableSelection({ ['document']: doc, onChange, onSelectionChange, }: UseTableSelectionOptions): UseTableSelectionReturn;

/**
 * Selection Highlight Hook
 *
 * A React hook that manages visual selection highlighting across multiple runs.
 * Uses a combination of CSS ::selection pseudo-element styling and optional
 * overlay rectangles for complex scenarios.
 *
 * Features:
 * - Consistent selection highlighting across all text runs
 * - Support for text with different backgrounds (highlighted, dark bg)
 * - Optional overlay rectangles for custom highlight effects
 * - Debounced updates for performance
 */

/**
 * Options for the useSelectionHighlight hook
 */
interface UseSelectionHighlightOptions {
    /** Reference to the container element */
    containerRef: React__default.RefObject<HTMLElement>;
    /** Whether to enable selection highlighting */
    enabled?: boolean;
    /** Custom highlight configuration */
    config?: SelectionHighlightConfig;
    /** Whether to use overlay rectangles (default: false, uses CSS) */
    useOverlay?: boolean;
    /** Debounce delay for rect updates in ms (default: 16) */
    debounceMs?: number;
    /** Callback when selection changes */
    onSelectionChange?: (hasSelection: boolean, text: string) => void;
}
/**
 * Return value from the useSelectionHighlight hook
 */
interface UseSelectionHighlightReturn {
    /** Whether there is an active selection */
    hasSelection: boolean;
    /** The selected text */
    selectedText: string;
    /** Highlight rectangles (only populated if useOverlay is true) */
    highlightRects: HighlightRect[];
    /** Whether selection is within the container */
    isSelectionInContainer: boolean;
    /** Refresh the highlight state */
    refresh: () => void;
    /** Get styles for a highlight rect overlay */
    getOverlayStyle: (rect: HighlightRect) => CSSProperties;
}
/**
 * Hook to manage selection highlighting in the editor
 */
declare function useSelectionHighlight(options: UseSelectionHighlightOptions): UseSelectionHighlightReturn;
/**
 * Props for selection overlay component
 */
interface SelectionOverlayProps {
    /** Highlight rectangles to render */
    rects: HighlightRect[];
    /** Style configuration */
    config?: SelectionHighlightConfig;
    /** Additional class name */
    className?: string;
}
/**
 * Generate selection overlay elements (for use in JSX)
 *
 * Usage:
 * ```tsx
 * const { highlightRects } = useSelectionHighlight({ ... });
 * return (
 *   <div style={{ position: 'relative' }}>
 *     {generateOverlayElements(highlightRects)}
 *     <div>... content ...</div>
 *   </div>
 * );
 * ```
 */
declare function generateOverlayElements(rects: HighlightRect[], config?: SelectionHighlightConfig): React__default.ReactNode[];

/**
 * useClipboard Hook
 *
 * Thin React wrapper around the framework-agnostic ClipboardManager.
 * Handles clipboard operations with formatting preservation.
 */

interface UseClipboardOptions {
    onCopy?: (selection: ClipboardSelection) => void;
    onCut?: (selection: ClipboardSelection) => void;
    onPaste?: (content: ParsedClipboardContent, asPlainText: boolean) => void;
    cleanWordFormatting?: boolean;
    editable?: boolean;
    onError?: (error: Error) => void;
    /** Document theme — used to resolve themed colors in the HTML clipboard payload. */
    theme?: Theme | null;
}
interface UseClipboardReturn {
    copy: (selection: ClipboardSelection) => Promise<boolean>;
    cut: (selection: ClipboardSelection) => Promise<boolean>;
    paste: (asPlainText?: boolean) => Promise<ParsedClipboardContent | null>;
    handleCopy: (event: ClipboardEvent) => void;
    handleCut: (event: ClipboardEvent) => void;
    handlePaste: (event: ClipboardEvent) => void;
    handleKeyDown: (event: KeyboardEvent) => void;
    isProcessing: boolean;
    lastPastedContent: ParsedClipboardContent | null;
}
declare function useClipboard(options?: UseClipboardOptions): UseClipboardReturn;

/**
 * useAutoSave Hook
 *
 * Thin React wrapper around the framework-agnostic AutoSaveManager.
 * Bridges AutoSaveManager's subscribe/getSnapshot pattern with React state.
 */

/** Options for useAutoSave hook */
interface UseAutoSaveOptions {
    /** Storage key for localStorage (default: 'docx-editor-autosave') */
    storageKey?: string;
    /** Save interval in milliseconds (default: 30000 - 30 seconds) */
    interval?: number;
    /** Whether auto-save is enabled (default: true) */
    enabled?: boolean;
    /** Maximum age of auto-save in milliseconds before it's considered stale (default: 24 hours) */
    maxAge?: number;
    /** Callback when save succeeds */
    onSave?: (timestamp: Date) => void;
    /** Callback when save fails */
    onError?: (error: Error) => void;
    /** Callback when recovery data is found */
    onRecoveryAvailable?: (savedDocument: SavedDocumentData) => void;
    /** Whether to save immediately when document changes (debounced) */
    saveOnChange?: boolean;
    /** Debounce delay for saveOnChange in milliseconds (default: 2000) */
    debounceDelay?: number;
}
/** Return value of useAutoSave hook */
interface UseAutoSaveReturn {
    status: AutoSaveStatus;
    lastSaveTime: Date | null;
    save: () => Promise<boolean>;
    clearAutoSave: () => void;
    hasRecoveryData: boolean;
    getRecoveryData: () => SavedDocumentData | null;
    acceptRecovery: () => Document | null;
    dismissRecovery: () => void;
    isEnabled: boolean;
    enable: () => void;
    disable: () => void;
}
declare function useAutoSave(docxDocument: Document | null | undefined, options?: UseAutoSaveOptions): UseAutoSaveReturn;

/**
 * Drag Auto-Scroll Hook
 *
 * When the user is drag-selecting text and moves the mouse near the
 * top or bottom edge of the scroll container, this hook auto-scrolls
 * the container and continues extending the selection.
 */
interface DragAutoScrollOptions {
    /** Ref to the pages container (used to find the scroll parent). */
    pagesContainerRef: React.RefObject<HTMLDivElement | null>;
    /** Called during auto-scroll to extend the selection at the current mouse position. */
    onScrollExtendSelection: (clientX: number, clientY: number) => void;
}
declare function useDragAutoScroll({ pagesContainerRef, onScrollExtendSelection, }: DragAutoScrollOptions): {
    updateMousePosition: (clientX: number, clientY: number) => void;
    stopAutoScroll: () => void;
};

/**
 * Hook for toolbar dropdowns that need position:fixed to escape overflow:auto/hidden ancestors.
 *
 * Returns refs and styles for a dropdown that positions itself below its trigger
 * using fixed coordinates (like MenuDropdown), so it isn't clipped by the toolbar's
 * overflow-x-auto container.
 */

interface UseFixedDropdownOptions {
    isOpen: boolean;
    onClose: () => void;
    /** 'left' aligns dropdown left edge to trigger, 'right' aligns right edge */
    align?: 'left' | 'right';
}
interface UseFixedDropdownReturn {
    containerRef: RefObject<HTMLDivElement | null>;
    dropdownRef: RefObject<HTMLDivElement | null>;
    dropdownStyle: CSSProperties;
    handleMouseDown: (e: React.MouseEvent) => void;
}
declare function useFixedDropdown({ isOpen, onClose, align, }: UseFixedDropdownOptions): UseFixedDropdownReturn;

/**
 * Width/height inputs with an optional aspect-ratio lock. `width`/`height`
 * are `number | ''` so a cleared field shows empty instead of 0.
 */
interface UseAspectLockedSizeReturn {
    width: number | '';
    height: number | '';
    lockAspect: boolean;
    setLockAspect: (locked: boolean) => void;
    /** Number-input onChange handlers. Empty string clears, otherwise clamps to >= 1. */
    handleWidthChange: (raw: string) => void;
    handleHeightChange: (raw: string) => void;
    /** Seed both fields and re-lock. Null/undefined values clear the input. */
    seed: (w: number | null | undefined, h: number | null | undefined) => void;
}
declare function useAspectLockedSize(): UseAspectLockedSizeReturn;

interface VisualLineNavigationOptions {
    pagesContainerRef: React.RefObject<HTMLDivElement | null>;
}
declare function useVisualLineNavigation({ pagesContainerRef }: VisualLineNavigationOptions): {
    stickyXRef: React$1.RefObject<number | null>;
    lastVisualLineIndexRef: React$1.RefObject<number>;
    getCaretClientX: (pmPos: number) => number | null;
    findLineElementAtPosition: (pmPos: number) => HTMLElement | null;
    findPositionOnLineAtClientX: (lineEl: HTMLElement, clientX: number) => number | null;
    handlePMKeyDown: (view: EditorView, event: KeyboardEvent) => boolean;
};

/**
 * useWheelZoom Hook
 *
 * Enables Ctrl+scroll (or Cmd+scroll on Mac) to zoom in/out.
 * Features:
 * - Configurable zoom range and step
 * - Smooth zoom transitions
 * - Pinch-to-zoom support on trackpads
 * - Zoom reset (Ctrl+0)
 * - Zoom in/out shortcuts (Ctrl++, Ctrl+-)
 */
/**
 * Options for useWheelZoom hook
 */
interface UseWheelZoomOptions {
    /** Initial zoom level (default: 1.0) */
    initialZoom?: number;
    /** Minimum zoom level (default: 0.25) */
    minZoom?: number;
    /** Maximum zoom level (default: 4.0) */
    maxZoom?: number;
    /** Zoom step for each scroll event (default: 0.1) */
    zoomStep?: number;
    /** Whether zoom is enabled (default: true) */
    enabled?: boolean;
    /** Container element ref to attach wheel listener */
    containerRef?: React.RefObject<HTMLElement>;
    /** Callback when zoom changes */
    onZoomChange?: (zoom: number) => void;
    /** Whether to enable keyboard shortcuts (Ctrl++, Ctrl+-, Ctrl+0) */
    enableKeyboardShortcuts?: boolean;
    /** Whether to prevent default browser zoom behavior */
    preventDefault?: boolean;
}
/**
 * Return value of useWheelZoom hook
 */
interface UseWheelZoomReturn {
    /** Current zoom level */
    zoom: number;
    /** Set zoom level directly */
    setZoom: (zoom: number) => void;
    /** Zoom in by step */
    zoomIn: () => void;
    /** Zoom out by step */
    zoomOut: () => void;
    /** Reset zoom to initial level */
    resetZoom: () => void;
    /** Reset zoom to 100% */
    zoomTo100: () => void;
    /** Zoom to fit width */
    zoomToFit: (containerWidth: number, contentWidth: number) => void;
    /** Whether currently at minimum zoom */
    isMinZoom: boolean;
    /** Whether currently at maximum zoom */
    isMaxZoom: boolean;
    /** Zoom percentage (e.g., 100 for zoom level 1.0) */
    zoomPercent: number;
    /** Wheel event handler (for manual attachment) */
    handleWheel: (event: WheelEvent) => void;
    /** Keyboard event handler (for manual attachment) */
    handleKeyDown: (event: KeyboardEvent) => void;
}
/**
 * Preset zoom levels for snapping
 */
declare const ZOOM_PRESETS: number[];
/**
 * React hook for Ctrl+scroll zoom functionality
 */
declare function useWheelZoom(options?: UseWheelZoomOptions): UseWheelZoomReturn;
/**
 * Get zoom presets
 */
declare function getZoomPresets(): number[];
/**
 * Find nearest zoom preset
 */
declare function findNearestZoomPreset(zoom: number): number;
/**
 * Get next zoom preset (for zoom in)
 */
declare function getNextZoomPreset(zoom: number): number;
/**
 * Get previous zoom preset (for zoom out)
 */
declare function getPreviousZoomPreset(zoom: number): number;
/**
 * Format zoom level for display
 */
declare function formatZoom(zoom: number): string;
/**
 * Parse zoom from percentage string
 */
declare function parseZoom(zoomString: string): number | null;
/**
 * Check if zoom level is at a preset
 */
declare function isZoomPreset(zoom: number): boolean;
/**
 * Clamp zoom to valid range
 */
declare function clampZoom(zoom: number, minZoom?: number, maxZoom?: number): number;

/**
 * Returns tracked changes (and the comment→revision overlap map for threading)
 * derived from the latest PM state. Memoized on state identity, so derivation
 * only re-runs when PM state changes (which happens on every doc-changing
 * transaction, including remote ones via ySyncPlugin).
 *
 * No debounce: a single O(N) doc walk, cheap enough to run per transaction.
 * If you see jank on huge documents, wrap the setter that drives the state
 * argument in `requestAnimationFrame` rather than reintroducing a delay here —
 * a delay makes the sidebar feel laggy.
 */
declare function useTrackedChanges(state: EditorState | null): TrackedChangesResult;

export { type DragAutoScrollOptions, type HistoryEntry, HistoryManager, type SelectionOverlayProps, type TableSelectionState, type UseAspectLockedSizeReturn, type UseAutoSaveOptions, type UseAutoSaveReturn, type UseClipboardOptions, type UseClipboardReturn, type UseFixedDropdownOptions, type UseFixedDropdownReturn, type UseHistoryOptions, type UseHistoryReturn, type UseSelectionHighlightOptions, type UseSelectionHighlightReturn, type UseTableSelectionOptions, type UseTableSelectionReturn, type UseWheelZoomOptions, type UseWheelZoomReturn, type VisualLineNavigationOptions, ZOOM_PRESETS, clampZoom, findNearestZoomPreset, formatZoom, generateOverlayElements, getNextZoomPreset, getPreviousZoomPreset, getZoomPresets, isZoomPreset, parseZoom, useAspectLockedSize, useAutoHistory, useAutoSave, useClipboard, useDocumentHistory, useDragAutoScroll, useFixedDropdown, useHistory, useSelectionHighlight, useTableSelection, useTrackedChanges, useVisualLineNavigation, useWheelZoom };
