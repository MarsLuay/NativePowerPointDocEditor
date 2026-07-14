import React__default, { CSSProperties, ReactNode } from 'react';
import { Table } from '@eigenpal/docx-editor-core/types/document';
import { FindMatch, FindOptions } from '@eigenpal/docx-editor-core/utils/findReplace';

/**
 * TableToolbar Component
 *
 * Provides controls for editing tables:
 * - Add row above/below
 * - Add column left/right
 * - Delete row/column
 * - Merge cells
 * - Split cell
 *
 * Shows when cursor is in a table.
 */

/**
 * Table editing action types
 */
type TableAction = 'addRowAbove' | 'addRowBelow' | 'addColumnLeft' | 'addColumnRight' | 'deleteRow' | 'deleteColumn' | 'mergeCells' | 'splitCell' | 'deleteTable' | 'selectTable' | 'selectRow' | 'selectColumn' | 'borderAll' | 'borderOutside' | 'borderInside' | 'borderNone' | 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight' | {
    type: 'cellFillColor';
    color: string | null;
} | {
    type: 'borderColor';
    color: string;
} | {
    type: 'borderWidth';
    size: number;
} | {
    type: 'cellBorder';
    side: 'top' | 'bottom' | 'left' | 'right' | 'all';
    style: string;
    size: number;
    color: string;
} | {
    type: 'cellVerticalAlign';
    align: 'top' | 'center' | 'bottom';
} | {
    type: 'cellMargins';
    margins: {
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
} | {
    type: 'cellTextDirection';
    direction: string | null;
} | {
    type: 'toggleNoWrap';
} | {
    type: 'rowHeight';
    height: number | null;
    rule?: 'auto' | 'atLeast' | 'exact';
} | {
    type: 'toggleHeaderRow';
} | {
    type: 'distributeColumns';
} | {
    type: 'autoFitContents';
} | {
    type: 'tableProperties';
    props: {
        width?: number | null;
        widthType?: string | null;
        justification?: 'left' | 'center' | 'right' | null;
    };
} | {
    type: 'openTableProperties';
} | {
    type: 'applyTableStyle';
    styleId: string;
};
/**
 * Selection within a table
 */
interface TableSelection {
    /** Index of the table in the document */
    tableIndex: number;
    /** Row index (0-indexed) */
    rowIndex: number;
    /** Column index (0-indexed) */
    columnIndex: number;
    /** Selected cell range for multi-cell selection */
    selectedCells?: {
        startRow: number;
        startCol: number;
        endRow: number;
        endCol: number;
    };
}
/**
 * Context for table operations
 */
interface TableContext {
    /** The table being edited */
    table: Table;
    /** Current selection within the table */
    selection: TableSelection;
    /** Whether multiple cells are selected (for merge) */
    hasMultiCellSelection: boolean;
    /** Whether current cell can be split */
    canSplitCell: boolean;
    /** Total number of rows */
    rowCount: number;
    /** Total number of columns */
    columnCount: number;
}
interface TableSplitConfig {
    minRows: number;
    minCols: number;
    initialRows: number;
    initialCols: number;
}
/**
 * Props for TableToolbar component
 */
interface TableToolbarProps {
    /** Current table context (null if cursor not in table) */
    context: TableContext | null;
    /** Callback when a table action is triggered */
    onAction?: (action: TableAction, context: TableContext) => void;
    /** Whether the toolbar is disabled */
    disabled?: boolean;
    /** Additional CSS class name */
    className?: string;
    /** Additional inline styles */
    style?: CSSProperties;
    /** Show labels next to icons */
    showLabels?: boolean;
    /** Compact mode */
    compact?: boolean;
    /** Position of the toolbar */
    position?: 'top' | 'floating';
    /** Custom render for additional buttons */
    children?: ReactNode;
}
/**
 * TableToolbar - Shows table manipulation controls when cursor is in a table
 */
declare function TableToolbar({ context, onAction, disabled, className, style, showLabels, compact, position, children, }: TableToolbarProps): React__default.ReactElement | null;

/**
 * useFindReplace Hook
 *
 * React hook for managing find/replace dialog state.
 * Extracted from FindReplaceDialog.tsx.
 */

/**
 * Options for the useFindReplace hook
 */
interface FindReplaceOptions {
    /** Whether to show replace functionality initially */
    initialReplaceMode?: boolean;
    /** Callback when matches change */
    onMatchesChange?: (matches: FindMatch[]) => void;
    /** Callback when current match changes */
    onCurrentMatchChange?: (match: FindMatch | null, index: number) => void;
}
/**
 * State for the find/replace hook
 */
interface FindReplaceState {
    /** Whether the dialog is open */
    isOpen: boolean;
    /** Current search text */
    searchText: string;
    /** Current replace text */
    replaceText: string;
    /** Find options */
    options: FindOptions;
    /** All matches found */
    matches: FindMatch[];
    /** Current match index */
    currentIndex: number;
    /** Whether in replace mode */
    replaceMode: boolean;
}
/**
 * Return type for the useFindReplace hook
 */
interface UseFindReplaceReturn {
    /** Current state */
    state: FindReplaceState;
    /** Open the find dialog */
    openFind: (selectedText?: string) => void;
    /** Open the replace dialog */
    openReplace: (selectedText?: string) => void;
    /** Close the dialog */
    close: () => void;
    /** Toggle dialog visibility */
    toggle: () => void;
    /** Update search text */
    setSearchText: (text: string) => void;
    /** Update replace text */
    setReplaceText: (text: string) => void;
    /** Update find options */
    setOptions: (options: Partial<FindOptions>) => void;
    /** Set search results */
    setMatches: (matches: FindMatch[], currentIndex?: number) => void;
    /** Go to next match */
    goToNextMatch: () => number;
    /** Go to previous match */
    goToPreviousMatch: () => number;
    /** Go to a specific match by index */
    goToMatch: (index: number) => void;
    /** Get current match */
    getCurrentMatch: () => FindMatch | null;
    /** Check if has matches */
    hasMatches: () => boolean;
}
/**
 * Hook for managing find/replace dialog state
 */
declare function useFindReplace(hookOptions?: FindReplaceOptions): UseFindReplaceReturn;

export { type FindReplaceOptions as F, type TableSelection as T, type UseFindReplaceReturn as U, type TableContext as a, type TableSplitConfig as b, type TableAction as c, type FindReplaceState as d, TableToolbar as e, type TableToolbarProps as f, useFindReplace as u };
