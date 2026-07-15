import { Node } from 'prosemirror-model';
import { EditorState, Command } from 'prosemirror-state';
import { C as ColorValue } from './colors-C3vA7HUU.js';

/**
 * Table selection context + navigation helpers.
 *
 * `getTableContext` walks the selection up from `$from` and reports which
 * table / row / cell the cursor is in, plus the table's row/column counts,
 * whether a multi-cell selection is active, and the current cell's border
 * + fill colors (so the toolbar's color pickers can show the live values).
 *
 * `goToNextCell` / `goToPrevCell` are tab-stop-style cell navigation commands
 * registered by the plugin extension.
 */

interface TableContextInfo {
    isInTable: boolean;
    table?: Node;
    tablePos?: number;
    rowIndex?: number;
    columnIndex?: number;
    rowCount?: number;
    columnCount?: number;
    hasMultiCellSelection?: boolean;
    canSplitCell?: boolean;
    /** Current cell's dominant border color, if any */
    cellBorderColor?: ColorValue;
    /** Current cell's background/fill color (RGB hex without #), if any */
    cellBackgroundColor?: string;
}
declare function getTableContext(state: EditorState): TableContextInfo;
declare function isInTableCell(state: EditorState): boolean;
declare function goToNextCell(): Command;
declare function goToPrevCell(): Command;

/**
 * Cell-border commands. Each command applies a preset / individual side /
 * color / width to the cells targeted by the current selection (single
 * cursor cell or active `CellSelection`).
 *
 * All four commands use the shared `buildTableGrid` lookup to find each
 * cell's neighbours in the grid, then sync the matching edge on the
 * adjacent cell — Google-Docs style edge-symmetric border editing.
 *
 * Schema-free: only attribute updates via `tr.setNodeMarkup`.
 */

type BorderPreset = 'all' | 'outside' | 'inside' | 'none';
type BorderSpec = {
    style: string;
    size: number;
    color: {
        rgb: string;
    };
};

export { type BorderPreset as B, type TableContextInfo as T, type BorderSpec as a, goToNextCell as b, goToPrevCell as c, getTableContext as g, isInTableCell as i };
