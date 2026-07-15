import { EditorView } from 'prosemirror-view';

/**
 * Table-resize commit + width-read helpers shared by the React and Vue
 * adapters. Pure `(view, …) → result` over the PM doc — the gesture state
 * machines (which DOM handle is grabbed, pixel deltas, listener wiring) stay
 * in each adapter.
 *
 * Widths and heights are stored in twips (1/20 of a point) on the PM doc,
 * matching OOXML's `w:tblGrid` and `w:tcW` units.
 */

/** 1px ≈ 15 twips at 96dpi (20 twips/pt × 72pt/in ÷ 96px/in). */
declare const TWIPS_PER_PIXEL = 15;
/** Minimum column width (~0.2"). */
declare const MIN_CELL_WIDTH_TWIPS = 300;
/** Minimum row height (~0.14"). */
declare const MIN_ROW_HEIGHT_TWIPS = 200;
/** Read the [left, right] column widths at `colIndex` in the table starting at `pmStart`. */
declare function readColumnWidths(view: EditorView, pmStart: number, colIndex: number): {
    left: number;
    right: number;
} | null;
/**
 * Read the row height (in twips) for `rowIndex` in the table starting at
 * `pmStart`. Returns null if the row has no explicit height — the caller
 * can fall back to measuring the rendered DOM cell.
 */
declare function readRowHeight(view: EditorView, pmStart: number, rowIndex: number): number | null;
/** Read the last-column width (the one being resized from the table's right edge). */
declare function readColumnWidthAt(view: EditorView, pmStart: number, colIndex: number): number | null;
/**
 * Bake a column-resize into the PM doc: update the table's `columnWidths`
 * attr and every cell at `colIdx` / `colIdx + 1` in every row.
 */
declare function commitColumnResize(view: EditorView, opts: {
    pmStart: number;
    colIdx: number;
    newLeft: number;
    newRight: number;
}): void;
/** Bake a row-resize into the PM doc: update the target row's `height` + `heightRule`. */
declare function commitRowResize(view: EditorView, opts: {
    pmStart: number;
    rowIdx: number;
    newHeight: number;
}): void;
/**
 * Bake a right-edge resize into the PM doc: grow only the last column.
 * Updates `columnWidths[colIdx]` and the cell at `colIdx` in every row.
 */
declare function commitRightEdgeResize(view: EditorView, opts: {
    pmStart: number;
    colIdx: number;
    newWidth: number;
}): void;

export { MIN_CELL_WIDTH_TWIPS, MIN_ROW_HEIGHT_TWIPS, TWIPS_PER_PIXEL, commitColumnResize, commitRightEdgeResize, commitRowResize, readColumnWidthAt, readColumnWidths, readRowHeight };
