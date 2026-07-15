import { EditorView } from 'prosemirror-view';

/**
 * Drag-to-select table cells, shared by the React and Vue pages-pointer
 * handlers. A pointer drag that starts inside a table cell and crosses into
 * another cell is promoted from a text selection to a prosemirror-tables
 * `CellSelection` (so multi-cell ops — delete row/column, fill, merge — become
 * reachable by dragging). Pure `(view, pmPos, clientX) → boolean` logic; the
 * adapters own the DOM event wiring and the active-view resolution (body vs
 * header/footer).
 */

/**
 * @deprecated No longer used. The tracker ignores pointer overflow so same-cell
 * drags stay text selections. Kept only for backward-compatible imports; will be
 * removed in a future major.
 */
declare const CELL_SELECT_OVERFLOW_PX = 5;
/**
 * Walk up from a PM position to the enclosing `tableCell`/`tableHeader` and
 * return its `before(d)` position (what `CellSelection` resolves against), or
 * null when the position isn't inside a table cell.
 */
declare function findCellPosFromPmPos(view: EditorView, pmPos: number): number | null;
/** Dispatch a `CellSelection` spanning the two cell positions. */
declare function applyCellSelection(view: EditorView, anchorCellPos: number, headCellPos: number): boolean;
/**
 * Per-gesture state machine for cell-drag selection. Create one per pointer
 * handler; drive it from mousedown/mousemove/mouseup.
 */
interface CellDragTracker {
    /** mousedown: record the cell under the press (or null when outside a table). */
    begin(cellPos: number | null): void;
    /**
     * mousemove: returns true when it applied a `CellSelection` and the caller
     * should NOT also run its text-drag selection for this move.
     */
    update(view: EditorView, pmPos: number, clientX: number): boolean;
    /** mouseup / gesture end: clear the transient drag flags. */
    end(): void;
    /** True while a cell drag is actively producing CellSelections. */
    readonly isCellDragging: boolean;
}
declare function createCellDragTracker(): CellDragTracker;

export { CELL_SELECT_OVERFLOW_PX, type CellDragTracker, applyCellSelection, createCellDragTracker, findCellPosFromPmPos };
