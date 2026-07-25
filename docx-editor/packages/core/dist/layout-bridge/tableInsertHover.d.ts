/**
 * Table-insert "+" hover hit-test.
 *
 * Pure DOM logic for the floating row/column insert button that shows
 * when the mouse is near the left or top edge of a layout table. Lives
 * in core so React + Vue + any future adapter can share the hit-test
 * and just wire up their own UI rendering of the button.
 *
 * The function is gated by `hfEditMode`: tables in inactive HF/body
 * regions don't surface the affordance. A header table only matches
 * when the user is editing the header; a body table only matches when
 * not in any HF edit mode.
 * @packageDocumentation
 * @public
 */
declare const TABLE_INSERT_EDGE_PROXIMITY = 30;
declare const TABLE_INSERT_HIDE_DELAY_MS = 200;
type TableInsertHoverHit = {
    type: 'row' | 'column';
    /** Client-coordinate anchor for the button. Caller converts to its UI frame. */
    clientX: number;
    clientY: number;
    /** PM position of the cell the button targets. */
    cellPmPos: number;
};
type TableInsertHoverInput = {
    mouseX: number;
    mouseY: number;
    pagesContainer: HTMLElement;
    /** Element under the cursor at the time of the event (e.target). */
    target: HTMLElement;
    hfEditMode: 'header' | 'footer' | null;
    /** Override edge proximity in pixels (defaults to TABLE_INSERT_EDGE_PROXIMITY). */
    edgeProximity?: number;
};
/**
 * Detect whether a mousemove should surface a row/column insert "+" button.
 *
 * Returns the button anchor + target cell PM position, or `null` if the
 * mouse isn't near a row's left edge or a column's top edge — or if the
 * relevant table belongs to an inactive HF/body region.
 */
declare function detectTableInsertHover(input: TableInsertHoverInput): TableInsertHoverHit | null;

export { TABLE_INSERT_EDGE_PROXIMITY, TABLE_INSERT_HIDE_DELAY_MS, type TableInsertHoverHit, type TableInsertHoverInput, detectTableInsertHover };
