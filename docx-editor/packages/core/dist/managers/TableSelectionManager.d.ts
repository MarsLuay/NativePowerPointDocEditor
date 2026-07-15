/**
 * TableSelectionManager
 *
 * Framework-agnostic class for managing table cell selection state.
 * Extracted from the React `useTableSelection` hook.
 *
 * Handles:
 * - Cell selection via data-attribute queries on the DOM
 * - Table document operations (add/delete rows/columns, merge/split)
 *
 * @remarks
 * Tagged `@internal` post-1.0 cut. Adapters bind to this class today,
 * but consumers should reach for `useTableSelection()` (React/Vue) instead
 * of this raw subpath. The subpath stays in `package.json` `exports` for
 * back-compat; expect it to move behind a public surface in a future major.
 *
 * @packageDocumentation
 * @internal
 */
import { S as Subscribable } from '../Subscribable-DOz6Ohoo.js';
import { TableSelectionSnapshot, CellCoordinates } from './types.js';
import { Document } from '../types/document.js';
import { T as Table } from '../content-B8ScSBzC.js';
import 'prosemirror-view';
import '../colors-C3vA7HUU.js';
import '../formatting-JhqWT_XM.js';
import '../lists-Bn29SzeS.js';
import '../docx/wrapTypes.js';
import '../watermark-D90356ZM.js';
import '../styles-BGGjYqnf.js';

/**
 * TableSelectionManager
 *
 * Framework-agnostic class for managing table cell selection state.
 * Extracted from the React `useTableSelection` hook.
 *
 * Handles:
 * - Cell selection via data-attribute queries on the DOM
 * - Table document operations (add/delete rows/columns, merge/split)
 *
 * @remarks
 * Tagged `@internal` post-1.0 cut. Adapters bind to this class today,
 * but consumers should reach for `useTableSelection()` (React/Vue) instead
 * of this raw subpath. The subpath stays in `package.json` `exports` for
 * back-compat; expect it to move behind a public surface in a future major.
 *
 * @packageDocumentation
 * @internal
 */

/**
 * Data attributes for table elements in the rendered DOM
 * @internal
 */
declare const TABLE_DATA_ATTRIBUTES: {
    readonly TABLE_INDEX: "data-table-index";
    readonly ROW_INDEX: "data-row";
    readonly COLUMN_INDEX: "data-col";
    readonly TABLE_CELL: "data-table-cell";
};
/**
 * Find table cell coordinates from a click target by walking up the DOM
 * and reading data attributes.
 * @internal
 */
declare function findTableFromClick(target: EventTarget | null, container?: HTMLElement | null): CellCoordinates | null;
/**
 * Get a table from the document by index.
 * @internal
 */
declare function getTableFromDocument(doc: Document, tableIndex: number): Table | null;
/**
 * Update a table in the document immutably.
 * @internal
 */
declare function updateTableInDocument(doc: Document, tableIndex: number, newTable: Table): Document;
/**
 * Delete a table from the document immutably.
 * @internal
 */
declare function deleteTableFromDocument(doc: Document, tableIndex: number): Document;
/** @internal */
declare class TableSelectionManager extends Subscribable<TableSelectionSnapshot> {
    constructor();
    /** Select a specific cell. */
    selectCell(coords: CellCoordinates): void;
    /** Clear the current selection. */
    clearSelection(): void;
    /** Check if a specific cell is selected. */
    isCellSelected(tableIndex: number, rowIndex: number, columnIndex: number): boolean;
    /** Get the currently selected cell coordinates, or null. */
    getSelectedCell(): CellCoordinates | null;
}

export { TABLE_DATA_ATTRIBUTES, TableSelectionManager, deleteTableFromDocument, findTableFromClick, getTableFromDocument, updateTableInDocument };
