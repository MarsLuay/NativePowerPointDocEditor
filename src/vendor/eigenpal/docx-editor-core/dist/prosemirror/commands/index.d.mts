/**
 * ProseMirror Commands
 *
 * Commands for formatting text and paragraphs.
 * @packageDocumentation
 * @public
 */
export { clearFontFamily, clearFontSize, clearFormatting, clearHighlight, clearTextColor, createRemoveMarkCommand, createSetMarkCommand, findHyperlinkRangeAt, getHyperlinkAttrs, getMarkAttr, getSelectedText, insertHyperlink, isHyperlinkActive, isMarkActive, removeHyperlink, setFontFamily, setFontSize, setHighlight, setHyperlink, setTextColor, setUnderlineStyle, toggleBold, toggleItalic, toggleStrike, toggleSubscript, toggleSuperscript, toggleUnderline } from './formatting.mjs';
export { ResolvedStyleAttrs, addTabStop, alignCenter, alignJustify, alignLeft, alignRight, applyStyle, clearStyle, decreaseIndent, decreaseListLevel, doubleSpacing, generateTOC, getListInfo, getParagraphAlignment, getParagraphBidi, getStyleId, increaseIndent, increaseListLevel, isInList, oneAndHalfSpacing, removeList, removeTabStop, setAlignment, setIndentFirstLine, setIndentLeft, setIndentRight, setLineSpacing, setLtr, setRtl, setSpaceAfter, setSpaceBefore, singleSpacing, toggleBulletList, toggleNumberedList } from './paragraph.mjs';
export { a as addColumnLeft, b as addColumnRight, c as addRowAbove, d as addRowBelow, e as applyTableStyle, f as autoFitContents, g as deleteColumn, h as deleteRow, i as deleteTable, j as distributeColumns, k as insertTable, m as mergeCells, r as removeTableBorders, s as selectColumn, l as selectRow, n as selectTable, o as setAllTableBorders, p as setCellBorder, q as setCellFillColor, t as setCellMargins, u as setCellTextDirection, v as setCellVerticalAlign, w as setInsideTableBorders, x as setOutsideTableBorders, y as setRowHeight, z as setTableBorderColor, A as setTableBorderWidth, B as setTableBorders, C as setTableProperties, D as splitCell, E as toggleHeaderRow, F as toggleNoWrap } from '../../table-_t_W8oQf.mjs';
export { insertPageBreak } from './pageBreak.mjs';
export { insertSectionBreakContinuous, insertSectionBreakNextPage } from './sectionBreak.mjs';
import { EditorState, Transaction, Command } from 'prosemirror-state';
import { Node } from 'prosemirror-model';
import { EditorView } from 'prosemirror-view';
import { I as ImageLayoutTarget, S as SetImageWrapTypeOptions } from '../../ImageExtension-BN327PNe.mjs';
export { A as AnchorWrapType } from '../../ImageExtension-BN327PNe.mjs';
import { W as Watermark } from '../../watermark-D90356ZM.mjs';
export { B as BorderPreset, T as TableContextInfo, g as getTableContext, i as isInTable } from '../../borders-BH_GoYU3.mjs';
import '../../formatting-DFtuRFQY.mjs';
import '../../colors-C3vA7HUU.mjs';
import '../../marks-DOMflBhB.mjs';
import '../../numberingParser-nSPZlzKr.mjs';
import '../../lists-CyGxd5Y2.mjs';
import '../../docx/wrapTypes.mjs';

/**
 * ProseMirror table cell splitting — dialog-backed split with explicit
 * row/column input.
 *
 * Uses the shared split algorithm from utils/tableSplitAlgorithm.ts for the
 * core layout computation, then maps the result back into ProseMirror
 * transactions.
 */

interface SplitCellDialogConfig {
    minRows: number;
    minCols: number;
    initialRows: number;
    initialCols: number;
    /** Captured row index of the target cell at dialog-open time */
    capturedCellRow: number;
    /** Captured column index of the target cell at dialog-open time */
    capturedCellCol: number;
}
declare function getSplitCellDialogConfig(state: EditorState): SplitCellDialogConfig | null;
declare function splitActiveTableCell(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, rows: number, cols: number,
/** Target cell row/col captured at dialog-open time. Falls back to current cursor. */
targetRow?: number, targetCol?: number): boolean;

/**
 * Image commands — thin re-exports from the extension system.
 *
 * Wrap-type transitions for floating images. Inline↔anchor conversions are
 * structural and live in a follow-up; this surface only covers anchor↔anchor.
 */

/**
 * Insert an image node at `pos`, wrapping with the `insertion` mark when
 * suggesting mode is active. Centralizes the tracked-image-insert flow
 * so React `useFileIO`, Vue `useImageActions`, and clipboard-paste
 * paths all share one source of truth — adding a fresh image in
 * suggesting mode always round-trips as `<w:ins>{run with drawing}</w:ins>`.
 *
 * Caller responsibility: produce the `image` node via `schema.nodes.image.create`.
 * This helper handles the dispatch + optional mark application.
 *
 * @public
 */
declare function insertImageNode(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, imageNode: Node, pos: number): boolean;
/**
 * Default max width (px) for an image inserted from a file picker — the content
 * area of a US Letter page at 96dpi (~6.375in). Images wider than this are
 * scaled down to fit the column, matching Word and keeping the painter's
 * `max-width: 100%` from shrinking the rendered height out from under the
 * reserved line height (which would leave a gap below the image).
 */
declare const INSERT_IMAGE_MAX_WIDTH_PX = 612;
/**
 * Read an image `File` (from a file picker or drop), fit it to the page width,
 * and insert it inline at the current selection. This is the single source of
 * truth for "insert an image from a file" — the React and Vue adapters both
 * call it, so insertion behaves identically: no intermediate dialog, the image
 * is sized to fit the column, and it round-trips as an inline drawing (with the
 * `insertion` mark applied in suggesting mode, via {@link insertImageNode}).
 *
 * The decode is async (FileReader → Image); `onError` reports a failed read or
 * decode, and `onInserted` runs after the node lands (e.g. to refocus).
 *
 * @public
 */
declare function insertImageFromFile(view: EditorView, file: File, opts?: {
    maxWidth?: number;
    onError?: (error: unknown) => void;
    onInserted?: () => void;
}): void;
/**
 * Change a floating image's wrap layout. `pos` is the PM document position of
 * the image node; `target` is either an OOXML wrap type (square / tight /
 * topAndBottom / behind / inFront / inline) or a directional convenience
 * choice (`squareLeft` / `squareRight`).
 *
 * `opts.initialPositionEmu` is used when promoting an inline image to an
 * anchor — the caller measures the image's current rendered offset relative
 * to the column origin in EMUs and passes it through, so the new float lands
 * exactly where the inline glyph used to sit (matches Word's behavior).
 */
declare function setImageWrapType(pos: number, target: ImageLayoutTarget, opts?: SetImageWrapTypeOptions): Command;

/**
 * Comment and Track Changes Commands
 *
 * PM commands for adding/removing comments and accepting/rejecting tracked changes.
 */

/**
 * Add a comment mark to the current selection.
 */
declare function addCommentMark(commentId: number): Command;
/**
 * Remove a comment mark by ID from the entire document.
 */
declare function removeCommentMark(commentId: number): Command;
/**
 * Accept a tracked change at the given range.
 * - Insertion: remove mark, keep text
 * - Deletion: remove mark AND text
 *
 * Use {@link acceptChangeById} when accepting a coalesced revision —
 * a single editing session can scatter sites across multiple paragraphs
 * and only the by-id resolver walks every site.
 *
 * @example
 * ```ts
 * import { acceptChange } from '@eigenpal/docx-editor-core/prosemirror/commands';
 * acceptChange(from, to)(view.state, view.dispatch);
 * ```
 */
declare function acceptChange(from: number, to: number): Command;
/**
 * Reject a tracked change at the given range.
 * - Insertion: remove mark AND text
 * - Deletion: remove mark, keep text
 *
 * Use {@link rejectChangeById} when rejecting a coalesced revision.
 */
declare function rejectChange(from: number, to: number): Command;
/**
 * Accept every tracked change in the document — inline marks plus
 * structural revisions (paragraph-mark, row, cell, property changes).
 *
 * Dispatches one transaction per distinct `revisionId` so each revision
 * remains individually undoable. The acceptance order follows document
 * order; later transactions read fresh state and skip ids whose sites
 * were already removed.
 *
 * @example
 * ```ts
 * import { acceptAllChanges } from '@eigenpal/docx-editor-core/prosemirror/commands';
 * acceptAllChanges()(view.state, view.dispatch);
 * ```
 */
declare function acceptAllChanges(): Command;
/**
 * Reject all tracked changes in the document — inverse of `acceptAllChanges`.
 */
declare function rejectAllChanges(): Command;
interface ChangeRange {
    from: number;
    to: number;
    type: 'insertion' | 'deletion';
}
/**
 * Find the next tracked-change range (inline insertion / deletion mark)
 * after `startPos`. Wraps to the document start when no later change is
 * found. Useful for "next change" / "previous change" toolbar buttons.
 *
 * Only walks inline marks — structural revisions (pPrIns / pPrDel / row
 * / cell) are not surfaced here. Use {@link extractTrackedChanges} for a
 * complete revision list.
 */
declare function findNextChange(state: EditorState, startPos: number): ChangeRange | null;
/**
 * Accept every site of a tracked revision in one PM transaction. Walks
 * the doc for all sites carrying `revisionId` — inline insertion/
 * deletion marks, paragraph-mark `pPrIns` / `pPrDel`, paragraph
 * property changes, table row / cell / table revisions — and applies
 * the accept semantics each requires.
 *
 * This is the right command for any coalesced revision (Enter chains,
 * replace pairs, multi-paragraph runs) because one editing session can
 * scatter sites across the doc; the range-based {@link acceptChange}
 * only clears sites within its `(from, to)`.
 *
 * Returns `false` (no-op) if the id is not present.
 *
 * @example
 * ```ts
 * import { acceptChangeById } from '@eigenpal/docx-editor-core/prosemirror/commands';
 * acceptChangeById(revisionId)(view.state, view.dispatch);
 * ```
 */
declare function acceptChangeById(revisionId: number): Command;
/**
 * Reject every site of a tracked revision in one PM transaction.
 * Inverse of {@link acceptChangeById} — inserts are removed, deletions
 * keep their text, paragraph-mark insertions join paragraphs back,
 * paragraph-mark deletions stay split, paragraph property changes
 * restore prior values, and tracked row insertions are removed.
 */
declare function rejectChangeById(revisionId: number): Command;
/**
 * Find the previous tracked-change range (inline insertion / deletion
 * mark) before `startPos`. Wraps to the document end when no earlier
 * change is found. Counterpart to {@link findNextChange}.
 */
declare function findPreviousChange(state: EditorState, startPos: number): ChangeRange | null;

/**
 * Watermark command.
 *
 * The document watermark lives as a `doc` node attribute (see DocExtension), so
 * setting it is a normal ProseMirror transaction — it rides undo/redo, the
 * toolbar undo/redo buttons, and Ctrl+Z like any other edit. The painter reads
 * the watermark from PM state; the conversion layer syncs it to/from
 * `HeaderFooter.watermark` for parse/serialize.
 */

/** Read the current watermark from a ProseMirror state's doc attrs. */
declare function getWatermarkFromState(state: EditorState): Watermark | null;
/**
 * Set (or clear, with `null`) the document watermark. Dispatches a
 * `setDocAttribute` transaction so the change is undoable.
 */
declare function setWatermark(watermark: Watermark | null): Command;

export { INSERT_IMAGE_MAX_WIDTH_PX, ImageLayoutTarget, SetImageWrapTypeOptions, type SplitCellDialogConfig, acceptAllChanges, acceptChange, acceptChangeById, addCommentMark, findNextChange, findPreviousChange, getSplitCellDialogConfig, getWatermarkFromState, insertImageFromFile, insertImageNode, rejectAllChanges, rejectChange, rejectChangeById, removeCommentMark, setImageWrapType, setWatermark, splitActiveTableCell };
