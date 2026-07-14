/**
 * Layout Bridge — measure, hit-test, and map between PM positions and pixels.
 *
 * @experimental Internal layer between the layout engine and rendering.
 * The named exports below are the public contract for adapter authors,
 * but the API is still evolving and may change in minor releases until
 * a third-party adapter validates it.
 * @packageDocumentation
 * @public
 */
export { ToFlowBlocksOptions, convertBorderSpecToLayout, resetBlockIdCounter, resolveListTemplate, toFlowBlocks } from './toFlowBlocks.js';
import { TableBlock, FlowBlock, Fragment, Measure, Page, TableFragment, TableMeasure, ParagraphBlock, ParagraphMeasure, Layout, PageMargins, ColumnLayout } from '../layout-engine/types.js';
export { FontMetrics, FontStyle, MeasureParagraphOptions, RunMeasurement, TextMeasurement, buildFontString, clampFloatingWrapMargins, clearAllCaches, clearFontMetricsCache, clearParagraphMeasureCache, clearTextWidthCache, findCharacterAtX, getCachedFontMetrics, getCachedParagraphMeasure, getCachedTextWidth, getCanvasContext, getFontCacheSize, getFontMetrics, getParagraphCacheSize, getRunCharWidths, getTextCacheSize, getTotalCacheSize, getXForCharacter, halfPtToPx, hashParagraphBlock, measureParagraph, measureParagraphs, measureRun, measureText, measureTextWidth, ptToPx, pxToHalfPt, pxToPt, pxToTwips, resetCanvasContext, setCachedFontMetrics, setCachedParagraphMeasure, setCachedTextWidth, setFontCacheSize, setParagraphCacheSize, setTextCacheSize, twipsToPx } from './measuring/index.js';
export { F as FloatPageGeometry, a as FloatingExclusionRect, b as FloatingImageZone, c as FloatingLineSegmentZone, M as MeasureBlockFn, g as getFloatingMargins, m as measureBlocksWithFloats, r as rectsToFloatingZones } from '../measureBlocksPipeline-iYNzAH-o.js';
export { DomCaretPosition, DomSelectionRect, clickToPositionDom, clipRectToTableWindow, getCaretPositionFromDom, getSelectionRectsFromDom, clickToPositionDom as mouseToPosition } from './clickToPositionDom.js';
import { EditorState } from 'prosemirror-state';
export { C as ConvertFootnoteOptions, a as ConvertHeaderFooterOptions, k as FOOTNOTE_COLUMN_GAP_PX, F as FOOTNOTE_SEPARATOR_HEIGHT, b as FootnoteRefLocation, H as HeaderFooterMetrics, M as MAX_FOOTNOTE_LAYOUT_PASSES, c as MeasureBlocksFn, S as StabilizeFootnoteLayoutArgs, d as StabilizeFootnoteLayoutResult, l as applyFootnotePresentation, e as buildFootnoteContentMap, f as buildFootnoteRenderItems, g as calculateFootnoteReservedHeights, n as calculateHeaderFooterVisualBounds, h as collectFootnoteRefs, o as computeHfCaretRectFromView, p as computeHfSelectionRectsFromView, q as contributesToHeaderFooterFlowHeight, r as convertFootnoteToContent, t as convertHeaderFooterPmDocToContent, i as convertHeaderFooterToContent, u as distributeFootnotesIntoColumns, j as footnoteReservedHeightsEqual, v as invalidateHfDomCache, m as mapFootnotesToPages, w as normalizeHeaderFooterMeasureBlocks, x as resolveHeaderFooterVisualTop, s as stabilizeFootnoteLayout } from '../headerFooterLayout-BbRwTj_j.js';
import { H as HeaderFooterContent } from '../footnotes-DMsicPGd.js';
export { TABLE_INSERT_EDGE_PROXIMITY, TABLE_INSERT_HIDE_DELAY_MS, TableInsertHoverHit, TableInsertHoverInput, detectTableInsertHover } from './tableInsertHover.js';
import { SectionLayoutConfig } from '../layout-engine/index.js';
import { Document } from '../types/document.js';
import { S as SectionProperties, y as HeaderFooter } from '../content-B8ScSBzC.js';
import 'prosemirror-model';
import '../styles-BGGjYqnf.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../lists-Bn29SzeS.js';
import '../docx/wrapTypes.js';
import '../watermark-D90356ZM.js';
import '../anchoredObjectPosition-CS-8BfhO.js';
import 'prosemirror-view';

/**
 * @internal Helpers for resolving DOCX table-width metadata into pixel widths.
 */

/**
 * Resolve a DOCX width pair to pixels. `pct` values are 50ths of a percent
 * (ECMA-376 §17.18.111 — 5000 means 100%). `dxa` / `auto` / unset are twips.
 *
 * @internal
 */
declare function resolveTableWidthPx(value: number | undefined, widthType: string | undefined, parentWidth: number): number | undefined;
/** A cell with its resolved grid position (column index honoring spans). */
interface ResolvedGridCell {
    rowIndex: number;
    cellIndex: number;
    columnIndex: number;
    colSpan: number;
    rowSpan: number;
}
/**
 * Resolve every cell's grid column index, accounting for `colSpan` and the
 * columns occupied by vertically-merged (`rowSpan`) cells from earlier rows.
 *
 * Single source of truth for table grid geometry — the measurer, the painter,
 * and the row-break paginator all consume this so they agree on which column a
 * cell lives in. Width-free on purpose: callers multiply `columnIndex` by their
 * own (possibly scaled) column widths to get an x offset.
 *
 * @internal
 */
declare function resolveCellGrid(tableBlock: TableBlock): ResolvedGridCell[];
/** Total grid columns, derived from the widest row's accumulated colSpans. */
declare function countTableColumns(tableBlock: TableBlock): number;
/**
 * Make `columnWidths` exactly `colCount` long with every entry positive.
 * Missing trailing columns inherit the average of existing positives; zero
 * or negative entries split the leftover `targetWidth` evenly. Callers
 * scale down totals that exceed the target — this helper only fills gaps.
 */
declare function normalizeTableColumnWidths(columnWidths: number[], colCount: number, targetWidth: number): number[];
/**
 * Resolve a table's per-column pixel widths from its grid metadata and width
 * budget — the width half of `measureTableBlock`, with NO cell-content
 * measurement. Factored out so a caller that only needs widths (e.g. deciding
 * whether a floating table is effectively full-width, before the main measure
 * pass) doesn't have to measure every cell.
 *
 * @internal
 */
declare function resolveTableColumnWidths(tableBlock: TableBlock, contentWidth: number): number[];
/**
 * Total pixel width of a table — sum of its resolved column widths, falling
 * back to the explicit table width or the content-width budget. No cell-content
 * measurement, so it is safe to call before the main measure pass. Mirrors the
 * `totalWidth` that `measureTableBlock` produces.
 *
 * @internal
 */
declare function resolveTableTotalWidthPx(tableBlock: TableBlock, contentWidth: number): number;

/**
 * Detect "block-like" floating tables.
 *
 * A DOCX table with `w:tblpPr` is a positioned ("floating") table — text wraps
 * around it. But contract templates routinely make a full-width form table
 * floating (often a paste artifact), leaving NO horizontal room for text on
 * either side. Word and Google Docs treat such a table as a normal in-flow
 * block: it paginates across pages and following content flows beneath it.
 *
 * Our floating path instead places the whole table as one fragment (no
 * pagination — it overflows past the page) AND registers a text-wrap zone the
 * height of the table, so the next paragraph's first line is pushed past the
 * entire table and lands off-page. Demoting these tables to inline (see
 * `computeLayout`) fixes both: they paginate via `layoutTable` and create no
 * wrap zone.
 *
 * The predicate mirrors the geometry in `layoutFloatingTable` (the existing
 * block-like cursor-advance) and `extractFloatingTableZone`; keep the three in
 * sync. Narrow floating tables — where text genuinely wraps beside them — fall
 * below the threshold and stay floating.
 */

/**
 * True when a floating table is effectively full-width — it leaves less than
 * {@link MIN_WRAP_SEGMENT_WIDTH} of usable text room on BOTH sides, so no text
 * can wrap beside it. Returns false for non-floating tables and for narrow
 * floats that leave real wrap room.
 *
 * @internal
 */
declare function isBlockLikeFloatingTable(block: TableBlock, contentWidth: number): boolean;
/**
 * In place, clear `floating` on every block-like floating table so the rest of
 * the pipeline treats it as a normal inline table — it then paginates via
 * `layoutTable` and registers no text-wrap zone. `blockWidths[i]` is the
 * block's section content width (from `computePerBlockWidths`); `fallbackWidth`
 * covers a missing entry. Demoted entries are replaced with shallow copies, so
 * any cached source block the caller holds is left untouched.
 *
 * @internal
 */
declare function demoteBlockLikeFloatingTables(blocks: FlowBlock[], blockWidths: number[], fallbackWidth: number): void;

/**
 * Hit Testing Utilities
 *
 * Maps click coordinates to pages, fragments, and document positions.
 * Used by the paged editor to translate user clicks into PM positions.
 */

/**
 * A 2D point in page coordinate space.
 */
type Point = {
    x: number;
    y: number;
};
/**
 * Result of hit-testing to find which page contains a point.
 */
type PageHit = {
    /** Index of the page (0-based). */
    pageIndex: number;
    /** The page that was hit. */
    page: Page;
    /** Y position relative to the page top. */
    pageY: number;
};
/**
 * Result of hit-testing to find which fragment contains a point.
 */
type FragmentHit = {
    /** The fragment that was hit. */
    fragment: Fragment;
    /** The corresponding flow block. */
    block: FlowBlock;
    /** The corresponding measurement. */
    measure: Measure;
    /** Page index (0-based). */
    pageIndex: number;
    /** Y position relative to the fragment top. */
    localY: number;
    /** X position relative to the fragment left. */
    localX: number;
};
/**
 * Result of hit-testing a table to find the cell.
 */
type TableCellHit = {
    /** The table fragment. */
    fragment: TableFragment;
    /** The table block. */
    block: TableBlock;
    /** The table measure. */
    measure: TableMeasure;
    /** Page index (0-based). */
    pageIndex: number;
    /** Row index (0-based). */
    rowIndex: number;
    /** Column index (0-based). */
    colIndex: number;
    /** Paragraph block within the cell (first one). */
    cellBlock?: ParagraphBlock;
    /** Paragraph measure within the cell. */
    cellMeasure?: ParagraphMeasure;
    /** X position relative to cell content area. */
    cellLocalX: number;
    /** Y position relative to cell content area. */
    cellLocalY: number;
};
/**
 * Hit-test pages to find which page contains a given Y coordinate.
 *
 * Pages are stacked vertically with optional gaps between them.
 * If the point is in a gap, returns the nearest page.
 *
 * @param layout - The layout containing pages.
 * @param point - The point to test (only Y is used for page lookup).
 * @returns PageHit if a page was found, null otherwise.
 */
declare function hitTestPage(layout: Layout, point: Point): PageHit | null;
/**
 * Calculate the Y offset of a page from the top of the document.
 */
declare function getPageTop(layout: Layout, pageIndex: number): number;
/**
 * Get the page index at a specific Y coordinate.
 * Returns null if the coordinate is outside all pages.
 */
declare function getPageIndexAtY(layout: Layout, y: number): number | null;
/**
 * Hit-test fragments on a page to find which fragment contains a point.
 *
 * Fragments are checked in visual order (sorted by Y then X).
 * Only considers paragraph and table fragments (not images for now).
 *
 * @param pageHit - The page hit result from hitTestPage.
 * @param blocks - All flow blocks in the document.
 * @param measures - All measurements corresponding to blocks.
 * @param pagePoint - Point in page-relative coordinates.
 * @returns FragmentHit if a fragment was found, null otherwise.
 */
declare function hitTestFragment(pageHit: PageHit, blocks: FlowBlock[], measures: Measure[], pagePoint: Point): FragmentHit | null;
/**
 * Hit-test to find image fragments (they may overlap other content).
 */
declare function hitTestImageFragment(pageHit: PageHit, blocks: FlowBlock[], measures: Measure[], pagePoint: Point): FragmentHit | null;
/**
 * Hit-test within a table fragment to find the specific cell.
 *
 * @param pageHit - The page hit result.
 * @param blocks - All flow blocks.
 * @param measures - All measurements.
 * @param pagePoint - Point in page-relative coordinates.
 * @returns TableCellHit if a table cell was found, null otherwise.
 */
declare function hitTestTableCell(pageHit: PageHit, blocks: FlowBlock[], measures: Measure[], pagePoint: Point): TableCellHit | null;
/**
 * Combined result of all hit testing.
 */
type HitTestResult = {
    /** Page that was hit. */
    pageHit: PageHit | null;
    /** Fragment that was hit (if any). */
    fragmentHit: FragmentHit | null;
    /** Table cell that was hit (if fragment is a table). */
    tableCellHit: TableCellHit | null;
};
/**
 * Perform complete hit testing from document coordinates to the most specific element.
 *
 * @param layout - The layout containing pages.
 * @param blocks - All flow blocks.
 * @param measures - All measurements.
 * @param point - Point in document coordinates (Y is cumulative from document top).
 * @returns Complete hit test result.
 */
declare function hitTest(layout: Layout, blocks: FlowBlock[], measures: Measure[], point: Point): HitTestResult;
/**
 * Get the total document height (all pages + gaps).
 */
declare function getTotalDocumentHeight(layout: Layout): number;
/**
 * Get Y coordinate to scroll to for a specific page.
 */
declare function getScrollYForPage(layout: Layout, pageIndex: number): number;
/**
 * Get page bounds (top and bottom Y coordinates).
 */
declare function getPageBounds(layout: Layout, pageIndex: number): {
    top: number;
    bottom: number;
} | null;

/**
 * Click-to-Position Mapping
 *
 * Maps click coordinates within a layout to ProseMirror document positions.
 * Uses geometry-based calculation with canvas text measurement for accuracy.
 *
 * The main entry point `clickToPosition` takes fragment hit data and local
 * coordinates and returns the PM position at the click point.
 */

/**
 * Result of click-to-position mapping.
 */
type PositionResult = {
    /** ProseMirror document position. */
    pmPosition: number;
    /** Character offset within the line (for debugging). */
    charOffset: number;
    /** Line index within the paragraph. */
    lineIndex: number;
};
/**
 * Map a click within a paragraph fragment to a PM position.
 *
 * @param fragmentHit - The fragment hit result from hitTestFragment.
 * @returns Position result, or null if mapping fails.
 */
declare function clickToPositionInParagraph(fragmentHit: FragmentHit): PositionResult | null;
/**
 * Map a click within a table cell to a PM position.
 *
 * @param tableCellHit - The table cell hit result from hitTestTableCell.
 * @returns PM position, or null if mapping fails.
 */
declare function clickToPositionInTableCell(tableCellHit: TableCellHit): number | null;
/**
 * Main entry point: Map a click to a PM position.
 *
 * This function takes the result of hit testing and returns the PM position.
 *
 * @param fragmentHit - Fragment hit from hitTestFragment.
 * @param tableCellHit - Optional table cell hit from hitTestTableCell.
 * @returns PM position, or null if mapping fails.
 */
declare function clickToPosition(fragmentHit: FragmentHit | null, tableCellHit?: TableCellHit | null): number | null;
/**
 * Map a PM position to X coordinates within a line (for caret positioning).
 *
 * @param block - The paragraph block.
 * @param measure - The paragraph measure.
 * @param pmPosition - The PM position to map.
 * @param fragmentWidth - Width of the fragment.
 * @returns X coordinate relative to fragment start, or null if not found.
 */
declare function positionToX(block: ParagraphBlock, measure: ParagraphMeasure, pmPosition: number, _fragmentWidth: number): {
    x: number;
    lineIndex: number;
} | null;
/**
 * Get the bounding rect for a PM position (for caret rendering).
 */
declare function getPositionRect(block: ParagraphBlock, measure: ParagraphMeasure, pmPosition: number, fragmentX: number, fragmentY: number, fragmentWidth: number, fromLine: number): {
    x: number;
    y: number;
    height: number;
} | null;

/**
 * Apply the visual cell-selection highlight class onto rendered table cells.
 *
 * The layout painter renders pages as static DOM with `data-pm-start` on each
 * table cell. This walk translates a ProseMirror `CellSelection` into the
 * `.layout-table-cell-selected` class on the matching painted cells. Shared by
 * the React and Vue adapters.
 */

/**
 * Apply the `.layout-table-cell-selected` class to painted layout cells
 * matching a CellSelection in the PM state. Clears the class everywhere
 * (within scope) first so toggling off (or moving to a TextSelection) erases
 * prior highlights.
 *
 * Duck-types CellSelection via `$anchorCell` / `forEachCell` rather than
 * `instanceof` to dodge bundling issues across `prosemirror-tables` copies
 * (the same trick used inline in the selection-overlay update).
 */
declare function applyCellSelectionHighlight(pagesContainer: HTMLElement, state: EditorState, options?: {
    scope?: 'body' | 'header' | 'footer';
}): void;

/**
 * Selection Rectangles
 *
 * Converts ProseMirror selection ranges into screen rectangles for rendering
 * selection highlights and the caret cursor.
 *
 * The main function `selectionToRects` takes a PM range and returns an array
 * of rectangles that cover the selected text across all pages and fragments.
 */

/**
 * A rectangle representing part of a selection.
 */
type SelectionRect = {
    /** X coordinate in container space. */
    x: number;
    /** Y coordinate in container space. */
    y: number;
    /** Width of the rectangle. */
    width: number;
    /** Height of the rectangle (typically line height). */
    height: number;
    /** Page index (0-based). */
    pageIndex: number;
};
/**
 * Caret position for collapsed selection.
 */
type CaretPosition = {
    /** X coordinate in container space. */
    x: number;
    /** Y coordinate in container space. */
    y: number;
    /** Height of the caret (line height). */
    height: number;
    /** Page index (0-based). */
    pageIndex: number;
};
/**
 * Convert a ProseMirror selection range to screen rectangles.
 *
 * @param layout - The document layout.
 * @param blocks - All flow blocks.
 * @param measures - All measurements.
 * @param from - Start PM position.
 * @param to - End PM position.
 * @returns Array of rectangles covering the selection.
 */
declare function selectionToRects(layout: Layout, blocks: FlowBlock[], measures: Measure[], from: number, to: number): SelectionRect[];
/**
 * Get caret position for a collapsed selection.
 *
 * @param layout - The document layout.
 * @param blocks - All flow blocks.
 * @param measures - All measurements.
 * @param pmPosition - The PM position.
 * @param startPageIndex - Page index to begin the search from. Pages/fragments
 *   are ordered by PM position, so callers resolving many positions in
 *   ascending order (e.g. sidebar anchors) can pass the previously matched
 *   page to turn a per-call O(pages) scan into an amortized O(1) advance.
 *   Defaults to 0 (full scan). Must be ≤ the page containing `pmPosition`.
 * @returns Caret position, or null if not found.
 */
declare function getCaretPosition(layout: Layout, blocks: FlowBlock[], measures: Measure[], pmPosition: number, startPageIndex?: number): CaretPosition | null;
/**
 * Check if a selection spans multiple pages.
 *
 * @param rects - Selection rectangles.
 * @returns True if selection spans multiple pages.
 */
declare function isMultiPageSelection(rects: SelectionRect[]): boolean;
/**
 * Get selection rectangles grouped by page.
 *
 * @param rects - Selection rectangles.
 * @returns Map of page index to rectangles on that page.
 */
declare function groupRectsByPage(rects: SelectionRect[]): Map<number, SelectionRect[]>;

/**
 * Header/footer body-margin extension.
 *
 * Word grows the header (or footer) band when its in-flow content is taller
 * than the authored top (or bottom) margin minus the header/footer distance,
 * pushing the body text down (or up). This module owns that computation so the
 * React and Vue adapters share one implementation instead of byte-identical
 * inline copies (the layout pipelines were drifting candidates — see
 * `docx-editor` engine-unification work, issue #696).
 *
 * Two correctness rules live here:
 *
 *  1. The band height is driven by `HeaderFooterContent.flowHeight` (in-flow
 *     content only), NOT `height` / `visualBottom`. A page/margin-anchored
 *     float — e.g. a full-page letterhead anchored in a header — is positioned
 *     on the page and does not push the body in Word. Counting it inflated the
 *     effective top margin past the page on real-world templates, so the
 *     paginator hard-threw "page size and margins yield no content area" and
 *     the document rendered blank (issue #705).
 *
 *  2. A clamp guarantees `top + bottom` never consumes the whole page, so a
 *     pathological in-flow header degrades to a thin content band with a
 *     warning instead of aborting pagination.
 */

/** @public */
interface ExtendMarginsForHeaderFooterInput {
    pageSize: {
        w: number;
        h: number;
    };
    /** Body fallback margins. */
    margins: PageMargins;
    /** Final-section margins (last `sectPr`). */
    finalMargins: PageMargins;
    /**
     * Body flow blocks. Each `sectionBreak` block's `margins` is extended IN
     * PLACE so multi-section documents paginate with the same band growth (the
     * layout engine prefers `sectionBreak.margins` over the body fallback).
     */
    bodyBlocks?: FlowBlock[];
    /** Header variants in play this layout (e.g. default + first-page). */
    headers?: Array<HeaderFooterContent | undefined>;
    /** Footer variants in play this layout. */
    footers?: Array<HeaderFooterContent | undefined>;
    /** Optional diagnostic sink for the clamp (adapters pass `console.warn`). */
    warn?: (message: string) => void;
}
/** @public */
interface ExtendMarginsForHeaderFooterResult {
    margins: PageMargins;
    finalMargins: PageMargins;
}
/**
 * Extend body margins so the body clears the header/footer bands, mirroring
 * Word. Returns new `margins` / `finalMargins`; mutates `sectionBreak.margins`
 * in place. When no extension is needed the original objects are returned
 * unchanged.
 *
 * @public
 */
declare function extendMarginsForHeaderFooter(input: ExtendMarginsForHeaderFooterInput): ExtendMarginsForHeaderFooterResult;

/**
 * Body-scoped DOM lookups for `data-pm-start` / `data-pm-end` markers.
 *
 * Headers and footers are parsed via a separate ProseMirror document
 * (see `convertHeaderFooterToContent`), so HF runs end up with PM
 * positions that overlap body PM positions. Any container-level
 * `querySelector(...)` that doesn't restrict to `.layout-page-content`
 * will match both trees and produce wrong results: phantom selection
 * rectangles in headers, scroll-restore latching onto an HF span,
 * caret resolution returning an HF position, etc.
 *
 * These helpers centralize the body-only scope so call sites can't
 * forget the prefix. The match for HF lookups is `.layout-page-header`
 * / `.layout-page-footer`; HF call sites should write their own queries
 * scoped to those classes.
 */
/**
 * All body-tree run spans carrying both `data-pm-start` and `data-pm-end`.
 *
 * This is the workhorse for caret resolution and selection-rect painting.
 */
declare function findBodyPmSpans(container: ParentNode): HTMLElement[];
/**
 * All body-tree empty-paragraph runs. Used as a caret fallback when a
 * paragraph has no painted text spans.
 */
declare function findBodyEmptyRuns(container: ParentNode): HTMLElement[];
/**
 * All body-tree elements carrying a `data-pm-start` attribute.
 *
 * Includes paragraph elements as well as run spans, which is what scroll-
 * anchor recovery needs. Distinct from {@link findBodyPmSpans}, which
 * filters down to run spans only.
 */
declare function findBodyPmAnchors(container: ParentNode): HTMLElement[];
/**
 * First body-tree element whose `data-pm-start` exactly matches `pmStart`.
 *
 * Used for scroll anchors and image `NodeSelection` resolution where the
 * caller already knows the exact PM position it wants to find. Returns
 * `null` for non-finite inputs so callers don't need to guard.
 */
declare function findBodyPmAnchor(container: ParentNode, pmStart: number): HTMLElement | null;

/**
 * Shared table measurement helper.
 *
 * Both React's PagedEditor and Vue's useDocxEditor measure tables the
 * same way: resolve column widths through `tableWidthUtils`, track
 * column occupancy across vertically merged cells, then call back into
 * the adapter's `measureBlock` for each cell's contained blocks.
 *
 * This module lives in core (not in either adapter) so a fix on one
 * side automatically applies to both. The recursive cell-content
 * measurement is delegated to the caller via `measureBlock` because
 * each adapter's block coverage differs (React supports floating
 * zones, footnotes, textBoxes; Vue is a subset).
 */

/**
 * Visual height of a single block inside a table cell.
 *
 * A one-line paragraph that contains only image runs is laid out at the
 * image's intrinsic height (plus the paragraph's explicit spacing), not
 * a full text line — matching Word's per-cell layout. Everything else
 * uses the measured `totalHeight` / `height`.
 */
declare function measureTableCellBlockVisualHeight(block: FlowBlock, blockMeasure: Measure): number;
/**
 * Measure a `TableBlock` against a content-width budget.
 *
 * `measureBlock` is the per-cell-content measurement callback the
 * adapter uses for everything inside a cell. The adapter passes its
 * own `measureBlock` so block coverage stays per-renderer.
 */
declare function measureTableBlock(tableBlock: TableBlock, contentWidth: number, measureBlock: (block: FlowBlock, contentWidth: number) => Measure): TableMeasure;

/**
 * Section-properties → page geometry helpers.
 *
 * Both adapters need the same translation from `SectionProperties` (twips
 * + DOCX field names) into the layout engine's pixel-shaped `PageSize`
 * and `PageMargins`. Living in core eliminates the subtle React/Vue
 * drift that grows when each adapter ships its own twipsToPixels math.
 */

/** US Letter at 96 DPI — Word's default page size. */
declare const DEFAULT_PAGE_WIDTH_PX = 816;
/** US Letter at 96 DPI. */
declare const DEFAULT_PAGE_HEIGHT_PX = 1056;
/** 1 inch at 96 DPI — Word's default body margins. */
declare const DEFAULT_BODY_MARGIN_PX = 96;
/** Word's default `headerDistance` / `footerDistance` (0.5in = 48px). */
declare const DEFAULT_HF_DISTANCE_PX = 48;
/** Convert twips to pixels (1 twip = 1/20 point, 96 pixels per inch). */
declare function twipsToPixels(twips: number): number;
/**
 * Convert SectionProperties page size (twips) → pixel `{ w, h }`.
 *
 * Page size is a SIZE: a literal `0` (malformed `w:pgSz`) defaults to Letter
 * rather than rendering a zero-area page — so the truthy guard is intentional
 * here (contrast `getMargins`, where `0` is honored). See `twipsToPxOr`.
 */
declare function getPageSize(sp: SectionProperties | null | undefined): {
    w: number;
    h: number;
};
/**
 * Convert SectionProperties margins (twips) → pixel `PageMargins`.
 *
 * Every distance is an OFFSET, so an explicit `0` is honored (full-bleed body
 * margins; a header/footer pinned to the page edge — issue #740). Only an
 * ABSENT distance falls back to Word's default. `header`/`footer` default to
 * 48px (Word's 0.5in) so the HF margin-extension math needn't special-case
 * undefined.
 */
declare function getMargins(sp: SectionProperties | null | undefined): PageMargins;
/**
 * Resolve the HeaderFooter pair (default + first-page) for a section.
 *
 * Mirrors React's lookup in DocxEditor.tsx: read `pkg.headers`/`footers`
 * (Maps keyed by rId), resolve through `sp.headerReferences` /
 * `footerReferences`. When `titlePg` is unset and only `first` HFs
 * exist, they serve as the default — same Word fallback both adapters
 * have shipped.
 */
declare function resolveHeaderFooter(doc: Document | null, sp: SectionProperties | null | undefined): {
    header: HeaderFooter | null;
    footer: HeaderFooter | null;
    firstHeader: HeaderFooter | null;
    firstFooter: HeaderFooter | null;
};
/**
 * Extract column layout from section properties.
 * Returns undefined for single-column (default) to avoid unnecessary paginator overhead.
 */
declare function getColumns(sectionProps: SectionProperties | null | undefined): ColumnLayout | undefined;
declare function columnWidthForSection(config: SectionLayoutConfig): number;
/**
 * Compute per-block measurement widths by scanning for section breaks.
 * Blocks must be measured with the page width/margins/columns of their own
 * section so that the layout engine can paginate them against the right
 * geometry without remeasuring.
 */
declare function computePerBlockWidths(blocks: FlowBlock[], initialConfig: SectionLayoutConfig, finalConfig: SectionLayoutConfig): number[];

export { type CaretPosition, DEFAULT_BODY_MARGIN_PX, DEFAULT_HF_DISTANCE_PX, DEFAULT_PAGE_HEIGHT_PX, DEFAULT_PAGE_WIDTH_PX, type ExtendMarginsForHeaderFooterInput, type ExtendMarginsForHeaderFooterResult, type FragmentHit, type HitTestResult, type PageHit, type Point, type PositionResult, type ResolvedGridCell, type SelectionRect, type TableCellHit, applyCellSelectionHighlight, clickToPosition, clickToPositionInParagraph, clickToPositionInTableCell, columnWidthForSection, computePerBlockWidths, countTableColumns, demoteBlockLikeFloatingTables, extendMarginsForHeaderFooter, findBodyEmptyRuns, findBodyPmAnchor, findBodyPmAnchors, findBodyPmSpans, getCaretPosition, getColumns, getMargins, getPageBounds, getPageIndexAtY, getPageSize, getPageTop, getPositionRect, getScrollYForPage, getTotalDocumentHeight, groupRectsByPage, hitTest, hitTestFragment, hitTestImageFragment, hitTestPage, hitTestTableCell, isBlockLikeFloatingTable, isMultiPageSelection, measureTableBlock, measureTableCellBlockVisualHeight, normalizeTableColumnWidths, positionToX, resolveCellGrid, resolveHeaderFooter, resolveTableColumnWidths, resolveTableTotalWidthPx, resolveTableWidthPx, selectionToRects, twipsToPixels };
