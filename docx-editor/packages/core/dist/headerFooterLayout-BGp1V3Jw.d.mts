import { Node } from 'prosemirror-model';
import { EditorView } from 'prosemirror-view';
import { FlowBlock, Measure, BlockId, LayoutOptions, FootnoteContent, Layout, Page, PageMargins, ImageRun } from './layout-engine/types.mjs';
import { F as Footnote, y as HeaderFooter } from './content-BZ9rYecc.mjs';
import { S as StyleDefinitions, T as Theme } from './styles-BNjUANte.mjs';
import { F as FootnoteRenderItem, H as HeaderFooterContent } from './footnotes-BZ24OTAT.mjs';
import { Document } from './types/document.mjs';

/**
 * Footnote Layout Utilities
 *
 * Footnote/endnote rendering pipeline plus page-mapping helpers:
 * - scanning FlowBlocks for footnote references and their PM positions
 * - mapping references to the page that ends up containing them
 * - converting a Footnote → FootnoteContent via the body pipeline
 *   (footnoteToProseDoc → toFlowBlocks → caller-supplied measureBlocks)
 * - reserving per-page footnote area heights for layout
 *
 * Everything that's pure OOXML / FlowBlock semantics lives here so the
 * React, Vue, and any future adapters can share the conversion logic
 * and just supply their own measurement function (which depends on
 * platform-specific Canvas/font metrics).
 */

/** Separator line height + vertical padding in pixels. */
declare const FOOTNOTE_SEPARATOR_HEIGHT = 12;
/**
 * Gutter between footnote columns when `w15:footnoteColumns` > 1, in pixels
 * (≈ 0.25in). Shared by the reserved-height/measurement path (core) and the
 * footnote painter so a footnote measured at column width paints into a column
 * of exactly that width. Single-column footnotes never consult it.
 */
declare const FOOTNOTE_COLUMN_GAP_PX = 24;
/**
 * Hard cap on the multi-pass footnote layout loop. Reserving footnote
 * space can move a reference to another page, so adapters keep remapping
 * until the page→height contract is stable. Dense layouts converge in
 * 2–3 passes in practice; 6 is a safe ceiling.
 */
declare const MAX_FOOTNOTE_LAYOUT_PASSES = 6;
/**
 * Compare two per-page footnote reservation maps. Used by the React +
 * Vue adapters to detect when the multi-pass loop has converged.
 */
declare function footnoteReservedHeightsEqual(a: Map<number, number>, b: Map<number, number>): boolean;
/**
 * Where a footnote reference lives, as found by {@link collectFootnoteRefs}.
 *
 * `pmPos` alone is enough to attribute a reference to a page for ordinary
 * (paragraph) content, whose fragments carry a per-page pm sub-range. A table
 * is different: it splits across pages by ROW, but every `TableFragment` keeps
 * the whole table's `pmStart`/`pmEnd` (those drive selection mapping and must
 * not be narrowed). So for a reference authored inside a table cell we also
 * record the OUTERMOST table's id and the index of the row that contains it,
 * letting {@link mapFootnotesToPages} attribute the reference to the page that
 * actually laid out that row.
 */
type FootnoteRefLocation = {
    footnoteId: number;
    pmPos: number;
    /** Id of the outermost enclosing table block, when the ref is in a table cell. */
    tableBlockId?: BlockId;
    /** Index (into the outermost table's `rows`) of the row holding the ref. */
    rowIndex?: number;
};
/**
 * Scan FlowBlocks for runs with footnoteRefId set.
 * Returns a list of {@link FootnoteRefLocation} in document order.
 *
 * Recurses into container blocks (table cells, text boxes) so footnote
 * references authored anywhere in the body reach the page-reservation
 * pass. Without this, a `footnoteRefId` nested inside a table cell never
 * gets mapped to a page and the per-page `.layout-footnote-area` silently
 * drops that entry even though the body still renders the in-line ref
 * marker.
 *
 * For refs inside a table, the OUTERMOST table's id and row index are
 * recorded (a nested table keeps the outer context, since the outer row is
 * what the paginator splits into per-page fragments).
 */
declare function collectFootnoteRefs(blocks: FlowBlock[]): FootnoteRefLocation[];
/**
 * After layout, determine which footnotes appear on which pages.
 * Checks each page's fragments to see if any footnoteRef PM positions fall within.
 *
 * Returns Map<pageNumber, footnoteId[]> in document order.
 */
declare function mapFootnotesToPages(pages: Page[], footnoteRefs: FootnoteRefLocation[]): Map<number, number[]>;
/**
 * Footnote-specific block normalization. Mirrors the spirit of
 * `normalizeHeaderFooterMeasureBlocks`: post-process the body-pipeline
 * output for a single footnote so it carries the correct visual prefix
 * (its display number, rendered as a superscript) and a default 8pt font
 * for any run that didn't specify a size.
 *
 * The displayNumber is prepended onto the FIRST paragraph as a fresh
 * superscript text run — visually matches Word's footnote numbering
 * without disturbing the authored runs.
 *
 * Exported for callers that want to compose their own conversion
 * pipeline; `convertFootnoteToContent` calls it as part of its flow.
 */
declare function applyFootnotePresentation(blocks: FlowBlock[], displayNumber: number): FlowBlock[];
/**
 * Adapter-supplied block measurement function. The caller (React /
 * Vue / etc.) supplies its platform's measure routine — at minimum
 * paragraph + table + image + textBox — so this core helper stays
 * Canvas-free.
 */
type MeasureBlocksFn = (blocks: FlowBlock[], contentWidth: number) => Measure[];
/**
 * Options for {@link convertFootnoteToContent}.
 */
type ConvertFootnoteOptions = {
    /** The document's parsed style definitions, threaded into the body pipeline. */
    styles?: StyleDefinitions | null;
    /** Theme for resolving themed fills / fonts inside the footnote. */
    theme?: Theme | null;
    /** Measure callback supplied by the rendering adapter. */
    measureBlocks: MeasureBlocksFn;
    /**
     * Doc-level `w:defaultTabStop` (twips) from the body so list markers
     * inside footnotes honor the same tab grid.
     */
    defaultTabStopTwips?: number | null;
};
/**
 * Convert a Footnote to renderable FootnoteContent via the body pipeline:
 * `footnoteToProseDoc → toFlowBlocks → applyFootnotePresentation →
 * measureBlocks`. Pre-PR (#378) this lived in a hand-rolled shadow stack
 * that silently dropped non-paragraph content; routing through the body
 * pipeline gives footnotes full block-kind support — paragraph + table
 * + image + textBox + fields.
 */
declare function convertFootnoteToContent(footnote: Footnote, displayNumber: number, contentWidth: number, options: ConvertFootnoteOptions): FootnoteContent;
/**
 * Build footnote content for all footnotes referenced in the document.
 * Display numbers are assigned by first-appearance order (the same way
 * Word renders them).
 */
declare function buildFootnoteContentMap(footnotes: Footnote[], footnoteRefs: Array<{
    footnoteId: number;
}>, contentWidth: number, options: ConvertFootnoteOptions): Map<number, FootnoteContent>;
/**
 * Distribute footnote items across `columns` balanced columns, preserving
 * document order (footnotes must still read in numeric sequence). Items fill
 * the first column until it reaches the balanced target height (≈ total / N),
 * then spill into the next column — the same order-preserving balance Word
 * applies to its footnote columns, not a greedy shortest-column packing
 * (which would scramble the reading order).
 *
 * `columns <= 1` (the default for ordinary single-column footnotes) returns a
 * single column unchanged, so callers that never opt into multi-column
 * footnotes are byte-for-byte unaffected.
 *
 * Pure and shared by the reserved-height calculation (core) and the footnote
 * painter (layout-painter) so the reserved area and the rendered columns are
 * computed from the same partition.
 */
declare function distributeFootnotesIntoColumns<T extends {
    height: number;
}>(items: T[], columns: number): T[][];
/**
 * Calculate per-page footnote reserved heights.
 * Returns Map<pageNumber, reservedHeight>.
 *
 * With `columns > 1` the footnotes are balanced across that many columns and
 * the reserved height is the tallest column (plus the separator), since the
 * columns sit side by side — not the sum of every footnote height.
 */
declare function calculateFootnoteReservedHeights(pageFootnoteMap: Map<number, number[]>, footnoteContentMap: Map<number, {
    height: number;
}>, columns?: number): Map<number, number>;
interface StabilizeFootnoteLayoutArgs {
    blocks: FlowBlock[];
    measures: Measure[];
    layoutOpts: LayoutOptions;
    footnoteRefs: FootnoteRefLocation[];
    footnoteContentMap: Map<number, FootnoteContent>;
    /** First-pass layout already computed by the caller without reserved heights. */
    initialLayout: Layout;
    /**
     * Number of columns the footnote area is laid out in (`w15:footnoteColumns`).
     * Defaults to 1. When > 1, reserved heights are balanced across the columns
     * (tallest column wins) instead of summing every footnote, and the value is
     * written onto each footnote-bearing page as `page.footnoteColumns`.
     */
    footnoteColumns?: number;
}
interface StabilizeFootnoteLayoutResult {
    layout: Layout;
    pageFootnoteMap: Map<number, number[]>;
    /** True if the loop converged before hitting MAX_FOOTNOTE_LAYOUT_PASSES. */
    converged: boolean;
}
/**
 * Run the multi-pass footnote layout loop. Reserving footnote space on a
 * page can move a reference to another page, which changes the reservation,
 * which can move references again. Iterate until the page→height contract
 * is the same one used by the latest layout, or `MAX_FOOTNOTE_LAYOUT_PASSES`
 * passes have run.
 *
 * Lives in core so the React + Vue adapters call the same loop and stay in
 * lockstep on convergence behaviour. Writes `page.footnoteIds` onto each
 * page in the returned layout so renderers can paint footnote areas.
 */
declare function stabilizeFootnoteLayout(args: StabilizeFootnoteLayoutArgs): StabilizeFootnoteLayoutResult;
/**
 * Turn the page→footnote-id map into the per-page render payload that
 * `renderPages` consumes via `footnotesByPage`. Skips non-`normal` notes
 * (separators, continuation notices), reads the display number out of the
 * content map, and pulls plain text via `getFootnoteText`.
 *
 * Lives in core (not in either adapter) so React + Vue both call the
 * same helper — same rule as the rest of this module.
 */
declare function buildFootnoteRenderItems(pageFootnoteMap: Map<number, number[]>, footnoteContentMap: Map<number, FootnoteContent>, doc: Document | null): Map<number, FootnoteRenderItem[]>;

/**
 * Header / Footer Layout Utilities
 *
 * The header/footer rendering pipeline lives here so any rendering adapter
 * (React, Vue, etc.) can share the conversion logic and just supply its
 * platform-specific {@link MeasureBlocksFn}. Mirrors the footnote pipeline
 * in `footnoteLayout.ts`.
 *
 * Pipeline:
 *   HF.content → headerFooterToProseDoc → toFlowBlocks
 *     → measureBlocks (caller-supplied, Canvas-aware)
 *     → HeaderFooterContent (blocks, measures, height, visualTop/Bottom)
 *
 * The render side uses the normalized block list so paint and measurement stay
 * in lockstep. Visual-bounds calculation still inspects the original block
 * list because floating images can paint above/below the nominal flow box even
 * when they do not contribute to flow height.
 */

type HeaderFooterMetrics = {
    section: 'header' | 'footer';
    pageSize: {
        w: number;
        h: number;
    };
    margins: PageMargins;
};
declare function normalizeHeaderFooterMeasureBlocks(blocks: FlowBlock[]): FlowBlock[];
declare function resolveHeaderFooterVisualTop(run: ImageRun, paragraphY: number, flowHeight: number, metrics: HeaderFooterMetrics): number;
/**
 * Whether a header/footer block participates in the in-flow band height that
 * pushes the body margin.
 *
 * OOXML semantics: Word grows the header/footer band — and shifts body text —
 * based only on the story's in-flow content. A floating/anchored object
 * (`wp:anchor` DrawingML or an absolutely-positioned VML shape, e.g. a
 * full-page letterhead anchored to the page in a header) is removed from the
 * text flow and positioned on the page; it does NOT grow the band or push the
 * body. So only inline-flow blocks count here. Anchored image *runs* inside a
 * paragraph are likewise out of flow, but they don't contribute to the
 * paragraph's measured line height, so paragraphs need no special handling.
 *
 * @public
 */
declare function contributesToHeaderFooterFlowHeight(block: FlowBlock): boolean;
declare function calculateHeaderFooterVisualBounds(blocks: FlowBlock[], measures: Measure[], flowHeight: number, metrics: HeaderFooterMetrics): {
    visualTop: number;
    visualBottom: number;
};
type ConvertHeaderFooterOptions = {
    styles?: StyleDefinitions | null;
    theme?: Theme | null;
    measureBlocks: MeasureBlocksFn;
    /**
     * `w:defaultTabStop` (twips) read from `state.doc.attrs.defaultTabStopTwips`
     * on the body doc — HF content doesn't carry its own doc-level setting,
     * so pass it through so list markers inside headers/footers honor the
     * same tab grid as the body.
     */
    defaultTabStopTwips?: number | null;
};
/**
 * Convert HeaderFooter (document type) to HeaderFooterContent (render type).
 *
 * Routes through the same pipeline as the body: HF.content →
 * headerFooterToProseDoc → toFlowBlocks → measureBlocks. The inline editor
 * uses the same conversion chain, so block support (paragraph, table, image,
 * textBox, fields) and the inline editor's content stay in lockstep.
 */
declare function convertHeaderFooterToContent(headerFooter: HeaderFooter | null | undefined, contentWidth: number, metrics: HeaderFooterMetrics, options: ConvertHeaderFooterOptions): HeaderFooterContent | undefined;
/**
 * Same pipeline as {@link convertHeaderFooterToContent}, but starts from an
 * already-built ProseMirror document instead of `HeaderFooter.content`.
 *
 * The unified HF editing model (see `openspec/changes/unify-hf-editing/`)
 * maintains one persistent hidden PM EditorView per HF `rId`. The painter
 * reads from that EditorView's current `state.doc` rather than re-parsing
 * the Document-model `HeaderFooter` every layout pass — this is what
 * actually makes the painter and the editor stay in lockstep.
 *
 * `headerFooterToProseDoc` is still the right entry point when there is no
 * mounted PM for the slot (cold load, or rId not yet projected).
 *
 * @public
 */
declare function convertHeaderFooterPmDocToContent(pmDoc: Node, contentWidth: number, metrics: HeaderFooterMetrics, options: ConvertHeaderFooterOptions): HeaderFooterContent | undefined;
/**
 * Drop the cached HF host + span lists. Hosts/painters call this after
 * a repaint (or HF mode toggle) so the next caret / selection compute
 * re-walks the DOM. Public so adapters can call it from their painter
 * commit signal.
 *
 * @public
 */
declare function invalidateHfDomCache(): void;
/**
 * TODO(unify-hf-editing follow-up): this function duplicates the
 * span-walking + Range/TreeWalker logic in
 * `packages/react/src/components/DocxEditor/internals/domSelection.ts:getCaretFromDom`
 * (body). The body's helper is scoped to `.layout-page-content` via
 * `findBodyPmSpans`; we walk the same shape scoped to `.layout-page-header /
 * .layout-page-footer` here. Unification path:
 *   1. Add `findHfPmSpans` / `findHfEmptyRuns` mirrors next to the body
 *      ones in `packages/core/src/layout-bridge/findBodyPmSpans.ts`.
 *   2. Add `scope: 'body' | 'hf'` param to `getCaretFromDom` +
 *      `computeSelectionRectsFromDom`; switch the helper internally.
 *   3. Move the (now scope-aware) helpers into core so React + Vue both
 *      call them.
 *   4. Delete this function and `computeHfSelectionRectsFromView` —
 *      `DocxEditorPagedArea` calls `getCaretFromDom(scope: 'hf', ...)`.
 * Reviewer estimate: ~30 LOC net deletion + body↔HF parity for free
 * (lineHeight from `.layout-line` ancestor, empty-paragraph fallback
 * via `findBodyEmptyRuns`, etc.). Deferred because it's a multi-file
 * shape change that doesn't affect observable behavior.
 *
 * @public
 */
declare function computeHfCaretRectFromView(view: EditorView, section: 'header' | 'footer', doc?: globalThis.Document): {
    top: number;
    left: number;
    height: number;
} | null;
/**
 * Selection-rect set for a non-empty HF selection, projected against the
 * painted HF spans. Mirror of `computeSelectionRectsFromDom` but scoped to
 * `.layout-page-header` / `.layout-page-footer` instead of the body. Used
 * so the painter draws a visible highlight when the user drag-selects
 * inside a header/footer in edit mode.
 *
 * Returns viewport-relative `{top, left, width, height}` rects. Empty
 * array when selection is collapsed or no painted spans overlap the range.
 *
 * @public
 */
declare function computeHfSelectionRectsFromView(view: EditorView, section: 'header' | 'footer', doc?: globalThis.Document): Array<{
    top: number;
    left: number;
    width: number;
    height: number;
}>;

export { type ConvertFootnoteOptions as C, FOOTNOTE_SEPARATOR_HEIGHT as F, type HeaderFooterMetrics as H, MAX_FOOTNOTE_LAYOUT_PASSES as M, type StabilizeFootnoteLayoutArgs as S, type ConvertHeaderFooterOptions as a, type FootnoteRefLocation as b, type MeasureBlocksFn as c, type StabilizeFootnoteLayoutResult as d, buildFootnoteContentMap as e, buildFootnoteRenderItems as f, calculateFootnoteReservedHeights as g, collectFootnoteRefs as h, convertHeaderFooterToContent as i, footnoteReservedHeightsEqual as j, FOOTNOTE_COLUMN_GAP_PX as k, applyFootnotePresentation as l, mapFootnotesToPages as m, calculateHeaderFooterVisualBounds as n, computeHfCaretRectFromView as o, computeHfSelectionRectsFromView as p, contributesToHeaderFooterFlowHeight as q, convertFootnoteToContent as r, stabilizeFootnoteLayout as s, convertHeaderFooterPmDocToContent as t, distributeFootnotesIntoColumns as u, invalidateHfDomCache as v, normalizeHeaderFooterMeasureBlocks as w, resolveHeaderFooterVisualTop as x };
