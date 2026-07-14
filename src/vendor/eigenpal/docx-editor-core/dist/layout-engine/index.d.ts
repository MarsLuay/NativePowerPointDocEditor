/**
 * Layout Engine - Main Entry Point
 *
 * Converts blocks + measures into positioned fragments on pages.
 *
 * @experimental Stable enough for the first-party React adapter, but the
 * API may change in minor releases until a third-party adapter validates
 * it. Pin a version range if you depend on this directly.
 * @packageDocumentation
 * @public
 */
import { Page, PageMargins, ColumnLayout, Fragment, FlowBlock, Measure, SectionBreakBlock, Layout, TextBoxBlock, TableMeasure, LayoutOptions } from './types.js';
export { BlockId, BorderStyle, CellBorderSpec, CellBorders, ColumnBreakBlock, ColumnBreakMeasure, DEFAULT_TEXTBOX_MARGINS, DEFAULT_TEXTBOX_WIDTH, DocumentPosition, FieldRun, FloatingTablePosition, FootnoteContent, FragmentBase, HeaderFooterContentHeights, HeaderFooterLayout, HitTestResult, HyperlinkInfo, ImageBlock, ImageFragment, ImageMeasure, ImageRun, ImageRunPosition, InlineSdtWidget, LineBreakRun, ListNumPr, MeasuredLine, MeasuredLineSegment, PageBreakBlock, PageBreakMeasure, ParagraphAttrs, ParagraphBlock, ParagraphBorders, ParagraphFragment, ParagraphIndent, ParagraphMeasure, ParagraphSpacing, Run, RunFormatting, SdtGroup, SectionBreakMeasure, TabAlignment, TabRun, TabStop, TableBlock, TableCell, TableCellMeasure, TableFragment, TableRow, TableRowMeasure, TextBoxFragment, TextBoxMeasure, TextRun, WrapTextDirection, assertExhaustiveFlowBlock } from './types.js';
export { isFloatingWrapType, isWrapNone, wrapsAroundText } from '../docx/wrapTypes.js';
import '../content-B8ScSBzC.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';

/**
 * Paginator - manages page state during layout
 *
 * Tracks the current page, cursor position, and available space.
 * Creates new pages when content doesn't fit.
 */

/**
 * Current state of a page being laid out.
 */
type PageState = {
    /** The page being built. */
    page: Page;
    /** Current Y position (cursor) from page top. */
    cursorY: number;
    /** Current column index (0-based). */
    columnIndex: number;
    /** Top margin of content area. */
    topMargin: number;
    /** Bottom boundary of content area (page height - bottom margin). */
    contentBottom: number;
    /** Accumulated trailing spacing (space after previous block). */
    trailingSpacing: number;
};
/**
 * Options for creating a paginator.
 */
type PaginatorOptions = {
    /** Page size (width, height). */
    pageSize: {
        w: number;
        h: number;
    };
    /** Page margins. */
    margins: PageMargins;
    /** Column configuration (optional). */
    columns?: ColumnLayout;
    /** Per-page footnote reserved heights (pageNumber → height in pixels). */
    footnoteReservedHeights?: Map<number, number>;
    /** Callback when a new page is created. */
    onNewPage?: (state: PageState) => void;
};
/**
 * Creates a paginator for managing page layout state.
 */
declare function createPaginator(options: PaginatorOptions): {
    /** All pages created so far. */
    pages: Page[];
    /** All page states. */
    states: PageState[];
    /** Column width in pixels (use getColumnWidth() for current value after updates). */
    readonly columnWidth: number;
    /** Get current column layout (returns copy to prevent external mutation). */
    readonly columns: {
        count: number;
        gap: number;
        equalWidth?: boolean;
        separator?: boolean;
    };
    /** Get current state. */
    getCurrentState: () => PageState;
    /** Get available height in current column. */
    getAvailableHeight: () => number;
    /** Get content width for the active section. */
    getContentWidth: () => number;
    /** Check if height fits in current column. */
    fits: (height: number) => boolean;
    /** Ensure height fits, advancing if needed. */
    ensureFits: (height: number) => PageState;
    /** Add a fragment to current page. */
    addFragment: (fragment: Fragment, height: number, spaceBefore?: number, spaceAfter?: number) => {
        state: PageState;
        x: number;
        y: number;
    };
    /** Force a page break. */
    forcePageBreak: () => PageState;
    /** Force a column break. */
    forceColumnBreak: () => PageState;
    /** Get X position for column. */
    getColumnX: (columnIndex: number) => number;
    /** Update column layout (for section breaks). */
    updateColumns: (newColumns: ColumnLayout) => void;
    /** Update page size/margins for subsequent pages. */
    updatePageLayout: (newPageSize?: {
        w: number;
        h: number;
    }, newMargins?: PageMargins, applyImmediately?: boolean) => void;
};
type Paginator = ReturnType<typeof createPaginator>;

/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

/**
 * A chain of consecutive keepNext paragraphs.
 */
type KeepNextChain = {
    /** Index of the first paragraph in the chain. */
    startIndex: number;
    /** Index of the last paragraph in the chain. */
    endIndex: number;
    /** All paragraph indices in the chain. */
    memberIndices: number[];
    /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
    anchorIndex: number;
};
/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A keepNext chain is a sequence of consecutive paragraphs with keepNext=true,
 * followed by an anchor paragraph (the first non-keepNext paragraph).
 * The entire chain must stay on the same page as the anchor's first line.
 *
 * Returns a map from chain start index to chain info.
 */
declare function computeKeepNextChains(blocks: FlowBlock[]): Map<number, KeepNextChain>;
/**
 * Calculate the total height needed to keep a chain together.
 *
 * Includes all chain members plus the first line of the anchor paragraph.
 */
declare function calculateChainHeight(chain: KeepNextChain, blocks: FlowBlock[], measures: Measure[]): number;
/**
 * Get the set of indices that are mid-chain (not chain starters).
 * These should skip the keepNext check since their chain starter already decided.
 */
declare function getMidChainIndices(chains: Map<number, KeepNextChain>): Set<number>;
/**
 * Check if a paragraph has keepLines property (all lines must stay together).
 */
declare function hasKeepLines(block: FlowBlock): boolean;
/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
declare function hasPageBreakBefore(block: FlowBlock): boolean;

/**
 * Section Breaks - Handle page layout changes at section boundaries
 *
 * Sections in DOCX can have different page sizes, margins, columns, and orientations.
 * This module manages the state transitions between sections during layout.
 */

/**
 * State tracking for sections during layout.
 * Uses active/pending pattern to schedule changes at page boundaries.
 */
type SectionState = {
    /** Currently active top margin. */
    activeTopMargin: number;
    /** Currently active bottom margin. */
    activeBottomMargin: number;
    /** Currently active left margin. */
    activeLeftMargin: number;
    /** Currently active right margin. */
    activeRightMargin: number;
    /** Scheduled top margin for next page. */
    pendingTopMargin: number | null;
    /** Scheduled bottom margin for next page. */
    pendingBottomMargin: number | null;
    /** Scheduled left margin for next page. */
    pendingLeftMargin: number | null;
    /** Scheduled right margin for next page. */
    pendingRightMargin: number | null;
    /** Currently active page size. */
    activePageSize: {
        w: number;
        h: number;
    };
    /** Scheduled page size for next page. */
    pendingPageSize: {
        w: number;
        h: number;
    } | null;
    /** Currently active column layout. */
    activeColumns: ColumnLayout;
    /** Scheduled column layout for next page. */
    pendingColumns: ColumnLayout | null;
    /** Currently active orientation. */
    activeOrientation: 'portrait' | 'landscape' | null;
    /** Scheduled orientation for next page. */
    pendingOrientation: 'portrait' | 'landscape' | null;
    /** Whether any pages have been created yet. */
    hasAnyPages: boolean;
};
/**
 * Decision about what happens at a section break.
 */
type BreakDecision = {
    /** Force a page break. */
    forcePageBreak: boolean;
    /** Force a mid-page region change (for column layout changes). */
    forceMidPageRegion: boolean;
    /** Required page parity (even or odd). */
    requiredParity?: 'even' | 'odd';
};
/**
 * Create initial section state from default options.
 */
declare function createInitialSectionState(margins: PageMargins, pageSize: {
    w: number;
    h: number;
}, columns?: ColumnLayout): SectionState;
/**
 * Schedule section break effects by analyzing the break type and updating state.
 *
 * This determines what layout changes should occur (page break, column changes)
 * and schedules the new section properties to be applied at the appropriate boundary.
 */
declare function scheduleSectionBreak(block: SectionBreakBlock, state: SectionState, _baseMargins: PageMargins): {
    decision: BreakDecision;
    state: SectionState;
};
/**
 * Apply pending section state to active state at a page boundary.
 * Transfers all pending values to active and clears pending.
 */
declare function applyPendingToActive(state: SectionState): SectionState;
/**
 * Get the effective margins for the current section state.
 * Returns active margins, or pending if scheduled.
 */
declare function getEffectiveMargins(state: SectionState): PageMargins;
/**
 * Get the effective page size for the current section state.
 */
declare function getEffectivePageSize(state: SectionState): {
    w: number;
    h: number;
};
/**
 * Get the effective columns for the current section state.
 */
declare function getEffectiveColumns(state: SectionState): ColumnLayout;

/**
 * Page index (0-based) whose layout fragments cover `pmPos`, or null if none.
 * Used when the painted DOM may not yet have `[data-pm-start]` for this position (virtualization).
 *
 * Range semantics: `[pmStart, pmEnd)` — half-open, matching ProseMirror's
 * `pos + nodeSize` convention. Boundary positions belong to the next fragment,
 * so when a fragment ends at the same position the next one starts, the next
 * fragment wins (avoids returning the previous page for the start of the
 * next paragraph).
 */
declare function findPageIndexContainingPmPos(layout: Layout, pmPos: number): number | null;

/**
 * Subset of {@link TextBoxBlock} needed to classify how a text box flows
 * relative to surrounding content. Kept narrow so callers (measure, layout,
 * paint) can pass partial views without rebuilding the full block.
 *
 * @public
 */
type TextBoxFlowAttrs = Pick<TextBoxBlock, 'displayMode' | 'wrapType'>;
/**
 * True when a text box participates in float layout — either via the
 * `float` display mode (CSS-style) or via an OOXML `wrapType` that
 * positions the box outside paragraph flow (`square`, `tight`, `through`,
 * `behind`, `inFront`, `topAndBottom`).
 *
 * `topAndBottom` is positioned (it honours its OOXML anchor, e.g. a banner
 * pinned to the top of the page) even though text never wraps beside it —
 * see {@link floatingTextBoxReservesBand}.
 *
 * @public
 */
declare function isFloatingTextBoxBlock(block: TextBoxFlowAttrs): boolean;
/**
 * True when a floating text box reserves a full-width vertical band rather
 * than a side exclusion. `topAndBottom` boxes break the text above and below
 * the box (no text beside it), so surrounding lines flow past the band.
 *
 * @public
 */
declare function floatingTextBoxReservesBand(block: TextBoxFlowAttrs): boolean;
/**
 * True when a floating text box reserves an exclusion zone that narrows
 * surrounding text lines. `wrapNone` (`behind` / `inFront`) and
 * `topAndBottom` are floats that don't wrap text on their sides.
 *
 * @public
 */
declare function floatingTextBoxWrapsText(block: TextBoxFlowAttrs): boolean;

/**
 * Layout Engine - Main Entry Point
 *
 * Converts blocks + measures into positioned fragments on pages.
 *
 * @experimental Stable enough for the first-party React adapter, but the
 * API may change in minor releases until a third-party adapter validates
 * it. Pin a version range if you depend on this directly.
 * @packageDocumentation
 * @public
 */

/**
 * Page-flow geometry resolved from a single section's properties.
 * Exported so the React paged editor can reuse the same shape when
 * measuring blocks per section width — keeping pagination and
 * measurement consistent.
 */
type SectionLayoutConfig = {
    pageSize: {
        w: number;
        h: number;
    };
    margins: PageMargins;
    /** Optional. Sections without explicit columns inherit `{ count: 1 }`. */
    columns?: ColumnLayout;
};
/**
 * Walk `blocks` once and collect per-section geometry. `configs` has one
 * entry per section break plus a trailing `finalConfig`. `breakIndices` is
 * 1-to-1 with the inner break entries (same length as `configs.length - 1`).
 * Callers that need the break `type` can read it from
 * `(blocks[breakIndices[i]] as SectionBreakBlock).type`.
 *
 * @internal
 */
declare function collectSectionConfigs(blocks: FlowBlock[], initialConfig: SectionLayoutConfig, finalConfig: SectionLayoutConfig): {
    configs: SectionLayoutConfig[];
    breakIndices: number[];
};
/**
 * Layout a document: convert blocks + measures into pages with positioned fragments.
 *
 * Algorithm:
 * 1. Walk blocks in order with their corresponding measures
 * 2. For each block, create appropriate fragment(s)
 * 3. Use paginator to manage page/column state
 * 4. Handle page breaks, section breaks, and keepNext chains
 */
declare function layoutDocument(blocks: FlowBlock[], measures: Measure[], options?: LayoutOptions): Layout;
/**
 * Calculate total height of header rows from their measures.
 */
declare function getHeaderRowsHeight(measure: TableMeasure, headerRowCount: number): number;

export { type BreakDecision, ColumnLayout, FlowBlock, Fragment, type KeepNextChain, Layout, LayoutOptions, Measure, Page, PageMargins, type PageState, type Paginator, type PaginatorOptions, SectionBreakBlock, type SectionLayoutConfig, type SectionState, TableMeasure, TextBoxBlock, type TextBoxFlowAttrs, applyPendingToActive, calculateChainHeight, collectSectionConfigs, computeKeepNextChains, createInitialSectionState, createPaginator, findPageIndexContainingPmPos, floatingTextBoxReservesBand, floatingTextBoxWrapsText, getEffectiveColumns, getEffectiveMargins, getEffectivePageSize, getHeaderRowsHeight, getMidChainIndices, hasKeepLines, hasPageBreakBefore, isFloatingTextBoxBlock, layoutDocument, scheduleSectionBreak };
