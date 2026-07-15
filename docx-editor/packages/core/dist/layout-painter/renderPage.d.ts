/**
 * Page Renderer
 *
 * Renders a single page from Layout data to DOM elements.
 * Each page contains positioned fragments within a content area.
 *
 * This file owns the single-page orchestrator (`renderPage`) plus page-level
 * styling (background, borders, content area) and floating-image extraction
 * from paragraphs. Header/footer rendering lives in ./renderPage/headerFooter.ts,
 * footnote area rendering in ./renderPage/footnotes.ts, and the multi-page
 * virtualization / IntersectionObserver layer in ./renderPage/virtualization.ts.
 * @packageDocumentation
 * @public
 */
import { ImageRun, Page, ParagraphBlock, MeasuredLine, TabStop, Run, ParagraphFragment, ParagraphMeasure, ParagraphBorders, Fragment, TableFragment, TableBlock, TableMeasure, ImageFragment, ImageBlock, ImageMeasure, TextBoxFragment, TextBoxBlock, TextBoxMeasure, FlowBlock, Measure, Layout } from '../layout-engine/types.js';
import '../anchoredObjectPosition-CS-8BfhO.js';
import { Node } from 'prosemirror-model';
import { I as ImageAttrs } from '../nodes-Bu9WifHU.js';
import { I as ImageLayoutTarget } from '../ImageExtension-BlI5afZD.js';
import { WrapType } from '../docx/wrapTypes.js';
import { F as FootnoteRenderItem, H as HeaderFooterContent } from '../footnotes-DMsicPGd.js';
export { a as HeaderFooterLayoutInfo, r as resolveHeaderFooterFloatLeft, b as resolveHeaderFooterFloatingTablePosition } from '../footnotes-DMsicPGd.js';
import { B as BorderSpec } from '../colors-C3vA7HUU.js';
import { W as Watermark } from '../watermark-D90356ZM.js';
import { T as Theme } from '../styles-2J4U-Lgk.js';
import '../content-B8ScSBzC.js';
import '../formatting-JhqWT_XM.js';
import '../lists-Bn29SzeS.js';

/**
 * Whether a floating image record reserves space in the text-wrap calculation.
 * Records reaching this predicate have already passed `isFloatingImageRun`, so
 * `wrapType=undefined` implies a `cssFloat`-driven float that wraps text.
 */
declare function floatingImageWrapsText(img: {
    wrapType?: string;
}): boolean;
declare function floatingImageIsBehindDoc(img: {
    wrapType?: string;
}): boolean;
/**
 * Check if an image run is a floating image positioned at page/cell level.
 */
declare function isFloatingImageRun(run: ImageRun): boolean;
/**
 * Check if a floating image should create text wrapping exclusion zones.
 * wrapNone images (`behind` / `inFront`) are positioned floats but do not
 * shrink line widths; text paints over or under them.
 */
declare function isTextWrappingFloatingImageRun(run: ImageRun): boolean;

/**
 * Minimum fields the floating-image painter needs. Page-level and cell-level
 * float records both satisfy this shape.
 */
interface FloatingImagePaintRecord {
    src: string;
    width: number;
    height: number;
    alt?: string;
    transform?: string;
    x: number;
    y: number;
    pmStart?: number;
    pmEnd?: number;
    /** wp:srcRect crop fractions in [0, 1]. */
    cropTop?: number;
    cropRight?: number;
    cropBottom?: number;
    cropLeft?: number;
    /** a:alphaModFix -> CSS opacity. */
    opacity?: number;
}
interface FloatingImagesLayerOptions {
    layerClass: string;
    itemClass: string;
    /**
     * `inset0` sizes the layer with `top/right/bottom/left = 0` (used at page level).
     * `fullSize` uses `width/height = 100%` and adds `overflow: hidden` (used inside table cells).
     */
    sizing: 'inset0' | 'fullSize';
    /** `behind` skips z-index so DOM order keeps the layer below body fragments. */
    layerMode: 'front' | 'behind';
}
/**
 * Render a layer of positioned floating images. Used at both page level and
 * inside table cells; the variant differs only in class names and sizing.
 */
declare function renderFloatingImagesLayer(floatingImages: FloatingImagePaintRecord[], doc: Document, options: FloatingImagesLayerOptions): HTMLElement;

/**
 * Multi-page rendering with virtualization.
 *
 * For documents under VIRTUALIZATION_THRESHOLD pages, all pages render
 * eagerly. Larger documents render only pages near the viewport — off-screen
 * pages are lightweight shells (correct dimensions, no fragment content) so
 * scroll position is preserved. An IntersectionObserver populates and clears
 * page content as the user scrolls. Incremental updates (re-rendering only
 * fingerprint-changed pages) avoid blink when the document model shifts.
 */

type RenderPagesUpdateKind = 'incremental' | 'full';
/**
 * Render multiple pages to a container with virtualization for large documents.
 *
 * For documents with fewer than VIRTUALIZATION_THRESHOLD pages, all pages
 * are rendered eagerly. For larger documents, only pages near the visible
 * viewport are fully rendered — off-screen pages are lightweight shells
 * with correct dimensions to preserve scroll position.
 *
 * An IntersectionObserver watches page elements and populates/clears
 * content as pages scroll into and out of view.
 */
declare function renderPages(pages: Page[], container: HTMLElement, options?: RenderPageOptions & {
    pageGap?: number;
    footnotesByPage?: Map<number, FootnoteRenderItem[]>;
}): RenderPagesUpdateKind;
/**
 * Force every virtualized page shell in `container` to be fully rendered.
 *
 * Virtualization keeps off-screen pages as empty shells so cloning the
 * pages container for print (or any DOM snapshot) yields blank pages past
 * the visible band. Callers that need every page populated — print,
 * export-to-HTML, pdf snapshot — should call this first.
 *
 * No-op for small documents (rendered eagerly) or containers that were
 * never managed by `renderPages`. Returns the number of shells populated
 * by this call (useful for tests).
 */
declare function renderAllPagesNow(container: HTMLElement): number;

/**
 * Page Renderer
 *
 * Renders a single page from Layout data to DOM elements.
 * Each page contains positioned fragments within a content area.
 *
 * This file owns the single-page orchestrator (`renderPage`) plus page-level
 * styling (background, borders, content area) and floating-image extraction
 * from paragraphs. Header/footer rendering lives in ./renderPage/headerFooter.ts,
 * footnote area rendering in ./renderPage/footnotes.ts, and the multi-page
 * virtualization / IntersectionObserver layer in ./renderPage/virtualization.ts.
 * @packageDocumentation
 * @public
 */

/**
 * CSS class names for page elements
 */
declare const PAGE_CLASS_NAMES: {
    page: string;
    content: string;
    header: string;
    footer: string;
};
/**
 * Context passed to fragment renderers
 */
interface RenderContext {
    /** Current page number (1-indexed) */
    pageNumber: number;
    /** Total number of pages */
    totalPages: number;
    /** Which section is being rendered */
    section: 'body' | 'header' | 'footer';
    /** Content width in pixels (page width minus margins) - used for justify */
    contentWidth?: number;
    /** When true, floating images render in-flow instead of being skipped (for table cells) */
    insideTableCell?: boolean;
    /** Comment IDs that are resolved — skip highlight for these */
    resolvedCommentIds?: Set<number>;
    /**
     * How the renderer should position its outer element. The body lays
     * fragments at absolute (x, y) on the page (`'absolute'`, the default),
     * while headers/footers and text boxes flow blocks vertically and let
     * normal document flow handle placement (`'flow'`). The caller passes
     * 'flow' instead of overwriting the renderer's inline styles after the
     * fact (#379).
     */
    positioning?: 'absolute' | 'flow';
}
/**
 * Options for rendering a page
 */
interface RenderPageOptions {
    /** Document to create elements in (default: window.document) */
    ['document']?: Document;
    /** Custom page class name */
    pageClassName?: string;
    /** Show page borders (for debugging) */
    showBorders?: boolean;
    /** Background color for pages */
    backgroundColor?: string;
    /** Drop shadow on pages */
    showShadow?: boolean;
    /** Header content to render (used for all pages, or pages 2+ when titlePg is set). */
    headerContent?: HeaderFooterContent;
    /** Footer content to render (used for all pages, or pages 2+ when titlePg is set). */
    footerContent?: HeaderFooterContent;
    /** Header content for the first page only (when titlePg is set). */
    firstPageHeaderContent?: HeaderFooterContent;
    /** Footer content for the first page only (when titlePg is set). */
    firstPageFooterContent?: HeaderFooterContent;
    /** Whether different first page headers/footers are enabled (w:titlePg). */
    titlePg?: boolean;
    /** Distance from page top to header content. */
    headerDistance?: number;
    /** Distance from page bottom to footer content. */
    footerDistance?: number;
    /** Block lookup for rendering actual content. */
    blockLookup?: BlockLookup;
    /** OOXML page borders from section properties. */
    pageBorders?: {
        top?: BorderSpec;
        bottom?: BorderSpec;
        left?: BorderSpec;
        right?: BorderSpec;
        display?: 'allPages' | 'firstPage' | 'notFirstPage';
        offsetFrom?: 'page' | 'text';
        zOrder?: 'front' | 'back';
    };
    /** Theme for resolving border colors. */
    theme?: Theme | null;
    /** Footnotes to render at the bottom of this page. */
    footnoteArea?: FootnoteRenderItem[];
    /** Comment IDs that are resolved — skip highlight for these */
    resolvedCommentIds?: Set<number>;
    /** Watermark to paint behind body content (resolved from the page's section header). */
    watermark?: Watermark;
}
/**
 * Apply page styles to an element. Exported because virtualization.ts uses it
 * to size lightweight shells before content lands in them.
 */
declare function applyPageStyles(element: HTMLElement, width: number, height: number, options: RenderPageOptions): void;
/**
 * Render a single page to DOM
 *
 * @param page - The page to render
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The page DOM element
 */
declare function renderPage(page: Page, context: RenderContext, options?: RenderPageOptions): HTMLElement;

/**
 * Line-level rendering.
 *
 * Owns `renderLine` and its helpers: slicing the paragraph's runs to the
 * line's character range, justify decisions, per-line floating margins,
 * tab-width calculation through the tabCalculator (explicit stops + default
 * intervals), inline image dedup, and field-value substitution width math.
 */

/**
 * Slice runs for a specific line
 *
 * @param block - The paragraph block
 * @param line - The line measurement
 * @returns Array of runs for this line
 */
declare function sliceRunsForLine(block: ParagraphBlock, line: MeasuredLine): Run[];
/**
 * Options for rendering a line with justify support
 */
interface RenderLineOptions {
    /** Available width for the line (content area width minus indentation) */
    availableWidth: number;
    /** Whether this is the last line of the paragraph */
    isLastLine: boolean;
    /** Whether this is the first line of the paragraph */
    isFirstLine: boolean;
    /** Whether the paragraph ends with a line break */
    paragraphEndsWithLineBreak: boolean;
    /** Tab stops from paragraph attributes */
    tabStops?: TabStop[];
    /** Render context for field substitution */
    context?: RenderContext;
    /** Left indent in pixels */
    leftIndentPx?: number;
    /** First line indent in pixels (positive) or hanging indent (negative) */
    firstLineIndentPx?: number;
    /** Line-specific floating image margins (calculated per-line based on Y overlap) */
    floatingMargins?: {
        leftMargin: number;
        rightMargin: number;
    };
    /** Track inline image runs already rendered in this paragraph fragment to prevent duplicates */
    renderedInlineImageKeys?: Set<string>;
    /**
     * Rightmost x where inline content may render, in content-area coords. Used
     * by the right-tab anchor; passed in directly (rather than recomposed from
     * `leftIndentPx + availableWidth`) because `availableWidth` excludes the
     * hung-out region for some inputs and would drift.
     */
    lineRightEdgePx?: number;
}
/**
 * Render a single line
 *
 * @param block - The paragraph block
 * @param line - The line measurement
 * @param alignment - Text alignment
 * @param doc - Document to create elements in
 * @param options - Additional options for justify calculation
 * @returns The line DOM element
 */
declare function renderLine(block: ParagraphBlock, line: MeasuredLine, alignment: 'left' | 'center' | 'right' | 'justify' | undefined, doc: Document, options?: RenderLineOptions): HTMLElement;

/**
 * Paragraph Fragment Renderer
 *
 * Renders paragraph fragments with lines and text runs to DOM.
 * Handles text formatting, alignment, and positioning.
 *
 * This file owns `renderParagraphFragment` (the orchestrator), the
 * border-grouping helpers, and the list-marker renderer. Per-run rendering
 * (text/tab/image/break/field) lives in ./renderParagraph/runs.ts and the
 * line-level walker is in ./renderParagraph/line.ts. The shared class-name
 * constants and run-type guards are in ./renderParagraph/shared.ts.
 */

/**
 * Options for rendering a paragraph
 */
interface RenderParagraphOptions {
    /** Document to create elements in */
    ['document']?: Document;
    /** Fragment's Y position relative to content area (for per-line margin calculation) */
    fragmentContentY?: number;
    /** Borders from the previous adjacent paragraph (for border grouping) */
    prevBorders?: ParagraphBorders;
    /** Borders from the next adjacent paragraph (for border grouping) */
    nextBorders?: ParagraphBorders;
    /** Inline image runs already rendered for this paragraph block */
    renderedInlineImageKeys?: Set<string>;
}
/**
 * Render a paragraph fragment
 *
 * @param fragment - The fragment to render
 * @param block - The paragraph block
 * @param measure - The paragraph measurement
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The fragment DOM element
 */
declare function renderParagraphFragment(fragment: ParagraphFragment, block: ParagraphBlock, measure: ParagraphMeasure, context: RenderContext, options?: RenderParagraphOptions): HTMLElement;

/**
 * Fragment Renderer
 *
 * Renders individual fragments (paragraphs, tables, images) to DOM.
 * Each fragment is positioned within a page's content area.
 */

/**
 * CSS class names for fragment elements
 */
declare const FRAGMENT_CLASS_NAMES: {
    fragment: string;
    paragraph: string;
    table: string;
    image: string;
    line: string;
    run: string;
};
/**
 * Options for rendering fragments
 */
interface RenderFragmentOptions {
    /** Document to create elements in */
    ['document']?: Document;
}
/**
 * Render a fragment to DOM
 *
 * @param fragment - The fragment to render
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The fragment DOM element
 */
declare function renderFragment(fragment: Fragment, context: RenderContext, options?: RenderFragmentOptions): HTMLElement;

/**
 * Table Renderer
 *
 * Renders table fragments to DOM. Handles:
 * - Multi-row tables split across pages
 * - Cell content (paragraphs within cells)
 * - Column widths and cell spans
 * - Basic cell styling (borders, backgrounds)
 */

/**
 * CSS class names for table elements
 */
declare const TABLE_CLASS_NAMES: {
    table: string;
    row: string;
    cell: string;
    cellContent: string;
    resizeHandle: string;
    rowResizeHandle: string;
    tableEdgeHandleBottom: string;
    tableEdgeHandleRight: string;
};
/**
 * Options for rendering a table fragment
 */
interface RenderTableFragmentOptions {
    ['document']?: Document;
}
/**
 * Render a table fragment to DOM
 *
 * @param fragment - The table fragment to render
 * @param block - The full table block
 * @param measure - The full table measure
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The table DOM element
 */
declare function renderTableFragment(fragment: TableFragment, block: TableBlock, measure: TableMeasure, context: RenderContext, options?: RenderTableFragmentOptions): HTMLElement;

/**
 * Image Renderer
 *
 * Renders image fragments to DOM. Handles:
 * - Inline images
 * - Anchored/floating images with z-index layering
 * - Basic image sizing
 */

/**
 * CSS class names for image elements
 */
declare const IMAGE_CLASS_NAMES: {
    image: string;
    imageAnchored: string;
};
/**
 * Options for rendering an image fragment
 */
interface RenderImageFragmentOptions {
    ['document']?: Document;
}
/**
 * Render an image fragment to DOM
 *
 * @param fragment - The image fragment to render
 * @param block - The full image block
 * @param measure - The image measure
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The image DOM element
 */
declare function renderImageFragment(fragment: ImageFragment, block: ImageBlock, _measure: ImageMeasure, _context: RenderContext, options?: RenderImageFragmentOptions): HTMLElement;

/**
 * Text Box Renderer
 *
 * Renders text box fragments to DOM. Handles:
 * - Background fill color
 * - Border/outline
 * - Internal padding (margins)
 * - Paragraph content inside the box (using pre-measured data)
 */

/**
 * CSS class names for text box elements
 */
declare const TEXTBOX_CLASS_NAMES: {
    textBox: string;
};
/**
 * Options for rendering a text box fragment
 */
interface RenderTextBoxFragmentOptions {
    ['document']?: Document;
}
/**
 * Render a text box fragment to DOM
 */
declare function renderTextBoxFragment(fragment: TextBoxFragment, block: TextBoxBlock, measure: TextBoxMeasure, context: RenderContext, options?: RenderTextBoxFragmentOptions): HTMLElement;

/**
 * Visible boundary chrome for block-level content controls (`w:sdt`) in the
 * paged view. Kept out of `renderPage.ts` so that file stays under its
 * line cap; the body painter calls `renderSdtBoundaryBoxes` once per page.
 */

/**
 * Group ids of every block-level content control that encloses the given
 * selection — used to keep a control's boundary visible while the caret is
 * inside it (Word-style focus), independent of mouse hover. The id matches
 * `toFlowBlocks`' `sdt@${pos}` scheme, where `pos` is the SDT node's position;
 * `$pos.before(d)` yields exactly that for an ancestor at depth `d`, so the
 * ids line up with the painted boxes' `data-sdt-group-id`. Both ends of a
 * range are collected so a selection straddling a control still lights it up.
 */
declare function enclosingSdtGroupIds(doc: Node, from: number, to: number): Set<string>;
/**
 * Toggle the `.is-focused` reveal class on the painted boundary boxes whose
 * control encloses the caret. Kept separate from the hover-driven `.is-active`
 * class so the two reveal paths never clear each other.
 */
declare function applySdtFocus(container: HTMLElement, focusedIds: Set<string>): void;

/**
 * Image layout helpers shared between framework adapters (React, Vue, ...).
 *
 * Everything here is framework-agnostic: pure DOM math + pure functions over
 * OOXML wrap-type vocabulary. The corresponding UI bindings (right-click menu,
 * toolbar dropdown) live in each framework adapter and call into these.
 */

declare const LAYOUT_IMAGE_CLASSES: {
    /** Inline image rendered inside `.layout-line`. */
    readonly runImage: "layout-run-image";
    /** Block (centered, topAndBottom) image. */
    readonly blockImage: "layout-block-image";
    /** Anchored image rendered in the page-level floating layer. */
    readonly pageFloatingImage: "layout-page-floating-image";
    /** Anchored image rendered inside a table cell's floating layer. */
    readonly cellFloatingImage: "layout-cell-floating-image";
    readonly pageContent: "layout-page-content";
    readonly paragraph: "layout-paragraph";
};
interface ImageHitTestResult {
    /** PM doc position of the image node, read from `data-pm-start`. */
    pos: number;
    /** The matched element — pass to `captureInlinePositionEmu` if it's inline. */
    imageEl: HTMLElement;
}
/**
 * Walk up from an event target looking for any rendered image element. Returns
 * the PM position embedded in `data-pm-start`, or null if the target isn't on
 * an image.
 */
declare function hitTestImage(target: EventTarget | null): ImageHitTestResult | null;
/**
 * Walk up from a click target to the nearest rendered image element, returning
 * just the element (no PM position parsing). Used by the left-click selection
 * path in both adapters — `data-pm-start` is read separately by callers that
 * need the position. Returns the element when the click was directly on an
 * inline `<img.layout-run-image>` OR inside one of the container classes
 * registered in `LAYOUT_IMAGE_CLASSES`.
 */
declare function findImageElement(target: EventTarget | null): HTMLElement | null;
/**
 * Capture the rendered position of an inline image as EMUs, normalised to
 * unzoomed coordinates. Returns horizontal offset relative to the page
 * content area (column origin) and vertical offset relative to the
 * containing paragraph — matches the OOXML attrs the resolver writes
 * (`relativeFrom: 'column'` / `relativeFrom: 'paragraph'`).
 *
 * `zoom` defaults to 1; pass the editor's current zoom factor when the
 * pages container has a CSS scale applied so `getBoundingClientRect` deltas
 * are converted back to authored-pixel space before going to EMU.
 *
 * Returns undefined for non-inline images or detached DOM.
 */
declare function captureInlinePositionEmu(imageEl: HTMLElement, zoom?: number): {
    horizontalEmu: number;
    verticalEmu: number;
} | undefined;
/**
 * Map an image's current OOXML attrs onto the menu's choice vocabulary so the
 * menu can highlight the active option. Returns null for `topAndBottom` and
 * any unknown wrap type — those don't have a directional menu entry.
 *
 * `cssFloat` accepts both `null` and `undefined` so framework adapters can
 * use either as their "unset" sentinel without an extra normalisation step.
 */
declare function deriveLayoutChoice(wrapType: WrapType, cssFloat?: ImageAttrs['cssFloat'] | null): ImageLayoutTarget | null;
/**
 * Hint to the framework's icon registry for which Material Symbol — or
 * equivalent — to render alongside each option. Bindings own the icon
 * component itself.
 */
type ImageLayoutIconHint = 'inline' | 'squareLeft' | 'squareRight' | 'behind' | 'inFront';
interface ImageLayoutOptionDef {
    /** Choice value — what gets dispatched on click. */
    choice: ImageLayoutTarget;
    /** i18n key under `imageWrap.menu.*`. */
    i18nLabelKey: string;
    /** i18n key under `imageWrap.menuDesc.*`. */
    i18nDescKey: string;
    /** Hint for the framework's icon registry. */
    iconHint: ImageLayoutIconHint;
}
/** Mirrors Word's Wrap Text menu — five directional options. */
declare const IMAGE_LAYOUT_OPTIONS: readonly ImageLayoutOptionDef[];
/**
 * Whether a given option is enabled for an image with the given current wrap
 * type. Every option stays clickable — picking the option that matches the
 * current state is a no-op (the PM command early-returns), which matches
 * Word's behavior. We don't grey out the current option, so the menu reads
 * consistently regardless of whether the image is inline or anchored.
 *
 * The flag is kept around for forward-compatibility (e.g. future read-only
 * mode), but currently always returns true.
 */
declare function isImageLayoutOptionEnabled(_option: ImageLayoutOptionDef, _currentWrapType: WrapType): boolean;
/**
 * Translate the legacy toolbar wrap-type vocabulary (`wrapLeft` / `wrapRight`
 * / `square` / `tight` / `through` / `topAndBottom` / `behind` / `inFront` /
 * `inline`) into a `ImageLayoutTarget` so toolbar dispatch shares the same PM
 * command path as the right-click menu.
 *
 * Returns `undefined` for unknown values; callers should treat that as
 * "no-op".
 */
declare function toolbarValueToLayoutTarget(value: string): ImageLayoutTarget | undefined;

/**
 * Layout Painter
 *
 * Main entry point for rendering Layout data to DOM.
 * Provides reconciliation for efficient incremental updates.
 *
 * @experimental Stable enough for the first-party React adapter, but the
 * API may change in minor releases until a third-party adapter validates
 * it. Pin a version range if you depend on this directly.
 * @packageDocumentation
 * @public
 */

/**
 * Block lookup entry for painter
 */
interface BlockLookupEntry {
    block: FlowBlock;
    measure: Measure;
    version?: string;
}
/**
 * Block lookup map type
 */
type BlockLookup = Map<string, BlockLookupEntry>;
/**
 * Build the painter's `block.id → { block, measure }` lookup from the parallel
 * blocks/measures arrays. Shared by both adapters' paint step.
 */
declare function buildBlockLookup(blocks: FlowBlock[], measures: Measure[]): BlockLookup;
/**
 * Painter options
 */
interface PainterOptions {
    /** Document to create elements in */
    ['document']?: Document;
    /** Gap between pages in pixels */
    pageGap?: number;
    /** Show page shadows */
    showShadow?: boolean;
    /** Background color for pages */
    pageBackground?: string;
    /** Container background color */
    containerBackground?: string;
}
/**
 * Layout Painter class
 *
 * Renders Layout data to DOM with efficient reconciliation.
 * Only updates changed pages and fragments for better performance.
 */
declare class LayoutPainter {
    private container;
    private blockLookup;
    private pageStates;
    private totalPages;
    private options;
    private doc;
    resolvedCommentIds: Set<number>;
    constructor(options?: PainterOptions);
    /**
     * Set the block lookup map for rendering fragments
     */
    setBlockLookup(lookup: BlockLookup): void;
    /**
     * Mount the painter to a container element
     */
    mount(container: HTMLElement): void;
    /**
     * Unmount the painter
     */
    unmount(): void;
    /**
     * Apply styles to the container
     */
    private applyContainerStyles;
    /**
     * Paint a layout to the container
     */
    paint(layout: Layout): void;
    /**
     * Render a page using block lookup for full fragment rendering
     */
    private renderPageWithLookup;
    /**
     * Render a fragment using block lookup for full content rendering
     */
    private renderFragmentWithLookup;
    /**
     * Apply positioning styles to a fragment element
     */
    private applyFragmentPosition;
    /**
     * Get the current page count
     */
    getPageCount(): number;
    /**
     * Get a page element by index
     */
    getPageElement(index: number): HTMLElement | null;
    /**
     * Scroll to a specific page
     */
    scrollToPage(pageNumber: number): void;
}
/**
 * Create a new LayoutPainter instance
 */
declare function createPainter(options?: PainterOptions): LayoutPainter;

export { type BlockLookup as B, FRAGMENT_CLASS_NAMES as F, type FloatingImagePaintRecord, type FloatingImagesLayerOptions, FootnoteRenderItem, HeaderFooterContent, IMAGE_CLASS_NAMES as I, LAYOUT_IMAGE_CLASSES as L, type PainterOptions as P, PAGE_CLASS_NAMES, type RenderContext, type RenderPageOptions, type RenderPagesUpdateKind, TABLE_CLASS_NAMES as T, type BlockLookupEntry as a, applyPageStyles, IMAGE_LAYOUT_OPTIONS as b, type ImageHitTestResult as c, type ImageLayoutIconHint as d, type ImageLayoutOptionDef as e, LayoutPainter as f, floatingImageIsBehindDoc, floatingImageWrapsText, TEXTBOX_CLASS_NAMES as g, applySdtFocus as h, buildBlockLookup as i, isFloatingImageRun, isTextWrappingFloatingImageRun, captureInlinePositionEmu as j, createPainter as k, deriveLayoutChoice as l, enclosingSdtGroupIds as m, findImageElement as n, hitTestImage as o, isImageLayoutOptionEnabled as p, renderImageFragment as q, renderFragment as r, renderAllPagesNow, renderFloatingImagesLayer, renderPage, renderPages, renderLine as s, renderParagraphFragment as t, renderTableFragment as u, renderTextBoxFragment as v, sliceRunsForLine as w, toolbarValueToLayoutTarget as x };
