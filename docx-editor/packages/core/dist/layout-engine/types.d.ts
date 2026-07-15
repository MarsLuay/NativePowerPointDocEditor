/**
 * Layout Engine Types
 *
 * Core types for the paginated layout engine.
 * Converts document blocks + measurements into positioned fragments on pages.
 * @packageDocumentation
 * @public
 */
import { a as RevisionInfo, ax as CellMarker } from '../content-B8ScSBzC.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../docx/wrapTypes.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';

/**
 * Inline SDT widget metadata carried by a painted text run.
 */
interface InlineSdtWidget {
    kind: 'checkbox';
    /** Stable per-document id derived from the ProseMirror node position. */
    groupId: string;
    /** ProseMirror position of the inline SDT node. */
    pos: number;
    /** Word tag value (`w:tag`). */
    tag?: string;
    /** Word alias value (`w:alias`). */
    alias?: string;
    /** Live checkbox glyph state. */
    checked?: boolean;
}

/**
 * Unique identifier for a block in the document.
 * Format: typically `${index}-${type}` or just the block index.
 */
type BlockId = string | number;
/**
 * Common run formatting properties applied to text runs.
 */
type RunFormatting = {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | {
        style?: string;
        color?: string;
    };
    strike?: boolean;
    color?: string;
    highlight?: string;
    fontFamily?: string;
    fontSize?: number;
    letterSpacing?: number;
    superscript?: boolean;
    subscript?: boolean;
    /** Render glyphs as uppercase regardless of source case (OOXML w:caps). */
    allCaps?: boolean;
    /** Render lowercase glyphs as small uppercase (OOXML w:smallCaps). */
    smallCaps?: boolean;
    /**
     * Vertical baseline shift in CSS pixels (positive = up). OOXML w:position is
     * authored in half-points; converted to px during the bridge so the painter
     * can apply it directly via vertical-align without re-doing the math.
     */
    positionPx?: number;
    /**
     * Horizontal text scale as a percentage (100 = normal, 50 = half-width,
     * 200 = double-width). OOXML w:w specifies pct in raw % (e.g. 90 means 90%).
     */
    horizontalScale?: number;
    /**
     * Minimum font size in points at which kerning kicks in (OOXML w:kern, half-
     * points in source). When this run's effective font size is at or above this
     * threshold, the painter enables font-kerning: normal.
     */
    kerningMinPt?: number;
    /** Engraved/imprint effect (OOXML w:imprint, §17.3.2.18). */
    imprint?: boolean;
    /** Embossed/raised effect (OOXML w:emboss, §17.3.2.13). */
    emboss?: boolean;
    /** Drop-shadow effect (OOXML w:shadow, §17.3.2.31). */
    textShadow?: boolean;
    /** Outlined / hollow text (OOXML w:outline, §17.3.2.23). */
    textOutline?: boolean;
    /**
     * CJK emphasis mark (OOXML w:em, §17.3.2.12). Maps to CSS `text-emphasis`
     * with a position and style; the painter handles the variant lookup.
     */
    emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot';
    /** Hidden run (OOXML w:vanish, §17.3.2.41). Painter skips the run. */
    hidden?: boolean;
    /**
     * Per-run right-to-left direction (OOXML w:rtl, §17.3.2.30). Independent
     * from the paragraph's bidi flag — a single run may flip direction within
     * an LTR paragraph.
     */
    rtl?: boolean;
    /**
     * Legacy text-effect animation (OOXML w:effect, §17.3.2.11). The painter
     * surfaces it as a class hook so host CSS can opt in to animations.
     */
    textEffect?: 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle';
    /** Hyperlink info if this run is a link */
    hyperlink?: HyperlinkInfo;
    /** Footnote reference ID (if this run contains a footnote reference) */
    footnoteRefId?: number;
    /** Endnote reference ID (if this run contains an endnote reference) */
    endnoteRefId?: number;
    /** Comment IDs if this run is within a comment range */
    commentIds?: number[];
    /** Whether this run is a tracked insertion */
    isInsertion?: boolean;
    /** Whether this run is a tracked deletion */
    isDeletion?: boolean;
    /** Author of the tracked change */
    changeAuthor?: string;
    /** Date of the tracked change */
    changeDate?: string;
    /** Revision ID of the tracked change (for sidebar matching) */
    changeRevisionId?: number;
};
/**
 * Hyperlink information for a run.
 */
type HyperlinkInfo = {
    href: string;
    tooltip?: string;
    /**
     * Skip the painter's "Word default" blue + underline fallback when the run
     * has no resolved color/underline. Set by the layout-bridge for hyperlinks
     * inside TOC paragraphs, where Word renders entries in the TOCx paragraph
     * color rather than the Hyperlink character style's color.
     */
    noDefaultStyle?: boolean;
};
/**
 * A text run within a paragraph.
 */
type TextRun = RunFormatting & {
    kind: 'text';
    text: string;
    /** Hyperlink information if this run is a link. */
    hyperlink?: HyperlinkInfo;
    /** Absolute ProseMirror position (inclusive) of first character. */
    pmStart?: number;
    /** Absolute ProseMirror position (exclusive) after last character. */
    pmEnd?: number;
    /** Inline content-control widget metadata when this run is the visible glyph. */
    inlineSdtWidget?: InlineSdtWidget;
};
/**
 * A tab character run.
 */
type TabRun = RunFormatting & {
    kind: 'tab';
    width?: number;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * Position data for floating/anchored images.
 */
type ImageRunPosition = {
    horizontal?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
    };
    vertical?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
    };
};
type WrapTextDirection = 'bothSides' | 'left' | 'right' | 'largest';
/**
 * An inline image run.
 */
type ImageRun = {
    kind: 'image';
    src: string;
    width: number;
    height: number;
    alt?: string;
    /** CSS transform string (rotation, flip) */
    transform?: string;
    /** Position for floating/anchored images */
    position?: ImageRunPosition;
    /** Wrap type from DOCX (inline, square, tight, through, topAndBottom, etc.) */
    wrapType?: string;
    /** Display mode for CSS rendering */
    displayMode?: 'inline' | 'block' | 'float';
    /** CSS float direction */
    cssFloat?: 'left' | 'right' | 'none';
    /** Wrap distances in pixels */
    distTop?: number;
    distBottom?: number;
    distLeft?: number;
    distRight?: number;
    /** wp:srcRect crop fractions in [0, 1]; emit as CSS clip-path inset. */
    cropTop?: number;
    cropRight?: number;
    cropBottom?: number;
    cropLeft?: number;
    /** a:alphaModFix → CSS opacity in [0, 1]. */
    opacity?: number;
    /** Whether this picture is itself a tracked insertion (`<w:ins>`). */
    isInsertion?: boolean;
    /** Whether this picture is itself a tracked deletion (`<w:del>`). */
    isDeletion?: boolean;
    /** Author of the tracked change wrapping the picture. */
    changeAuthor?: string;
    /** Date of the tracked change wrapping the picture. */
    changeDate?: string;
    /** Revision id of the tracked change (for sidebar matching). */
    changeRevisionId?: number;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * A line break run.
 */
type LineBreakRun = {
    kind: 'lineBreak';
    pmStart?: number;
    pmEnd?: number;
};
/**
 * A field run (PAGE, NUMPAGES, etc.) that gets substituted at render time.
 */
type FieldRun = RunFormatting & {
    kind: 'field';
    fieldType: 'PAGE' | 'NUMPAGES' | 'DATE' | 'TIME' | 'OTHER';
    /** Fallback text if field can't be resolved */
    fallback?: string;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * Union of all run types.
 */
type Run = TextRun | TabRun | ImageRun | LineBreakRun | FieldRun;
/**
 * Paragraph spacing configuration.
 */
type ParagraphSpacing = {
    before?: number;
    after?: number;
    line?: number;
    lineUnit?: 'px' | 'multiplier';
    lineRule?: 'auto' | 'exact' | 'atLeast';
};
/**
 * Paragraph indentation configuration.
 */
type ParagraphIndent = {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
};
/**
 * Tab stop alignment types
 */
type TabAlignment = 'start' | 'end' | 'center' | 'decimal' | 'bar' | 'clear';
/**
 * Tab stop definition
 */
type TabStop = {
    /** Tab alignment mode */
    val: TabAlignment;
    /** Position in twips from left margin */
    pos: number;
    /** Optional leader character */
    leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
};
/**
 * Border specification for paragraphs.
 */
type BorderStyle = {
    style?: string;
    width?: number;
    color?: string;
    space?: number;
};
/**
 * Paragraph borders.
 */
type ParagraphBorders = {
    top?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    right?: BorderStyle;
    between?: BorderStyle;
    bar?: BorderStyle;
};
/**
 * List numbering properties for a paragraph.
 */
type ListNumPr = {
    numId?: number;
    ilvl?: number;
};
/**
 * Paragraph block attributes.
 */
type ParagraphAttrs = {
    alignment?: 'left' | 'center' | 'right' | 'justify';
    spacing?: ParagraphSpacing;
    /** See ParagraphFormatting.spacingExplicit. */
    spacingExplicit?: {
        before?: boolean;
        after?: boolean;
    };
    indent?: ParagraphIndent;
    keepNext?: boolean;
    keepLines?: boolean;
    pageBreakBefore?: boolean;
    styleId?: string;
    contextualSpacing?: boolean;
    /** Right-to-left paragraph direction */
    bidi?: boolean;
    borders?: ParagraphBorders;
    shading?: string;
    tabs?: TabStop[];
    numPr?: ListNumPr;
    listMarker?: string;
    listIsBullet?: boolean;
    listMarkerHidden?: boolean;
    listMarkerFontFamily?: string;
    listMarkerFontSize?: number;
    listMarkerSuffix?: 'tab' | 'space' | 'nothing';
    /**
     * Tracked-change state of the list numbering itself. When a list is applied
     * (or removed) under suggesting mode, the marker paints in the insertion /
     * deletion color so an inserted list item's number reads as part of the
     * suggestion, matching Word. `undefined` = numbering is not a pending change.
     */
    listMarkerRevision?: 'ins' | 'del';
    /** Document-wide `w:defaultTabStop` in twips (§17.6.13). Default 720. */
    defaultTabStopTwips?: number;
    defaultFontSize?: number;
    defaultFontFamily?: string;
    /**
     * Skip the empty-paragraph line-height fallback. Used by HF measurement
     * for the canonical OOXML "trailing empty paragraph after a table" pattern
     * — Word renders that paragraph as a zero-height anchor, not as a full
     * line-height of phantom space. See `normalizeHeaderFooterMeasureBlocks`
     * (#381).
     */
    suppressEmptyParagraphHeight?: boolean;
    /**
     * Tracked-change marker on the paragraph mark itself
     * (`<w:pPr><w:rPr><w:ins/>`). The painter renders a pilcrow and margin
     * change bar; the cache key includes presence so two paragraphs with
     * different revisions don't share a measurement.
     */
    pPrIns?: RevisionInfo | null;
    /** Tracked-change marker on the paragraph mark (`<w:pPr><w:rPr><w:del/>`). */
    pPrDel?: RevisionInfo | null;
};
/**
 * A paragraph block containing runs.
 */
type ParagraphBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'paragraph';
    id: BlockId;
    /** Stable Word `w14:paraId` / PM `paraId`, when available. */
    paraId?: string;
    runs: Run[];
    attrs?: ParagraphAttrs;
    /** ProseMirror start position for this block. */
    pmStart?: number;
    /** ProseMirror end position for this block. */
    pmEnd?: number;
};
/**
 * Cell border specification for rendering.
 */
type CellBorderSpec = {
    width?: number;
    color?: string;
    style?: string;
};
/**
 * Cell borders (all four sides).
 */
type CellBorders = {
    top?: CellBorderSpec;
    bottom?: CellBorderSpec;
    left?: CellBorderSpec;
    right?: CellBorderSpec;
};
/**
 * A table cell with content.
 */
type TableCell = {
    id: BlockId;
    blocks: FlowBlock[];
    colSpan?: number;
    rowSpan?: number;
    width?: number;
    /** Original DOCX cell width value, before unit conversion. */
    widthValue?: number;
    /** Original DOCX cell width type ('auto', 'pct', 'dxa', 'nil'). */
    widthType?: string;
    verticalAlign?: 'top' | 'center' | 'bottom';
    background?: string;
    borders?: CellBorders;
    /** Per-cell padding in pixels (from w:tcMar or table-level w:tblCellMar) */
    padding?: {
        top: number;
        right: number;
        bottom: number;
        left: number;
    };
    /**
     * `w:noWrap`: when true, the cell forbids text wrapping inside it. The
     * painter renders this as `white-space: nowrap` on the content container
     * — content stays on one line and the cell expands horizontally.
     */
    noWrap?: boolean;
    /** Tracked cell marker (`<w:cellIns>` / `<w:cellDel>` / `<w:cellMerge>`) — painter colors the top border. */
    trackedMarker?: CellMarker;
};
/**
 * A table row containing cells.
 */
type TableRow = {
    id: BlockId;
    cells: TableCell[];
    height?: number;
    heightRule?: 'auto' | 'atLeast' | 'exact';
    isHeader?: boolean;
    /**
     * `w:cantSplit` (§17.4.6): the row may not break across a page boundary.
     * The layout engine keeps such a row whole, moving it wholesale to the next
     * page rather than splitting its content.
     */
    cantSplit?: boolean;
    /** Tracked row ins / del (`<w:trPr><w:ins/>` / `<w:del/>`). */
    trackedIns?: RevisionInfo;
    /** see trackedIns */ trackedDel?: RevisionInfo;
};
/**
 * Floating table positioning info (pixel values).
 */
type FloatingTablePosition = {
    horzAnchor?: 'margin' | 'page' | 'text';
    vertAnchor?: 'margin' | 'page' | 'text';
    tblpX?: number;
    tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
    tblpY?: number;
    tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
    topFromText?: number;
    bottomFromText?: number;
    leftFromText?: number;
    rightFromText?: number;
};
/**
 * A table block containing rows.
 */
type TableBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'table';
    id: BlockId;
    rows: TableRow[];
    columnWidths?: number[];
    /** Table width value (twips for dxa, 50ths of percent for pct). */
    width?: number;
    /** Table width type ('auto', 'pct', 'dxa', 'nil'). */
    widthType?: string;
    /** Table horizontal alignment */
    justification?: 'left' | 'center' | 'right';
    /** Visual RTL column order (`w:bidiVisual`): painter renders logical column 0 rightmost. */
    bidi?: boolean;
    /** Table indent from left margin (in pixels, from w:tblInd) */
    indent?: number;
    /** Floating table properties (pixel values). */
    floating?: FloatingTablePosition;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * An anchored/floating image block.
 */
type ImageBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'image';
    id: BlockId;
    src: string;
    width: number;
    height: number;
    alt?: string;
    /** CSS transform string (rotation, flip) */
    transform?: string;
    anchor?: {
        isAnchored?: boolean;
        offsetH?: number;
        offsetV?: number;
        behindDoc?: boolean;
    };
    /** Hyperlink URL for clickable image */
    hlinkHref?: string;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * Section break block defining page layout changes.
 */
type SectionBreakBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'sectionBreak';
    id: BlockId;
    type?: 'continuous' | 'nextPage' | 'evenPage' | 'oddPage';
    pageSize?: {
        w: number;
        h: number;
    };
    orientation?: 'portrait' | 'landscape';
    margins?: PageMargins;
    columns?: ColumnLayout;
};
/**
 * Explicit page break block.
 */
type PageBreakBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'pageBreak';
    id: BlockId;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * Column break block.
 */
type ColumnBreakBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'columnBreak';
    id: BlockId;
    pmStart?: number;
    pmEnd?: number;
};
/** Default internal margins for text boxes (OOXML defaults in pixels) */
declare const DEFAULT_TEXTBOX_MARGINS: {
    top: number;
    bottom: number;
    left: number;
    right: number;
};
/** Default text box width in pixels when no width is specified */
declare const DEFAULT_TEXTBOX_WIDTH = 200;
/**
 * Text box block — positioned container with paragraph content.
 */
type TextBoxBlock = {
    /** Enclosing block-SDT group memberships (outermost first), if any. */
    sdtGroups?: SdtGroup[];
    kind: 'textBox';
    id: BlockId;
    /** Width in pixels */
    width: number;
    /** Height in pixels (may be auto-calculated) */
    height?: number;
    /** Fill/background color */
    fillColor?: string;
    /** Border width in pixels */
    outlineWidth?: number;
    /** Border color */
    outlineColor?: string;
    /** Border style */
    outlineStyle?: string;
    /** Internal padding */
    margins?: {
        top: number;
        bottom: number;
        left: number;
        right: number;
    };
    /** Paragraph blocks inside the text box */
    content: ParagraphBlock[];
    /** Display mode copied from the ProseMirror text box node */
    displayMode?: 'inline' | 'float' | 'block';
    /** CSS float direction copied from the ProseMirror text box node */
    cssFloat?: 'left' | 'right' | 'none';
    /** OOXML wrap type for anchored text boxes */
    wrapType?: string;
    /** OOXML wrapText direction */
    wrapText?: WrapTextDirection;
    /** Anchor target used during DOCX import/export */
    anchorTarget?: 'followingBlock';
    /** Position for floating/anchored text boxes */
    position?: ImageRunPosition;
    /** Wrap distances in pixels */
    distTop?: number;
    distBottom?: number;
    distLeft?: number;
    distRight?: number;
    pmStart?: number;
    pmEnd?: number;
};
/**
 * Identity of a block-level Structured Document Tag (content control) enclosing
 * a run of flow blocks. Block SDTs flatten into their child flow blocks for
 * pagination; each child carries its group(s) so the painter draws the boundary.
 */
interface SdtGroup {
    id: string;
    sdtType: string;
    tag?: string;
    alias?: string;
    lock?: string;
    checked?: boolean;
    bound?: boolean;
    repeatingItem?: boolean;
}
/**
 * Union of every block kind the layout engine knows about.
 *
 * Three switches over `block.kind` must stay in sync with this type:
 * - `runLayoutPipeline` in `layout-engine/index.ts` (this package)
 * - `measureBlock` in `packages/react/src/paged-editor/PagedEditor.tsx`
 * - `measureBlock` in `packages/vue/src/composables/useDocxEditor.ts`
 *
 * All three end in `assertExhaustiveFlowBlock(block, '<site>')` so adding
 * a new variant here without updating every site is a typecheck error.
 */
type FlowBlock = ParagraphBlock | TableBlock | ImageBlock | TextBoxBlock | SectionBreakBlock | PageBreakBlock | ColumnBreakBlock;
/**
 * Exhaustiveness guard for `FlowBlock`-shaped switches. Call from the
 * `default` arm with the still-typed value; TypeScript will refuse to
 * compile if any variant of `FlowBlock` was missed. The thrown error
 * names the calling site so runtime failures (e.g. an old adapter
 * compiled against a newer core) point future debuggers at the contract.
 */
declare function assertExhaustiveFlowBlock(block: never, site: string): never;
/**
 * A measured line within a paragraph.
 */
type MeasuredLine = {
    /** Starting run index (inclusive). */
    fromRun: number;
    /** Starting character index within fromRun. */
    fromChar: number;
    /** Ending run index (inclusive). */
    toRun: number;
    /** Ending character index within toRun (exclusive). */
    toChar: number;
    /** Total width of the line in pixels. */
    width: number;
    /** Ascent (height above baseline) in pixels. */
    ascent: number;
    /** Descent (height below baseline) in pixels. */
    descent: number;
    /** Total line height in pixels. */
    lineHeight: number;
    /** Left offset from floating images (pixels from content left edge). */
    leftOffset?: number;
    /** Right offset from floating images (pixels from content right edge). */
    rightOffset?: number;
    /** Optional split segments for centered floating exclusions. */
    segments?: MeasuredLineSegment[];
    /**
     * Vertical space inserted before this line to skip past floats that leave
     * no usable horizontal width at the natural line Y. Painters render this
     * as marginTop on the line element; measurement adds it to totalHeight.
     */
    floatSkipBefore?: number;
};
type MeasuredLineSegment = {
    fromRun: number;
    fromChar: number;
    toRun: number;
    toChar: number;
    width: number;
    leftOffset: number;
    availableWidth: number;
};
/**
 * Measurement result for a paragraph block.
 */
type ParagraphMeasure = {
    kind: 'paragraph';
    lines: MeasuredLine[];
    totalHeight: number;
};
/**
 * Measurement result for an image block.
 */
type ImageMeasure = {
    kind: 'image';
    width: number;
    height: number;
};
/**
 * Measurement result for a table cell.
 */
type TableCellMeasure = {
    blocks: Measure[];
    width: number;
    height: number;
    colSpan?: number;
    rowSpan?: number;
};
/**
 * Measurement result for a table row.
 */
type TableRowMeasure = {
    cells: TableCellMeasure[];
    height: number;
};
/**
 * Measurement result for a table block.
 */
type TableMeasure = {
    kind: 'table';
    rows: TableRowMeasure[];
    columnWidths: number[];
    totalWidth: number;
    totalHeight: number;
};
/**
 * Measurement result for section break (no visual size).
 */
type SectionBreakMeasure = {
    kind: 'sectionBreak';
};
/**
 * Measurement result for page break (no visual size).
 */
type PageBreakMeasure = {
    kind: 'pageBreak';
};
/**
 * Measurement result for column break (no visual size).
 */
type ColumnBreakMeasure = {
    kind: 'columnBreak';
};
/**
 * Measurement result for a text box block.
 */
type TextBoxMeasure = {
    kind: 'textBox';
    width: number;
    height: number;
    /** Pre-measured inner paragraph measures (avoids re-measuring during render) */
    innerMeasures: ParagraphMeasure[];
};
/**
 * Union of all measurement types.
 */
type Measure = ParagraphMeasure | ImageMeasure | TableMeasure | TextBoxMeasure | SectionBreakMeasure | PageBreakMeasure | ColumnBreakMeasure;
/**
 * Base fragment properties common to all fragment types.
 */
type FragmentBase = {
    /** Block ID this fragment belongs to. */
    blockId: BlockId;
    /** X position on page (relative to page left). */
    x: number;
    /** Y position on page (relative to page top). */
    y: number;
    /** Width of the fragment. */
    width: number;
    /** ProseMirror start position (for click mapping). */
    pmStart?: number;
    /** ProseMirror end position (for click mapping). */
    pmEnd?: number;
};
/**
 * A paragraph fragment positioned on a page.
 * May span only part of the paragraph's lines if split across pages.
 */
type ParagraphFragment = FragmentBase & {
    kind: 'paragraph';
    /** First line index (inclusive) from the measure. */
    fromLine: number;
    /** Last line index (exclusive) from the measure. */
    toLine: number;
    /** Height of this fragment. */
    height: number;
    /** True if this continues from a previous page. */
    continuesFromPrev?: boolean;
    /** True if this continues onto the next page. */
    continuesOnNext?: boolean;
};
/**
 * A table fragment positioned on a page.
 * May span only part of the table's rows if split across pages.
 */
type TableFragment = FragmentBase & {
    kind: 'table';
    /** First row index (inclusive). */
    fromRow: number;
    /** Last row index (exclusive). */
    toRow: number;
    /** Height of this fragment. */
    height: number;
    /** True if this is a floating table. */
    isFloating?: boolean;
    /** True if this continues from a previous page. */
    continuesFromPrev?: boolean;
    /** True if this continues onto the next page. */
    continuesOnNext?: boolean;
    /** Number of header rows prepended to this continuation fragment (0 or undefined for first fragment). */
    headerRowCount?: number;
    /**
     * Pixels to skip from the top of `fromRow`. Non-zero when this fragment's
     * first row is the continuation of a row that broke across a page boundary
     * (Word's "allow row to break across pages"). The painter renders the row
     * shifted up by this amount so the already-shown top slice is clipped.
     */
    topClip?: number;
    /**
     * Visible height (px) measured from the top of the LAST row (`toRow - 1`).
     * Set when that row breaks mid-content onto the next page; `undefined`
     * means the last row is fully visible. When `fromRow === toRow - 1`, the
     * visible band of that single row is `[topClip, bottomClip)`.
     */
    bottomClip?: number;
};
/**
 * An image fragment positioned on a page.
 */
type ImageFragment = FragmentBase & {
    kind: 'image';
    /** Height of the image. */
    height: number;
    /** True if this is an anchored/floating image. */
    isAnchored?: boolean;
    /** Z-index for layering. */
    zIndex?: number;
};
/**
 * A text box fragment positioned on a page.
 */
type TextBoxFragment = FragmentBase & {
    kind: 'textBox';
    /** Height of the text box. */
    height: number;
    /** True when positioned outside normal document flow. */
    isFloating?: boolean;
    /** Stack order hint for anchored text boxes. */
    zIndex?: number;
};
/**
 * Union of all fragment types.
 */
type Fragment = ParagraphFragment | TableFragment | ImageFragment | TextBoxFragment;
/**
 * Page margin configuration.
 */
type PageMargins = {
    top: number;
    right: number;
    bottom: number;
    left: number;
    /** Distance from page top to header content. */
    header?: number;
    /** Distance from page bottom to footer content. */
    footer?: number;
};
/**
 * A rendered page containing positioned fragments.
 */
type Page = {
    /** Page number (1-indexed). */
    number: number;
    /** Fragments positioned on this page. */
    fragments: Fragment[];
    /** Page margins. */
    margins: PageMargins;
    /** Page size (width, height). */
    size: {
        w: number;
        h: number;
    };
    /** Page orientation. */
    orientation?: 'portrait' | 'landscape';
    /** Section index this page belongs to. */
    sectionIndex?: number;
    /** Header/footer references for this page. */
    headerFooterRefs?: {
        headerDefault?: string;
        headerFirst?: string;
        headerEven?: string;
        footerDefault?: string;
        footerFirst?: string;
        footerEven?: string;
    };
    /** Footnote IDs that appear on this page (for rendering). */
    footnoteIds?: number[];
    /** Height reserved for the footnote area at page bottom (pixels). */
    footnoteReservedHeight?: number;
    /** Footnote-area columns (`w15:footnoteColumns`); absent/1 = single column. */
    footnoteColumns?: number;
    /** Column layout for this page (if multi-column). */
    columns?: ColumnLayout;
};
/**
 * Column layout configuration.
 */
type ColumnLayout = {
    count: number;
    gap: number;
    equalWidth?: boolean;
    /** Draw vertical separator line between columns (w:sep). */
    separator?: boolean;
};
/**
 * Header/footer layout for a specific type.
 */
type HeaderFooterLayout = {
    height: number;
    fragments: Fragment[];
};
/**
 * Final layout output ready for rendering/painting.
 */
type Layout = {
    /** Default page size for the document. */
    pageSize: {
        w: number;
        h: number;
    };
    /** All rendered pages with positioned fragments. */
    pages: Page[];
    /** Column configuration (if multi-column). */
    columns?: ColumnLayout;
    /** Header layouts by type (default, first, even). */
    headers?: Record<string, HeaderFooterLayout>;
    /** Footer layouts by type (default, first, even). */
    footers?: Record<string, HeaderFooterLayout>;
    /** Gap between pages in pixels (for rendering). */
    pageGap?: number;
};
/**
 * Header/footer content heights by variant type.
 */
type HeaderFooterContentHeights = Partial<Record<'default' | 'first' | 'even' | 'odd', number>>;
/**
 * Pre-calculated footnote content for layout and rendering.
 */
type FootnoteContent = {
    /** Footnote ID. */
    id: number;
    /** Display number (e.g. 1, 2, 3). */
    displayNumber: number;
    /** FlowBlocks for rendering the footnote content. */
    blocks: FlowBlock[];
    /** Measurements for the blocks. */
    measures: Measure[];
    /** Total height in pixels. */
    height: number;
};
/**
 * Options for the layout engine.
 */
type LayoutOptions = {
    /** Initial page size. */
    pageSize: {
        w: number;
        h: number;
    };
    /** Initial page margins. */
    margins: PageMargins;
    /** Body-level (final section) page size, used after the last explicit section break. */
    finalPageSize?: {
        w: number;
        h: number;
    };
    /** Body-level (final section) margins, used after the last explicit section break. */
    finalMargins?: PageMargins;
    /** Column configuration. */
    columns?: ColumnLayout;
    /** Gap between rendered pages (for UI). */
    pageGap?: number;
    /** Default line height multiplier. */
    defaultLineHeight?: number;
    /** Header content heights by variant. */
    headerContentHeights?: HeaderFooterContentHeights;
    /** Footer content heights by variant. */
    footerContentHeights?: HeaderFooterContentHeights;
    /** Whether section has different first page header/footer. */
    titlePage?: boolean;
    /** Whether section has different even/odd headers/footers. */
    evenAndOddHeaders?: boolean;
    /** Per-page footnote reserved heights (pageNumber → height in pixels). */
    footnoteReservedHeights?: Map<number, number>;
    /** Section break type for the body-level (final) section (for section transition logic). */
    bodyBreakType?: 'continuous' | 'nextPage' | 'evenPage' | 'oddPage';
};
/**
 * Result of hit-testing a click position.
 */
type HitTestResult = {
    /** Page index (0-based). */
    pageIndex: number;
    /** Fragment that was hit, if any. */
    fragment?: Fragment;
    /** Local X coordinate within the fragment. */
    localX?: number;
    /** Local Y coordinate within the fragment. */
    localY?: number;
};
/**
 * Position within the document model.
 */
type DocumentPosition = {
    /** Block index. */
    blockIndex: number;
    /** Run index within the block (for paragraphs). */
    runIndex?: number;
    /** Character offset within the run. */
    charOffset?: number;
    /** ProseMirror position. */
    pmPos?: number;
};

export { type BlockId, type BorderStyle, type CellBorderSpec, type CellBorders, type ColumnBreakBlock, type ColumnBreakMeasure, type ColumnLayout, DEFAULT_TEXTBOX_MARGINS, DEFAULT_TEXTBOX_WIDTH, type DocumentPosition, type FieldRun, type FloatingTablePosition, type FlowBlock, type FootnoteContent, type Fragment, type FragmentBase, type HeaderFooterContentHeights, type HeaderFooterLayout, type HitTestResult, type HyperlinkInfo, type ImageBlock, type ImageFragment, type ImageMeasure, type ImageRun, type ImageRunPosition, type InlineSdtWidget, type Layout, type LayoutOptions, type LineBreakRun, type ListNumPr, type Measure, type MeasuredLine, type MeasuredLineSegment, type Page, type PageBreakBlock, type PageBreakMeasure, type PageMargins, type ParagraphAttrs, type ParagraphBlock, type ParagraphBorders, type ParagraphFragment, type ParagraphIndent, type ParagraphMeasure, type ParagraphSpacing, type Run, type RunFormatting, type SdtGroup, type SectionBreakBlock, type SectionBreakMeasure, type TabAlignment, type TabRun, type TabStop, type TableBlock, type TableCell, type TableCellMeasure, type TableFragment, type TableMeasure, type TableRow, type TableRowMeasure, type TextBoxBlock, type TextBoxFragment, type TextBoxMeasure, type TextRun, type WrapTextDirection, assertExhaustiveFlowBlock };
