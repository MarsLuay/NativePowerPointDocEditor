/**
 * Color & Styling Primitives
 *
 * Basic types used throughout OOXML for colors, borders, and shading.
 */
/**
 * Theme color slots from theme1.xml
 */
type ThemeColorSlot = 'dk1' | 'lt1' | 'dk2' | 'lt2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink' | 'background1' | 'text1' | 'background2' | 'text2';
/**
 * ECMA-376 color reference — either a direct RGB hex, a theme slot
 * reference (with optional tint/shade), or `auto` for context-dependent
 * defaults (usually black for text on light backgrounds). When both
 * `rgb` and `themeColor` are set, the theme wins on Word import and the
 * `rgb` acts as a fallback for renderers without theme support.
 *
 * See ECMA-376 §17.18.39 (`ST_ThemeColor`).
 */
interface ColorValue {
    /** RGB hex value without # (e.g., "FF0000") */
    rgb?: string;
    /** Theme color slot reference */
    themeColor?: ThemeColorSlot;
    /** Tint modifier (0-255 as hex string, e.g., "80") - makes color lighter */
    themeTint?: string;
    /** Shade modifier (0-255 as hex string) - makes color darker */
    themeShade?: string;
    /** Auto color - context-dependent (usually black for text) */
    auto?: boolean;
}
/**
 * One side of a border — style, color, width, spacing. Used by paragraph
 * borders, table borders (per-cell or whole-table), and page borders.
 * `size` is in eighths of a point (Word's wire format); `space` is in
 * points.
 *
 * See ECMA-376 §17.18.2 (`ST_Border`).
 */
interface BorderSpec {
    /** Border style */
    style: 'none' | 'single' | 'double' | 'dotted' | 'dashed' | 'thick' | 'triple' | 'thinThickSmallGap' | 'thickThinSmallGap' | 'thinThickMediumGap' | 'thickThinMediumGap' | 'thinThickLargeGap' | 'thickThinLargeGap' | 'wave' | 'doubleWave' | 'dashSmallGap' | 'dashDotStroked' | 'threeDEmboss' | 'threeDEngrave' | 'outset' | 'inset' | 'nil';
    /** Color of the border */
    color?: ColorValue;
    /** Width in eighths of a point (1/8 pt) */
    size?: number;
    /** Spacing from text in points */
    space?: number;
    /** Shadow effect */
    shadow?: boolean;
    /** Frame effect */
    frame?: boolean;
}
/**
 * Cell/paragraph/run shading — Word's combined "fill + pattern overlay"
 * model. `fill` is the solid background; `color` is the pattern overlay
 * drawn on top; `pattern` selects the pattern type (defaults to
 * `'clear'` = solid `fill`, no pattern).
 *
 * See ECMA-376 §17.4.32 (`CT_Shd`).
 */
interface ShadingProperties {
    /** Pattern fill color */
    color?: ColorValue;
    /** Background fill color */
    fill?: ColorValue;
    /** Shading pattern type */
    pattern?: 'clear' | 'solid' | 'horzStripe' | 'vertStripe' | 'reverseDiagStripe' | 'diagStripe' | 'horzCross' | 'diagCross' | 'thinHorzStripe' | 'thinVertStripe' | 'thinReverseDiagStripe' | 'thinDiagStripe' | 'thinHorzCross' | 'thinDiagCross' | 'pct5' | 'pct10' | 'pct12' | 'pct15' | 'pct20' | 'pct25' | 'pct30' | 'pct35' | 'pct37' | 'pct40' | 'pct45' | 'pct50' | 'pct55' | 'pct60' | 'pct62' | 'pct65' | 'pct70' | 'pct75' | 'pct80' | 'pct85' | 'pct87' | 'pct90' | 'pct95' | 'nil';
}

/**
 * Text, Paragraph, and Table Formatting Types
 *
 * Properties that control how text, paragraphs, and table structures
 * are formatted in OOXML (w:rPr, w:pPr, w:tblPr, etc.).
 */

/**
 * Underline style options
 */
type UnderlineStyle = 'none' | 'single' | 'words' | 'double' | 'thick' | 'dotted' | 'dottedHeavy' | 'dash' | 'dashedHeavy' | 'dashLong' | 'dashLongHeavy' | 'dotDash' | 'dashDotHeavy' | 'dotDotDash' | 'dashDotDotHeavy' | 'wave' | 'wavyHeavy' | 'wavyDouble';
/**
 * Text effect animations
 */
type TextEffect = 'none' | 'blinkBackground' | 'lights' | 'antsBlack' | 'antsRed' | 'shimmer' | 'sparkle';
/**
 * Emphasis mark type
 */
type EmphasisMark = 'none' | 'dot' | 'comma' | 'circle' | 'underDot';
/**
 * Character-level formatting (`w:rPr`) — the full set of run properties
 * Word supports: weight, slant, font, size, color, highlight, underline,
 * strikethrough, vertical position, language, complex-script variants,
 * spacing/kerning, emphasis marks, and more.
 *
 * Most fields mirror their ECMA-376 element names (see §17.3.2). Missing
 * keys inherit from the run's paragraph style → linked style → document
 * defaults chain.
 */
interface TextFormatting {
    /** Bold (w:b) */
    bold?: boolean;
    /** Bold complex script (w:bCs) */
    boldCs?: boolean;
    /** Italic (w:i) */
    italic?: boolean;
    /** Italic complex script (w:iCs) */
    italicCs?: boolean;
    /** Underline style and color (w:u) */
    underline?: {
        style: UnderlineStyle;
        color?: ColorValue;
    };
    /** Strikethrough (w:strike) */
    strike?: boolean;
    /** Double strikethrough (w:dstrike) */
    doubleStrike?: boolean;
    /** Superscript/subscript (w:vertAlign) */
    vertAlign?: 'baseline' | 'superscript' | 'subscript';
    /** Small caps (w:smallCaps) */
    smallCaps?: boolean;
    /** All caps (w:caps) */
    allCaps?: boolean;
    /** Hidden text (w:vanish) */
    hidden?: boolean;
    /** Text color (w:color) */
    color?: ColorValue;
    /** Highlight/background color (w:highlight) */
    highlight?: 'black' | 'blue' | 'cyan' | 'darkBlue' | 'darkCyan' | 'darkGray' | 'darkGreen' | 'darkMagenta' | 'darkRed' | 'darkYellow' | 'green' | 'lightGray' | 'magenta' | 'none' | 'red' | 'white' | 'yellow';
    /** Character shading (w:shd) */
    shading?: ShadingProperties;
    /** Font size in half-points (w:sz) - e.g., 24 = 12pt */
    fontSize?: number;
    /** Font size complex script (w:szCs) */
    fontSizeCs?: number;
    /** Font family (w:rFonts) */
    fontFamily?: {
        ascii?: string;
        hAnsi?: string;
        eastAsia?: string;
        cs?: string;
        /** Theme font reference */
        asciiTheme?: 'majorAscii' | 'majorHAnsi' | 'majorEastAsia' | 'majorBidi' | 'minorAscii' | 'minorHAnsi' | 'minorEastAsia' | 'minorBidi';
        hAnsiTheme?: string;
        eastAsiaTheme?: string;
        csTheme?: string;
    };
    /** Character spacing in twips (w:spacing) */
    spacing?: number;
    /** Raised/lowered text position in half-points (w:position) */
    position?: number;
    /** Horizontal text scale percentage (w:w) */
    scale?: number;
    /** Kerning threshold in half-points (w:kern) */
    kerning?: number;
    /** Text effect animation (w:effect) */
    effect?: TextEffect;
    /** Emphasis mark (w:em) */
    emphasisMark?: EmphasisMark;
    /** Emboss effect (w:emboss) */
    emboss?: boolean;
    /** Imprint/engrave effect (w:imprint) */
    imprint?: boolean;
    /** Outline effect (w:outline) */
    outline?: boolean;
    /** Shadow effect (w:shadow) */
    shadow?: boolean;
    /** Right-to-left text (w:rtl) */
    rtl?: boolean;
    /** Complex script formatting (w:cs) */
    cs?: boolean;
    /** Character style ID (w:rStyle) */
    styleId?: string;
}
/**
 * Tab stop alignment
 */
type TabStopAlignment = 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';
/**
 * Tab leader character
 */
type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
/**
 * Tab stop definition
 */
interface TabStop {
    /** Position in twips from left margin */
    position: number;
    /** Alignment at tab stop */
    alignment: TabStopAlignment;
    /** Leader character */
    leader?: TabLeader;
}
/**
 * Line spacing rule
 */
type LineSpacingRule = 'auto' | 'exact' | 'atLeast';
/**
 * Paragraph alignment/justification
 */
type ParagraphAlignment = 'left' | 'center' | 'right' | 'both' | 'distribute' | 'mediumKashida' | 'highKashida' | 'lowKashida' | 'thaiDistribute';
/**
 * Per-side flags identifying which `<w:spacing>` attrs were inline (not
 * inherited from a style chain). Used to suppress style-only spacing on
 * empty paragraphs per Word's behavior.
 */
type SpacingExplicit = {
    before?: boolean;
    after?: boolean;
};
/**
 * Paragraph-level formatting (`w:pPr`) — alignment, indentation, spacing
 * (before/after, line height), pagination flags (keepNext, keepLines,
 * pageBreakBefore, widowControl), tabs, borders, shading, numbering
 * reference, style reference, and frame/anchored-text properties.
 *
 * Most fields mirror their ECMA-376 element names (see §17.3.1).
 * Inheritance: direct formatting beats the linked style which beats
 * document defaults.
 */
interface ParagraphFormatting {
    /** Paragraph alignment (w:jc) */
    alignment?: ParagraphAlignment;
    /** Text direction (w:bidi) */
    bidi?: boolean;
    /** Spacing before in twips (w:spacing/@w:before) */
    spaceBefore?: number;
    /** Spacing after in twips (w:spacing/@w:after) */
    spaceAfter?: number;
    /** Line spacing value (w:spacing/@w:line) */
    lineSpacing?: number;
    /** Line spacing rule (w:spacing/@w:lineRule) */
    lineSpacingRule?: LineSpacingRule;
    /** Auto space before (w:spacing/@w:beforeAutospacing) */
    beforeAutospacing?: boolean;
    /** Auto space after (w:spacing/@w:afterAutospacing) */
    afterAutospacing?: boolean;
    /**
     * Per-side flags marking which `<w:spacing>` attrs came from this
     * paragraph's own pPr (vs inherited). Word collapses style-inherited
     * spacing on empty paragraphs but honors the explicit values.
     */
    spacingExplicit?: SpacingExplicit;
    /** Left indent in twips (w:ind/@w:left) */
    indentLeft?: number;
    /** Right indent in twips (w:ind/@w:right) */
    indentRight?: number;
    /** First line indent in twips - positive for indent, negative for hanging (w:ind/@w:firstLine or @w:hanging) */
    indentFirstLine?: number;
    /** Whether first line is hanging indent */
    hangingIndent?: boolean;
    /** Paragraph borders (w:pBdr) */
    borders?: {
        top?: BorderSpec;
        bottom?: BorderSpec;
        left?: BorderSpec;
        right?: BorderSpec;
        between?: BorderSpec;
        bar?: BorderSpec;
    };
    /** Paragraph shading (w:shd) */
    shading?: ShadingProperties;
    /** Custom tab stops (w:tabs) */
    tabs?: TabStop[];
    /** Keep with next paragraph (w:keepNext) */
    keepNext?: boolean;
    /** Keep lines together (w:keepLines) */
    keepLines?: boolean;
    /** Widow/orphan control (w:widowControl) */
    widowControl?: boolean;
    /** Page break before (w:pageBreakBefore) */
    pageBreakBefore?: boolean;
    /** Contextual spacing — suppress space between paragraphs of the same style (w:contextualSpacing) */
    contextualSpacing?: boolean;
    /** Numbering properties (w:numPr) */
    numPr?: {
        /** Numbering definition ID (w:numId) */
        numId?: number;
        /** List level (0-8) (w:ilvl) */
        ilvl?: number;
    };
    /**
     * When `numPr` was resolved from the paragraph STYLE's pPr rather than the
     * paragraph's own `<w:numPr>`, this records the style-sourced value. The
     * serializer omits `numPr` while it still equals this value — writing it as
     * direct formatting would flip Word's indent precedence (a directly
     * referenced level's indents beat the style's; a style-referenced level's
     * do not) and break the document on save/reload. Cleared the moment the
     * user changes the numbering (values diverge).
     */
    numPrFromStyle?: {
        numId?: number;
        ilvl?: number;
    };
    /** Outline level 0-9 (w:outlineLvl) */
    outlineLevel?: number;
    /** Paragraph style ID (w:pStyle) */
    styleId?: string;
    /** Text frame properties (w:framePr) */
    frame?: {
        width?: number;
        height?: number;
        hAnchor?: 'text' | 'margin' | 'page';
        vAnchor?: 'text' | 'margin' | 'page';
        x?: number;
        y?: number;
        xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
        yAlign?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
        wrap?: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight';
    };
    /** Suppress line numbers (w:suppressLineNumbers) */
    suppressLineNumbers?: boolean;
    /** Suppress auto hyphens (w:suppressAutoHyphens) */
    suppressAutoHyphens?: boolean;
    /** Run properties to apply to all runs (w:rPr) */
    runProperties?: TextFormatting;
}
/**
 * Table width type
 */
type TableWidthType = 'auto' | 'dxa' | 'nil' | 'pct';
/**
 * Table measurement (width or height)
 */
interface TableMeasurement {
    /** Value in twips (for dxa) or fifths of a percent (for pct) */
    value: number;
    /** Measurement type */
    type: TableWidthType;
}
/**
 * Table borders
 */
interface TableBorders {
    top?: BorderSpec;
    bottom?: BorderSpec;
    left?: BorderSpec;
    right?: BorderSpec;
    insideH?: BorderSpec;
    insideV?: BorderSpec;
}
/**
 * Cell margins
 */
interface CellMargins {
    top?: TableMeasurement;
    bottom?: TableMeasurement;
    left?: TableMeasurement;
    right?: TableMeasurement;
}
/**
 * Table look flags (for table styles)
 */
interface TableLook {
    firstColumn?: boolean;
    firstRow?: boolean;
    lastColumn?: boolean;
    lastRow?: boolean;
    noHBand?: boolean;
    noVBand?: boolean;
}
/**
 * Floating table properties
 */
interface FloatingTableProperties {
    /** Horizontal anchor */
    horzAnchor?: 'margin' | 'page' | 'text';
    /** Vertical anchor */
    vertAnchor?: 'margin' | 'page' | 'text';
    /** Horizontal position */
    tblpX?: number;
    tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
    /** Vertical position */
    tblpY?: number;
    tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
    /** Distance from surrounding text */
    topFromText?: number;
    bottomFromText?: number;
    leftFromText?: number;
    rightFromText?: number;
}
/**
 * Table formatting properties (w:tblPr)
 */
interface TableFormatting {
    /** Table width */
    width?: TableMeasurement;
    /** Table justification */
    justification?: 'left' | 'center' | 'right';
    /** Cell spacing */
    cellSpacing?: TableMeasurement;
    /** Table indent from left margin */
    indent?: TableMeasurement;
    /** Table borders */
    borders?: TableBorders;
    /** Default cell margins */
    cellMargins?: CellMargins;
    /** Table layout */
    layout?: 'fixed' | 'autofit';
    /** Table style ID */
    styleId?: string;
    /** Table look (conditional formatting flags) */
    look?: TableLook;
    /** Shading/background */
    shading?: ShadingProperties;
    /** Overlap for floating tables */
    overlap?: 'never' | 'overlap';
    /** Floating table properties */
    floating?: FloatingTableProperties;
    /** Right to left table */
    bidi?: boolean;
}
/**
 * Table row formatting properties (w:trPr)
 */
interface TableRowFormatting {
    /** Row height */
    height?: TableMeasurement;
    /** Height rule */
    heightRule?: 'auto' | 'atLeast' | 'exact';
    /** Header row (repeats on each page) */
    header?: boolean;
    /** Allow row to break across pages */
    cantSplit?: boolean;
    /** Row justification */
    justification?: 'left' | 'center' | 'right';
    /** Hidden row */
    hidden?: boolean;
    /** Conditional format style */
    conditionalFormat?: ConditionalFormatStyle;
}
/**
 * Conditional format style
 */
interface ConditionalFormatStyle {
    /** First row */
    firstRow?: boolean;
    /** Last row */
    lastRow?: boolean;
    /** First column */
    firstColumn?: boolean;
    /** Last column */
    lastColumn?: boolean;
    /** Odd horizontal band */
    oddHBand?: boolean;
    /** Even horizontal band */
    evenHBand?: boolean;
    /** Odd vertical band */
    oddVBand?: boolean;
    /** Even vertical band */
    evenVBand?: boolean;
    /** Northwest corner */
    nwCell?: boolean;
    /** Northeast corner */
    neCell?: boolean;
    /** Southwest corner */
    swCell?: boolean;
    /** Southeast corner */
    seCell?: boolean;
}
/**
 * Table cell formatting properties (w:tcPr)
 */
interface TableCellFormatting {
    /** Cell width */
    width?: TableMeasurement;
    /** Cell borders */
    borders?: TableBorders;
    /** Cell margins (override table default) */
    margins?: CellMargins;
    /** Cell shading/background */
    shading?: ShadingProperties;
    /** Vertical alignment */
    verticalAlign?: 'top' | 'center' | 'bottom';
    /** Text direction */
    textDirection?: 'lr' | 'lrV' | 'rl' | 'rlV' | 'tb' | 'tbV' | 'tbRl' | 'tbRlV' | 'btLr';
    /** Grid span (horizontal merge) */
    gridSpan?: number;
    /** Vertical merge */
    vMerge?: 'restart' | 'continue';
    /** Fit text to cell width */
    fitText?: boolean;
    /** Wrap text */
    noWrap?: boolean;
    /** Hide cell marker */
    hideMark?: boolean;
    /** Conditional format style */
    conditionalFormat?: ConditionalFormatStyle;
}

/**
 * Lists & Numbering Types
 *
 * Types for bullet lists, numbered lists, and numbering definitions.
 */

/**
 * Number format type
 */
type NumberFormat = 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter' | 'ordinal' | 'cardinalText' | 'ordinalText' | 'hex' | 'chicago' | 'ideographDigital' | 'japaneseCounting' | 'aiueo' | 'iroha' | 'decimalFullWidth' | 'decimalHalfWidth' | 'japaneseLegal' | 'japaneseDigitalTenThousand' | 'decimalEnclosedCircle' | 'decimalFullWidth2' | 'aiueoFullWidth' | 'irohaFullWidth' | 'decimalZero' | 'decimalZero3' | 'decimalZero4' | 'decimalZero5' | 'bullet' | 'ganada' | 'chosung' | 'decimalEnclosedFullstop' | 'decimalEnclosedParen' | 'decimalEnclosedCircleChinese' | 'ideographEnclosedCircle' | 'ideographTraditional' | 'ideographZodiac' | 'ideographZodiacTraditional' | 'taiwaneseCounting' | 'ideographLegalTraditional' | 'taiwaneseCountingThousand' | 'taiwaneseDigital' | 'chineseCounting' | 'chineseLegalSimplified' | 'chineseCountingThousand' | 'koreanDigital' | 'koreanCounting' | 'koreanLegal' | 'koreanDigital2' | 'vietnameseCounting' | 'russianLower' | 'russianUpper' | 'none' | 'numberInDash' | 'hebrew1' | 'hebrew2' | 'arabicAlpha' | 'arabicAbjad' | 'hindiVowels' | 'hindiConsonants' | 'hindiNumbers' | 'hindiCounting' | 'thaiLetters' | 'thaiNumbers' | 'thaiCounting';
/**
 * Multi-level suffix (what follows the number)
 */
type LevelSuffix = 'tab' | 'space' | 'nothing';
/**
 * One indentation level of an abstract numbering definition (`w:lvl`).
 * Carries the number format, the marker template (`lvlText` — e.g.
 * `"%1.%2."`), the level's paragraph properties (indent, hanging) and
 * character properties (font, size, color for the marker itself).
 *
 * `ilvl` ranges 0-8 in standard Word documents.
 */
interface ListLevel {
    /** Level index (0-8) */
    ilvl: number;
    /** Starting number */
    start?: number;
    /** Number format */
    numFmt: NumberFormat;
    /** Level text (e.g., "%1." or "•") */
    lvlText: string;
    /** Justification */
    lvlJc?: 'left' | 'center' | 'right';
    /** Suffix after number */
    suffix?: LevelSuffix;
    /** Paragraph properties for this level */
    pPr?: ParagraphFormatting;
    /** Run properties for the number/bullet */
    rPr?: TextFormatting;
    /** Restart numbering from higher level */
    lvlRestart?: number;
    /** Is legal numbering style */
    isLgl?: boolean;
    /** Legacy settings */
    legacy?: {
        legacy?: boolean;
        legacySpace?: number;
        legacyIndent?: number;
    };
}
/**
 * Abstract numbering definition (`w:abstractNum`) — the reusable template
 * for a list: which `NumberFormat` at each indentation level, what
 * marker text, what paragraph/character formatting. Multiple
 * `NumberingInstance`s (`w:num`) can reference one abstract numbering
 * to share the template while keeping independent counters.
 *
 * See ECMA-376 §17.9.
 */
interface AbstractNumbering {
    /** Abstract numbering ID */
    abstractNumId: number;
    /** Multi-level type */
    multiLevelType?: 'hybridMultilevel' | 'multilevel' | 'singleLevel';
    /** Numbering style link */
    numStyleLink?: string;
    /** Style link */
    styleLink?: string;
    /** Level definitions */
    levels: ListLevel[];
    /** Name */
    name?: string;
}
/**
 * Numbering instance (w:num)
 */
interface NumberingInstance {
    /** Numbering ID (referenced by paragraphs) */
    numId: number;
    /** Reference to abstract numbering */
    abstractNumId: number;
    /** Level overrides */
    levelOverrides?: Array<{
        ilvl: number;
        startOverride?: number;
        lvl?: ListLevel;
    }>;
}
/**
 * Computed list marker for one paragraph — what the layout engine and
 * painter need to render the "1.", "a)", "•" prefix. Not part of the
 * wire format; the parser fills this from the `numbering.xml` chain plus
 * the paragraph's `numPr`. Paragraphs without list rendering omit it.
 */
interface ListRendering {
    /** Computed marker text (e.g., "1.", "a)", "•") */
    marker: string;
    /** List level (0-8) */
    level: number;
    /** Numbering ID */
    numId: number;
    /** Whether this is a bullet or numbered list */
    isBullet: boolean;
    /** Number format type (decimal, lowerRoman, upperRoman, etc.) */
    numFmt?: NumberFormat;
    /** Whether the list marker is hidden (w:vanish on level rPr) */
    markerHidden?: boolean;
    /** Marker font family from numbering level rPr (ascii name) */
    markerFontFamily?: string;
    /** Marker font size from numbering level rPr, in points */
    markerFontSize?: number;
    /**
     * Suffix character placed after the marker before body text (§17.9.25).
     * Default is `tab`; `space` inserts a single space; `nothing` no gap.
     * Drives marker-slot sizing in `getListMarkerInlineWidth`.
     */
    markerSuffix?: LevelSuffix;
    /**
     * NumberFormat for each level from 0..ilvl (inclusive).
     * Used to resolve multi-level templates like "%1.%2." where each %N
     * may need a different format (e.g., upperRoman parent + decimal child).
     */
    levelNumFmts?: NumberFormat[];
    /** abstractNumId the paragraph's numId points to (counters key on this). */
    abstractNumId?: number;
    /**
     * Start value from the numId's lvlOverride for the paragraph's ilvl, if any.
     * Per ECMA-376 §17.9.18, this resets the shared abstractNum counter the
     * first time the numId appears.
     */
    startOverride?: number;
}
/**
 * Top-level numbering data from `numbering.xml` — the set of abstract
 * templates and the per-document `NumberingInstance`s that reference
 * them. Paragraphs reference a `numId` (instance), not an
 * `abstractNumId` directly.
 */
interface NumberingDefinitions {
    /** Abstract numbering definitions */
    abstractNums: AbstractNumbering[];
    /** Numbering instances */
    nums: NumberingInstance[];
}

/**
 * OOXML image wrap-type taxonomy.
 *
 * `wp:inline` flows in the line. `wp:anchor` covers all positioned variants:
 *   - `square` / `tight` / `through` — text wraps around the image
 *   - `topAndBottom` — text breaks above and below
 *   - `behind` / `inFront` (`wp:wrapNone`) — image paints out of flow
 * @packageDocumentation
 * @public
 */
type WrapType = 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';

/**
 * Hyperlinks (`w:hyperlink`), bookmark markers (`w:bookmarkStart`/`End`),
 * and field types (`w:fldSimple`, complex `w:fldChar` runs).
 */

/**
 * Hyperlink (`w:hyperlink`) — wraps runs in a clickable link. External
 * targets resolve through the relationships part (`rId` → `href`);
 * internal targets reference a `BookmarkStart` anchor by name.
 */
interface Hyperlink {
    type: 'hyperlink';
    /** Relationship ID for external link */
    rId?: string;
    /** Resolved URL (from relationships) */
    href?: string;
    /** Internal bookmark anchor */
    anchor?: string;
    /** Tooltip text */
    tooltip?: string;
    /** Target frame */
    target?: string;
    /** Link history tracking */
    history?: boolean;
    /** Document location */
    docLocation?: string;
    /** Child runs */
    children: (Run | BookmarkStart | BookmarkEnd)[];
}
/**
 * Bookmark start marker (w:bookmarkStart)
 */
interface BookmarkStart {
    type: 'bookmarkStart';
    /** Bookmark ID */
    id: number;
    /** Bookmark name */
    name: string;
    /** Column index for table bookmarks */
    colFirst?: number;
    colLast?: number;
}
/**
 * Bookmark end marker (w:bookmarkEnd)
 */
interface BookmarkEnd {
    type: 'bookmarkEnd';
    /** Bookmark ID */
    id: number;
}
/**
 * Known field types
 */
type FieldType = 'PAGE' | 'NUMPAGES' | 'NUMWORDS' | 'NUMCHARS' | 'DATE' | 'TIME' | 'CREATEDATE' | 'SAVEDATE' | 'PRINTDATE' | 'AUTHOR' | 'TITLE' | 'SUBJECT' | 'KEYWORDS' | 'COMMENTS' | 'FILENAME' | 'FILESIZE' | 'TEMPLATE' | 'DOCPROPERTY' | 'DOCVARIABLE' | 'REF' | 'PAGEREF' | 'NOTEREF' | 'HYPERLINK' | 'TOC' | 'TOA' | 'INDEX' | 'SEQ' | 'STYLEREF' | 'AUTONUM' | 'AUTONUMLGL' | 'AUTONUMOUT' | 'IF' | 'MERGEFIELD' | 'NEXT' | 'NEXTIF' | 'ASK' | 'SET' | 'QUOTE' | 'INCLUDETEXT' | 'INCLUDEPICTURE' | 'SYMBOL' | 'ADVANCE' | 'EDITTIME' | 'REVNUM' | 'SECTION' | 'SECTIONPAGES' | 'USERADDRESS' | 'USERNAME' | 'USERINITIALS' | 'UNKNOWN';
/**
 * Simple field (w:fldSimple)
 */
interface SimpleField {
    type: 'simpleField';
    /** Field instruction (e.g., "PAGE \\* MERGEFORMAT") */
    instruction: string;
    /** Parsed field type */
    fieldType: FieldType;
    /** Current display value */
    content: (Run | Hyperlink)[];
    /** Field is locked */
    fldLock?: boolean;
    /** Field is dirty */
    dirty?: boolean;
}
/**
 * Complex field (w:fldChar begin/separate/end with w:instrText)
 */
interface ComplexField {
    type: 'complexField';
    /** Field instruction */
    instruction: string;
    /** Parsed field type */
    fieldType: FieldType;
    /** Field code runs */
    fieldCode: Run[];
    /** Display result runs */
    fieldResult: Run[];
    /**
     * Run formatting carried by the field's structural runs (the runs holding
     * the `w:fldChar` begin/separate/end). Word styles the field result with
     * this `w:rPr` when there is no separate result run (e.g. a `PAGE` field
     * collapsed into a single run). Used as a fallback for rendering and
     * serialization so the formatting survives the round-trip.
     */
    formatting?: TextFormatting;
    /** Field is locked */
    fldLock?: boolean;
    /** Field is dirty */
    dirty?: boolean;
}

/**
 * Math equations (`m:oMath`, `m:oMathPara`). OMML XML is round-tripped
 * verbatim to preserve fidelity Word/Pages/Docs can disagree on.
 */
/**
 * Math equation content (m:oMath or m:oMathPara)
 */
interface MathEquation {
    type: 'mathEquation';
    /** Whether this is a block (oMathPara) or inline (oMath) equation */
    display: 'inline' | 'block';
    /** Raw OMML XML for round-trip preservation */
    ommlXml: string;
    /** Plain text representation for accessibility/fallback */
    plainText?: string;
}

/**
 * Tracked-changes model — insertion/deletion/move wrappers, range
 * markers, and per-element property-change wrappers (`w:rPrChange`,
 * `w:pPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`) plus
 * structural changes (row/cell insert/delete/merge).
 */

/**
 * Tracked change metadata (w:ins, w:del attributes)
 */
interface TrackedChangeInfo {
    /** Revision ID */
    id: number;
    /** Author who made the change */
    author: string;
    /** Date of the change */
    date?: string;
}
/**
 * Generic tracked property-change wrapper metadata (w:*PrChange)
 */
interface PropertyChangeInfo extends TrackedChangeInfo {
    /** Optional revision session ID */
    rsid?: string;
}
/**
 * Insertion wrapper (w:ins) — runs inserted by tracked changes
 */
interface Insertion {
    type: 'insertion';
    /** Tracked change metadata */
    info: TrackedChangeInfo;
    /** Inserted content */
    content: (Run | Hyperlink)[];
}
/**
 * Deletion wrapper (w:del) — runs deleted by tracked changes
 */
interface Deletion {
    type: 'deletion';
    /** Tracked change metadata */
    info: TrackedChangeInfo;
    /** Deleted content */
    content: (Run | Hyperlink)[];
}
/**
 * Move-from wrapper (w:moveFrom) â€” content moved away from this position
 */
interface MoveFrom {
    type: 'moveFrom';
    /** Tracked change metadata */
    info: TrackedChangeInfo;
    /** Moved content */
    content: (Run | Hyperlink)[];
}
/**
 * Move-to wrapper (w:moveTo) â€” content moved into this position
 */
interface MoveTo {
    type: 'moveTo';
    /** Tracked change metadata */
    info: TrackedChangeInfo;
    /** Moved content */
    content: (Run | Hyperlink)[];
}
/**
 * Move-from range start marker (w:moveFromRangeStart) — ECMA-376 §17.13.5.22
 * Pairs with moveFromRangeEnd to delimit the source of a move in the document.
 */
interface MoveFromRangeStart {
    type: 'moveFromRangeStart';
    id: number;
    name: string;
}
/**
 * Move-from range end marker (w:moveFromRangeEnd)
 */
interface MoveFromRangeEnd {
    type: 'moveFromRangeEnd';
    id: number;
}
/**
 * Move-to range start marker (w:moveToRangeStart) — ECMA-376 §17.13.5.24
 * Pairs with moveToRangeEnd to delimit the destination of a move.
 */
interface MoveToRangeStart {
    type: 'moveToRangeStart';
    id: number;
    name: string;
}
/**
 * Move-to range end marker (w:moveToRangeEnd)
 */
interface MoveToRangeEnd {
    type: 'moveToRangeEnd';
    id: number;
}
/**
 * Run property change (w:rPrChange)
 */
interface RunPropertyChange {
    type: 'runPropertyChange';
    /** Tracked change metadata */
    info: PropertyChangeInfo;
    /** Run properties before the tracked change */
    previousFormatting?: TextFormatting;
    /** Run properties after the tracked change (editor model convenience) */
    currentFormatting?: TextFormatting;
}
/**
 * Paragraph property change (w:pPrChange)
 */
interface ParagraphPropertyChange {
    type: 'paragraphPropertyChange';
    /** Tracked change metadata */
    info: PropertyChangeInfo;
    /** Paragraph properties before the tracked change */
    previousFormatting?: ParagraphFormatting;
    /** Paragraph properties after the tracked change (editor model convenience) */
    currentFormatting?: ParagraphFormatting;
}
/**
 * Table property change (w:tblPrChange)
 */
interface TablePropertyChange {
    type: 'tablePropertyChange';
    /** Tracked change metadata */
    info: PropertyChangeInfo;
    /** Table properties before the tracked change */
    previousFormatting?: TableFormatting;
    /** Table properties after the tracked change (editor model convenience) */
    currentFormatting?: TableFormatting;
}
/**
 * Table row property change (w:trPrChange)
 */
interface TableRowPropertyChange {
    type: 'tableRowPropertyChange';
    /** Tracked change metadata */
    info: PropertyChangeInfo;
    /** Row properties before the tracked change */
    previousFormatting?: TableRowFormatting;
    /** Row properties after the tracked change (editor model convenience) */
    currentFormatting?: TableRowFormatting;
}
/**
 * Table cell property change (w:tcPrChange)
 */
interface TableCellPropertyChange {
    type: 'tableCellPropertyChange';
    /** Tracked change metadata */
    info: PropertyChangeInfo;
    /** Cell properties before the tracked change */
    previousFormatting?: TableCellFormatting;
    /** Cell properties after the tracked change (editor model convenience) */
    currentFormatting?: TableCellFormatting;
}
/**
 * Table structural tracked change metadata (row/cell insert/delete/merge)
 */
interface TableStructuralChangeInfo {
    type: 'tableRowInsertion' | 'tableRowDeletion' | 'tableCellInsertion' | 'tableCellDeletion' | 'tableCellMerge';
    /** Tracked change metadata */
    info: TrackedChangeInfo;
    /**
     * `<w:cellMerge w:vMerge="…">` value, only meaningful for `tableCellMerge`.
     * Schema `ST_AnnotationVMerge`: `"rest"` = anchor (start of merged span),
     * `"cont"` = continuation (merged into predecessor). Word's default for a
     * tracked merge is `"cont"` (most edits track "this cell got merged INTO
     * the one above"); we preserve the on-disk value when present.
     */
    vMerge?: 'rest' | 'cont';
    /** `<w:cellMerge w:vMergeOrig="…">` — the pre-merge vMerge state. */
    vMergeOrig?: 'rest' | 'cont';
}

/**
 * Tables (`w:tbl`), rows (`w:tr`), and cells (`w:tc`).
 */

/**
 * Table cell (`w:tc`). Holds nested block content (paragraphs and nested
 * tables), cell-level formatting (borders, shading, vertical merge),
 * tracked property changes, and tracked structural changes for cell
 * insert/delete/merge operations.
 */
interface TableCell {
    type: 'tableCell';
    /** Cell formatting */
    formatting?: TableCellFormatting;
    /** Cell-level tracked property changes (w:tcPrChange) */
    propertyChanges?: TableCellPropertyChange[];
    /** Tracked structural changes (cell insert/delete/merge) */
    structuralChange?: TableStructuralChangeInfo;
    /** Cell content (paragraphs, tables, etc.) */
    content: (Paragraph | Table)[];
}
/**
 * Table row (`w:tr`) — an ordered list of `TableCell` plus row-level
 * formatting (height, repeated header, cantSplit) and tracked changes
 * for inserts/deletes.
 */
interface TableRow {
    type: 'tableRow';
    /** Row formatting */
    formatting?: TableRowFormatting;
    /** Row-level tracked property changes (w:trPrChange) */
    propertyChanges?: TableRowPropertyChange[];
    /** Tracked structural changes (row insert/delete) */
    structuralChange?: TableStructuralChangeInfo;
    /** Cells in this row */
    cells: TableCell[];
}
/**
 * Table (`w:tbl`) — a block-level grid of rows × cells. Tables carry
 * their own formatting layer (borders, shading, alignment, indent,
 * floating placement) and an explicit column-width grid in twips. Tables
 * can nest arbitrarily through `TableCell.content`.
 *
 * See ECMA-376 §17.4.
 */
interface Table {
    type: 'table';
    /** Table formatting */
    formatting?: TableFormatting;
    /** Table-level tracked property changes (w:tblPrChange) */
    propertyChanges?: TablePropertyChange[];
    /** Column widths in twips */
    columnWidths?: number[];
    /** Table rows */
    rows: TableRow[];
}

/**
 * Watermark — the MS Word "Design → Watermark" feature.
 *
 * In OOXML a watermark is legacy VML stored inside a header part, inside a
 * paragraph run: `<w:pict><v:shape>…</v:shape></w:pict>`. It renders behind
 * the body content on every page of the section. Word supports two kinds:
 *
 * - **Text** — `<v:shape type="#_x0000_t136"><v:textpath string="DRAFT"/>`
 *   with a fill color, font, optional rotation (diagonal vs horizontal) and a
 *   reduced-opacity "semitransparent" look.
 * - **Picture** — `<v:shape><v:imagedata r:id="rIdN"/>` referencing a media
 *   file, scaled, optionally "washed out" (lightened).
 *
 * We model the watermark as a dedicated field on the owning `HeaderFooter`
 * rather than as editable run content, so it stays out of the ProseMirror text
 * flow while still round-tripping through parse → render → serialize.
 */
/**
 * Text watermark (e.g. "CONFIDENTIAL", "DRAFT").
 */
interface TextWatermark {
    kind: 'text';
    /** The watermark text. */
    text: string;
    /** Font family (e.g. 'Calibri'). */
    font: string;
    /** Fill color as a CSS hex string (e.g. '#C0C0C0'). */
    color: string;
    /** Word's "Semitransparent" checkbox — renders at reduced opacity. */
    semitransparent: boolean;
    /** Diagonal (≈ -45°) or horizontal layout. */
    layout: 'diagonal' | 'horizontal';
    /** Font size in points. When undefined the renderer auto-sizes to the page (Word's "Auto"). */
    fontSize?: number;
}
/**
 * Picture watermark — a scaled, optionally washed-out background image.
 */
interface PictureWatermark {
    kind: 'picture';
    /** Header-part relationship id of the media (set for images parsed from an existing file). */
    relId?: string;
    /** Package path of the media, e.g. 'word/media/image1.png'. */
    mediaPath?: string;
    /** Raw bytes for an image added in-editor (no rId yet). */
    data?: Uint8Array;
    /** MIME type for `data`. */
    contentType?: string;
    /** Ready-to-use data URL for rendering (resolved from media or `data`). */
    dataUrl?: string;
    /** Scale factor; 1 = 100% / Word's "Auto". */
    scale: number;
    /** Word's "Washout" checkbox — lightens the image so text stays readable. */
    washout: boolean;
    /** Natural width in EMUs (when known). */
    widthEmu?: number;
    /** Natural height in EMUs (when known). */
    heightEmu?: number;
}
/**
 * A document watermark — text or picture.
 */
type Watermark = TextWatermark | PictureWatermark;

/**
 * Page furniture — headers (`w:hdr`), footers (`w:ftr`), footnotes
 * (`w:footnote`), and endnotes (`w:endnote`), plus the section-level
 * properties (`w:footnotePr`/`w:endnotePr`) that configure note layout.
 */

/**
 * Header/footer type
 */
type HeaderFooterType = 'default' | 'first' | 'even';
/**
 * Header or footer reference
 */
interface HeaderReference {
    type: HeaderFooterType;
    rId: string;
}
interface FooterReference {
    type: HeaderFooterType;
    rId: string;
}
/**
 * Header or footer content
 */
interface HeaderFooter {
    type: 'header' | 'footer';
    /** Header/footer type */
    hdrFtrType: HeaderFooterType;
    /** Content (paragraphs, tables, etc.) */
    content: BlockContent[];
    /**
     * Watermark stored on this header (MS Word "Design → Watermark"). Lives
     * here, not in `content`, so it stays out of the editable text flow while
     * still round-tripping. Only headers carry watermarks; footers never do.
     */
    watermark?: Watermark;
}
/**
 * Footnote position
 */
type FootnotePosition = 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';
/**
 * Endnote position
 */
type EndnotePosition = 'sectEnd' | 'docEnd';
/**
 * Number restart type
 */
type NoteNumberRestart = 'continuous' | 'eachSect' | 'eachPage';
/**
 * Footnote properties
 */
interface FootnoteProperties {
    position?: FootnotePosition;
    numFmt?: NumberFormat;
    numStart?: number;
    numRestart?: NoteNumberRestart;
}
/**
 * Endnote properties
 */
interface EndnoteProperties {
    position?: EndnotePosition;
    numFmt?: NumberFormat;
    numStart?: number;
    numRestart?: NoteNumberRestart;
}
/**
 * Footnote (w:footnote)
 */
interface Footnote {
    type: 'footnote';
    /** Footnote ID */
    id: number;
    /** Special footnote type */
    noteType?: 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice';
    /**
     * Content. Per ECMA-376 §17.11.10 footnotes can hold the same blocks as
     * the body, so the note parser reuses the body's `parseBlockContent`: the
     * full block model — paragraphs, tables, and block-level `w:sdt` content
     * controls (as `BlockSdt`) — flows through the body pipeline
     * (toProseDoc → toFlowBlocks) and stays editable on round-trip.
     */
    content: BlockContent[];
    /**
     * Verbatim original XML of the entire `<w:footnote>` element, captured at
     * parse time ONLY when the note body carries a block-level construct the
     * model still can't represent — note-level bookmarks
     * (`w:bookmarkStart`/`w:bookmarkEnd`) or `w:customXml`. Block-level `w:sdt`
     * is NOT a trigger: it round-trips through the model as `BlockSdt`. When
     * present the serializer re-emits these bytes instead of rebuilding from
     * `content`, restoring pre-#646 fidelity for the unmodeled constructs.
     * See `parseNoteBlockContent` / `serializeNote` for the gate (#646 F3).
     */
    verbatimXml?: string;
}
/**
 * Endnote (w:endnote)
 */
interface Endnote {
    type: 'endnote';
    /** Endnote ID */
    id: number;
    /** Special endnote type */
    noteType?: 'normal' | 'separator' | 'continuationSeparator' | 'continuationNotice';
    /**
     * Content. Per ECMA-376 §17.11.4 endnotes can hold the same blocks as
     * the body — paragraphs, tables, and block-level content controls. See note
     * on `Footnote.content`.
     */
    content: BlockContent[];
    /** Verbatim original XML — see `Footnote.verbatimXml` (#646 F3). */
    verbatimXml?: string;
}

/**
 * Comments (`w:comment` in `comments.xml`) and the inline range markers
 * (`w:commentRangeStart`/`End`) that anchor them inside paragraphs.
 */

/**
 * A comment from `comments.xml` — the top-level entity for review
 * comments and replies. `id` matches the inline `CommentRangeStart` /
 * `CommentRangeEnd` markers that anchor it inside a paragraph; `parentId`
 * threads replies under their parent; `done` reflects Word's "Resolve"
 * state (`w15:done`).
 */
interface Comment {
    /** Comment ID (matches commentRangeStart/End) */
    id: number;
    /** Author name */
    author: string;
    /** Author initials */
    initials?: string;
    /** Date */
    date?: string;
    /** Comment content (paragraphs) */
    content: Paragraph[];
    /** Parent comment ID (for replies) */
    parentId?: number;
    /** Whether the comment is resolved/done */
    done?: boolean;
}
/**
 * Comment range start marker in paragraph content
 */
interface CommentRangeStart {
    type: 'commentRangeStart';
    id: number;
}
/**
 * Comment range end marker in paragraph content
 */
interface CommentRangeEnd {
    type: 'commentRangeEnd';
    id: number;
}

/**
 * Section properties (`w:sectPr`) — page size and margins, columns,
 * header/footer references, line numbers, page borders, document grid,
 * paper sources — plus the section and document-body containers that
 * group block-level content.
 */

/**
 * Page orientation
 */
type PageOrientation = 'portrait' | 'landscape';
/**
 * Section start type
 */
type SectionStart = 'continuous' | 'nextPage' | 'oddPage' | 'evenPage' | 'nextColumn';
/**
 * Vertical alignment
 */
type VerticalAlign = 'top' | 'center' | 'both' | 'bottom';
/**
 * Line number restart type
 */
type LineNumberRestart = 'continuous' | 'newPage' | 'newSection';
/**
 * Column definition
 */
interface Column {
    /** Column width in twips */
    width?: number;
    /** Space after column in twips */
    space?: number;
}
/**
 * Section properties (`w:sectPr`) — page geometry, margins, columns,
 * header/footer references, and page numbering for one section of the
 * document. Sections are introduced by inline `sectPr` markers on the
 * terminating paragraph (`Paragraph.sectionProperties`) and the body's
 * final `sectPr`.
 *
 * All distance units are twips (1/20 of a point) on the wire. The layout
 * engine converts to pixels.
 *
 * See ECMA-376 §17.6.
 */
interface SectionProperties {
    /** Page width in twips */
    pageWidth?: number;
    /** Page height in twips */
    pageHeight?: number;
    /** Page orientation */
    orientation?: PageOrientation;
    /** Top margin in twips */
    marginTop?: number;
    /** Bottom margin in twips */
    marginBottom?: number;
    /** Left margin in twips */
    marginLeft?: number;
    /** Right margin in twips */
    marginRight?: number;
    /** Header distance from top in twips */
    headerDistance?: number;
    /** Footer distance from bottom in twips */
    footerDistance?: number;
    /** Gutter margin in twips */
    gutter?: number;
    /** Number of columns */
    columnCount?: number;
    /** Space between columns in twips */
    columnSpace?: number;
    /** Equal width columns */
    equalWidth?: boolean;
    /** Separator line between columns */
    separator?: boolean;
    /** Individual column definitions */
    columns?: Column[];
    /**
     * Number of columns the footnote area is laid out in (`w15:footnoteColumns`).
     * Word's "Footnote layout → Columns" setting, independent of the body column
     * count above. Undefined/1 means the footnote area follows the body (single
     * column for a single-column section). See ECMA-376 + the w15 extension.
     */
    footnoteColumns?: number;
    /** Section start type */
    sectionStart?: SectionStart;
    /** Vertical alignment of text */
    verticalAlign?: VerticalAlign;
    /** Right-to-left section */
    bidi?: boolean;
    /** Header references */
    headerReferences?: HeaderReference[];
    /** Footer references */
    footerReferences?: FooterReference[];
    /** Different first page header/footer */
    titlePg?: boolean;
    /** Different odd/even page headers/footers */
    evenAndOddHeaders?: boolean;
    /** Line numbering settings */
    lineNumbers?: {
        start?: number;
        countBy?: number;
        distance?: number;
        restart?: LineNumberRestart;
    };
    /** Page borders */
    pageBorders?: {
        top?: BorderSpec;
        bottom?: BorderSpec;
        left?: BorderSpec;
        right?: BorderSpec;
        /** Display setting */
        display?: 'allPages' | 'firstPage' | 'notFirstPage';
        /** Offset from */
        offsetFrom?: 'page' | 'text';
        /** Z-order */
        zOrder?: 'front' | 'back';
    };
    /** Page background */
    background?: {
        color?: ColorValue;
        themeColor?: ThemeColorSlot;
        themeTint?: string;
        themeShade?: string;
    };
    /** Footnote properties for this section */
    footnotePr?: FootnoteProperties;
    /** Endnote properties for this section */
    endnotePr?: EndnoteProperties;
    /** Document grid */
    docGrid?: {
        type?: 'default' | 'lines' | 'linesAndChars' | 'snapToChars';
        linePitch?: number;
        charSpace?: number;
    };
    /** First page paper source */
    paperSrcFirst?: number;
    /** Other pages paper source */
    paperSrcOther?: number;
}
/**
 * Block-level content types
 */
type BlockContent = Paragraph | Table | BlockSdt;
/**
 * One section of the document — a `SectionProperties` plus the block
 * content (`Paragraph`s and `Table`s) that lives under those properties.
 *
 * Sections are derived during parse: every paragraph carrying an inline
 * `sectPr` ends a section, and the body's final `sectPr` defines the
 * last section. Each section may carry its own headers/footers map.
 */
interface Section {
    /** Section properties */
    properties: SectionProperties;
    /** Content in this section */
    content: BlockContent[];
    /** Headers for this section */
    headers?: Map<HeaderFooterType, HeaderFooter>;
    /** Footers for this section */
    footers?: Map<HeaderFooterType, HeaderFooter>;
}
/**
 * Document body (`w:body`) — the editable content of the document.
 *
 * Contains the ordered block content (paragraphs and tables), the section
 * layout chain derived from inline `sectPr` markers, the final `sectPr`,
 * and any document-level comments. This is what most edit operations
 * mutate; headers/footers/styles live elsewhere in the package.
 */
interface DocumentBody {
    /** All content (paragraphs, tables) */
    content: BlockContent[];
    /** Sections (derived from sectPr in paragraphs and final sectPr) */
    sections?: Section[];
    /** Final section properties (from body's sectPr) */
    finalSectionProperties?: SectionProperties;
    /** Comments from comments.xml */
    comments?: Comment[];
}

/**
 * Structured Document Tags / content controls (`w:sdt`) — inline and
 * block variants, plus properties (alias, tag, lock, list items,
 * checkbox state) for the supported SDT types.
 */

/**
 * SDT type (content control type).
 *
 * Values mirror the `w:sdtPr` type-marker element names from ECMA-376
 * §17.5.2 (`CT_SdtPr`), with two deliberate exceptions:
 * - `checkbox` is the `w14:checkbox` (Office 2010) extension, not a base
 *   OOXML type marker.
 * - `buildingBlockGallery` covers both `w:docPartObj` and `w:docPartList`.
 *
 * A `w:sdtPr` with no type marker means `richText` (the spec default). A
 * type marker the parser does not model maps to `unknown` — it is never
 * coerced to `richText`, so the projection stays honest. Round-trip
 * fidelity does not depend on this enum: the raw `w:sdtPr` is replayed
 * verbatim (see `rawPropertiesXml`).
 */
type SdtType = 'richText' | 'plainText' | 'date' | 'dropDownList' | 'comboBox' | 'checkbox' | 'picture' | 'buildingBlockGallery' | 'group' | 'equation' | 'citation' | 'bibliography' | 'unknown';
/**
 * XML data binding (`w:dataBinding`) — links a content control to a node in a
 * Custom XML data store. Modeled read-only; the binding round-trips verbatim
 * via `rawPropertiesXml` (this projection is for inspection, e.g. "which
 * controls are bound, and to what XPath"). The editor does not resolve or
 * sync bound values.
 */
interface SdtDataBinding {
    /** XPath into the bound Custom XML part (`w:xpath`). */
    xpath?: string;
    /** Target Custom XML store id (`w:storeItemID`). */
    storeItemID?: string;
    /** Namespace prefix mappings used by the XPath (`w:prefixMappings`). */
    prefixMappings?: string;
}
/**
 * SDT properties (`w:sdtPr`).
 *
 * The modeled fields are a **read-only projection** for downstream tooling
 * (tag/alias addressing, template extraction). They are NOT the
 * serialization source: the original `w:sdtPr` is captured verbatim in
 * `rawPropertiesXml` and replayed on save, which preserves element order
 * (`CT_SdtPr` is an `xsd:sequence`), avoids double-emission, and keeps
 * unmodeled features (data binding, `w15:*`, `@lastValue`) lossless.
 */
interface SdtProperties {
    /** SDT type (projection; see {@link SdtType}). */
    sdtType: SdtType;
    /** Unique numeric id (`w:id`, signed). */
    id?: number;
    /** Alias (friendly name, `w:alias`). */
    alias?: string;
    /** Tag (developer identifier, `w:tag`). */
    tag?: string;
    /** Lock setting (`w:lock`). */
    lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';
    /**
     * Placeholder building-block name (`w:placeholder/w:docPart@w:val`).
     * This is a reference to a glossary docPart that supplies the placeholder
     * content — NOT the literal placeholder text.
     */
    placeholder?: string;
    /** Whether the control is currently showing its placeholder (`w:showingPlcHdr`). */
    showingPlaceholder?: boolean;
    /** Date display format for date controls (`w:date/w:dateFormat@w:val`). */
    dateFormat?: string;
    /** Dropdown/combobox list items. */
    listItems?: {
        displayText: string;
        value: string;
    }[];
    /** Checkbox checked state (`w14:checkbox`). */
    checked?: boolean;
    /** XML data binding (`w:dataBinding`), if the control is bound. */
    dataBinding?: SdtDataBinding;
    /**
     * The original `<w:sdtPr>` serialized verbatim as an XML string, captured
     * at parse time. Replayed unchanged on save so the properties block
     * round-trips losslessly. Stored as a string (not an `XmlElement`) so the
     * types layer stays free of the parser/`xml-js` dependency. Absent for
     * SDTs created programmatically — the serializer then synthesizes a
     * minimal, sequence-valid `w:sdtPr` from the modeled fields.
     */
    rawPropertiesXml?: string;
    /** The original `<w:sdtEndPr>` serialized verbatim, if present. */
    rawEndPropertiesXml?: string;
}
/**
 * Inline SDT (content control within a paragraph)
 */
interface InlineSdt {
    type: 'inlineSdt';
    /** SDT properties */
    properties: SdtProperties;
    /**
     * Inline content held inside the control. OOXML allows runs,
     * hyperlinks, simple/complex fields, nested SDTs, and math at this
     * level; the renderer must descend into all of them so docProps-bound
     * fields and similar template content survive paged rendering.
     */
    content: (Run | Hyperlink | SimpleField | ComplexField | InlineSdt | MathEquation)[];
}
/**
 * Block-level SDT (content control wrapping block content).
 *
 * `content` is `BlockContent[]` (not just paragraphs/tables) so a nested
 * block SDT survives the round trip. `CT_SdtContentBlock` also permits
 * run-level content (bookmarks, etc.); that is carried through the same
 * block-content parsing as elsewhere in the document.
 */
interface BlockSdt {
    type: 'blockSdt';
    /** SDT properties */
    properties: SdtProperties;
    /** Block content inside the control */
    content: BlockContent[];
}

/**
 * Paragraph (`w:p`) — the union of inline content that can sit inside a
 * paragraph (runs, hyperlinks, bookmarks, fields, SDT, comment ranges,
 * tracked-change wrappers, math) plus paragraph-level metadata
 * (formatting, list rendering, optional terminating section properties).
 */

/**
 * Inline content that can appear inside a paragraph. Covers runs (text),
 * hyperlinks, bookmarks, fields, structured document tags, comment range
 * markers, tracked-change wrappers, and math equations. Every node in
 * this union carries a `type` discriminator so consumers can narrow at
 * runtime.
 */
type ParagraphContent = Run | Hyperlink | BookmarkStart | BookmarkEnd | SimpleField | ComplexField | InlineSdt | CommentRangeStart | CommentRangeEnd | Insertion | Deletion | MoveFrom | MoveTo | MoveFromRangeStart | MoveFromRangeEnd | MoveToRangeStart | MoveToRangeEnd | MathEquation;
/**
 * Paragraph (`w:p`) — the primary block-level container in a Word document.
 *
 * Every paragraph carries direct formatting (`formatting`), tracked
 * property changes (`propertyChanges`), inline content (`content`), and
 * optional list rendering / section break metadata. `paraId` is Word's
 * stable identifier (`w14:paraId`) and is what `EditorBridge` and the
 * agent toolkit use to address paragraphs.
 *
 * See ECMA-376 §17.3.1.
 */
interface Paragraph {
    type: 'paragraph';
    /** Unique paragraph ID */
    paraId?: string;
    /** Text ID */
    textId?: string;
    /** Paragraph formatting */
    formatting?: ParagraphFormatting;
    /** Paragraph-level tracked property changes (w:pPrChange) */
    propertyChanges?: ParagraphPropertyChange[];
    /**
     * Paragraph-mark insertion tracking (`<w:pPr><w:rPr><w:ins/>`). Set when
     * this paragraph's terminating pilcrow was added as a tracked change —
     * e.g., the user pressed Enter mid-paragraph in suggesting mode. Reject
     * joins this paragraph with the following one.
     */
    pPrIns?: TrackedChangeInfo;
    /**
     * Paragraph-mark deletion tracking (`<w:pPr><w:rPr><w:del/>`). Set when
     * this paragraph's terminating pilcrow was deleted as a tracked change —
     * e.g., the user pressed Backspace at the start of the next paragraph in
     * suggesting mode. Accept joins this paragraph with the following one.
     */
    pPrDel?: TrackedChangeInfo;
    /** Paragraph content */
    content: ParagraphContent[];
    /** Computed list rendering (if this is a list item) */
    listRendering?: ListRendering;
    /** Word's cached layout says this paragraph started on a new rendered page. */
    renderedPageBreakBefore?: boolean;
    /** Section properties (if this paragraph ends a section) */
    sectionProperties?: SectionProperties;
}

/**
 * DrawingML shapes (`wps:wsp`) and text boxes — preset shape types,
 * fill, outline, shape text body, transform.
 */

/**
 * Shape types
 */
type ShapeType = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rtTriangle' | 'parallelogram' | 'trapezoid' | 'pentagon' | 'hexagon' | 'heptagon' | 'octagon' | 'decagon' | 'dodecagon' | 'star4' | 'star5' | 'star6' | 'star7' | 'star8' | 'star10' | 'star12' | 'star16' | 'star24' | 'star32' | 'line' | 'straightConnector1' | 'bentConnector2' | 'bentConnector3' | 'bentConnector4' | 'bentConnector5' | 'curvedConnector2' | 'curvedConnector3' | 'curvedConnector4' | 'curvedConnector5' | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'leftRightArrow' | 'upDownArrow' | 'quadArrow' | 'leftRightUpArrow' | 'bentArrow' | 'uturnArrow' | 'leftUpArrow' | 'bentUpArrow' | 'curvedRightArrow' | 'curvedLeftArrow' | 'curvedUpArrow' | 'curvedDownArrow' | 'stripedRightArrow' | 'notchedRightArrow' | 'homePlate' | 'chevron' | 'rightArrowCallout' | 'downArrowCallout' | 'leftArrowCallout' | 'upArrowCallout' | 'leftRightArrowCallout' | 'quadArrowCallout' | 'circularArrow' | 'flowChartProcess' | 'flowChartAlternateProcess' | 'flowChartDecision' | 'flowChartInputOutput' | 'flowChartPredefinedProcess' | 'flowChartInternalStorage' | 'flowChartDocument' | 'flowChartMultidocument' | 'flowChartTerminator' | 'flowChartPreparation' | 'flowChartManualInput' | 'flowChartManualOperation' | 'flowChartConnector' | 'flowChartOffpageConnector' | 'flowChartPunchedCard' | 'flowChartPunchedTape' | 'flowChartSummingJunction' | 'flowChartOr' | 'flowChartCollate' | 'flowChartSort' | 'flowChartExtract' | 'flowChartMerge' | 'flowChartOnlineStorage' | 'flowChartDelay' | 'flowChartMagneticTape' | 'flowChartMagneticDisk' | 'flowChartMagneticDrum' | 'flowChartDisplay' | 'wedgeRectCallout' | 'wedgeRoundRectCallout' | 'wedgeEllipseCallout' | 'cloudCallout' | 'borderCallout1' | 'borderCallout2' | 'borderCallout3' | 'accentCallout1' | 'accentCallout2' | 'accentCallout3' | 'callout1' | 'callout2' | 'callout3' | 'accentBorderCallout1' | 'accentBorderCallout2' | 'accentBorderCallout3' | 'actionButtonBlank' | 'actionButtonHome' | 'actionButtonHelp' | 'actionButtonInformation' | 'actionButtonBackPrevious' | 'actionButtonForwardNext' | 'actionButtonBeginning' | 'actionButtonEnd' | 'actionButtonReturn' | 'actionButtonDocument' | 'actionButtonSound' | 'actionButtonMovie' | 'irregularSeal1' | 'irregularSeal2' | 'frame' | 'halfFrame' | 'corner' | 'diagStripe' | 'chord' | 'arc' | 'bracketPair' | 'bracePair' | 'leftBracket' | 'rightBracket' | 'leftBrace' | 'rightBrace' | 'can' | 'cube' | 'bevel' | 'donut' | 'noSmoking' | 'blockArc' | 'foldedCorner' | 'smileyFace' | 'heart' | 'lightningBolt' | 'sun' | 'moon' | 'cloud' | 'snip1Rect' | 'snip2SameRect' | 'snip2DiagRect' | 'snipRoundRect' | 'round1Rect' | 'round2SameRect' | 'round2DiagRect' | 'plaque' | 'teardrop' | 'mathPlus' | 'mathMinus' | 'mathMultiply' | 'mathDivide' | 'mathEqual' | 'mathNotEqual' | 'gear6' | 'gear9' | 'funnel' | 'pieWedge' | 'pie' | 'leftCircularArrow' | 'leftRightCircularArrow' | 'swooshArrow' | 'textBox';
/**
 * Shape fill type
 */
interface ShapeFill {
    type: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture';
    /** Solid fill color */
    color?: ColorValue;
    /** Gradient stops for gradient fill */
    gradient?: {
        type: 'linear' | 'radial' | 'rectangular' | 'path';
        angle?: number;
        stops: Array<{
            position: number;
            color: ColorValue;
        }>;
    };
}
/**
 * Shape outline/stroke
 */
interface ShapeOutline {
    /** Line width in EMUs */
    width?: number;
    /** Line color */
    color?: ColorValue;
    /** Line style */
    style?: 'solid' | 'dot' | 'dash' | 'lgDash' | 'dashDot' | 'lgDashDot' | 'lgDashDotDot' | 'sysDot' | 'sysDash' | 'sysDashDot' | 'sysDashDotDot';
    /** Line cap */
    cap?: 'flat' | 'round' | 'square';
    /** Line join */
    join?: 'bevel' | 'miter' | 'round';
    /** Head arrow */
    headEnd?: {
        type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
        width?: 'sm' | 'med' | 'lg';
        length?: 'sm' | 'med' | 'lg';
    };
    /** Tail arrow */
    tailEnd?: {
        type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
        width?: 'sm' | 'med' | 'lg';
        length?: 'sm' | 'med' | 'lg';
    };
}
/**
 * Text body inside a shape
 */
interface ShapeTextBody {
    /** Text direction */
    vertical?: boolean;
    /** Rotation */
    rotation?: number;
    /** Anchor/vertical alignment */
    anchor?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justified';
    /** Anchor center */
    anchorCenter?: boolean;
    /** Auto fit */
    autoFit?: 'none' | 'normal' | 'shape';
    /** Text margins */
    margins?: {
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
    /** Paragraphs inside the shape */
    content: Paragraph[];
}
/**
 * Shape/drawing object (wps:wsp)
 */
interface Shape {
    type: 'shape';
    /** Shape type preset */
    shapeType: ShapeType;
    /** Unique ID */
    id?: string;
    /** Name */
    name?: string;
    /** Size in EMUs */
    size: ImageSize;
    /** Position for floating shapes */
    position?: ImagePosition;
    /** Wrap settings */
    wrap?: ImageWrap;
    /** Fill */
    fill?: ShapeFill;
    /** Outline/stroke */
    outline?: ShapeOutline;
    /** Transform */
    transform?: ImageTransform;
    /** Text content inside the shape */
    textBody?: ShapeTextBody;
    /** Custom geometry points */
    customGeometry?: string;
}

/**
 * Embedded images (`w:drawing` → `pic:pic`): size, wrap, position,
 * transform, padding, crop.
 */

/**
 * Image size specification
 */
interface ImageSize {
    /** Width in EMUs (English Metric Units) */
    width: number;
    /** Height in EMUs */
    height: number;
}
/**
 * Image wrap type for floating images
 */
interface ImageWrap {
    type: WrapType;
    /** Wrap text direction */
    wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
    /** Distance from text */
    distT?: number;
    distB?: number;
    distL?: number;
    distR?: number;
}
/**
 * Position for floating images
 */
interface ImagePosition {
    /** Horizontal positioning */
    horizontal: {
        relativeTo: 'character' | 'column' | 'insideMargin' | 'leftMargin' | 'margin' | 'outsideMargin' | 'page' | 'rightMargin';
        alignment?: 'left' | 'right' | 'center' | 'inside' | 'outside';
        posOffset?: number;
    };
    /** Vertical positioning */
    vertical: {
        relativeTo: 'insideMargin' | 'line' | 'margin' | 'outsideMargin' | 'page' | 'paragraph' | 'topMargin' | 'bottomMargin';
        alignment?: 'top' | 'bottom' | 'center' | 'inside' | 'outside';
        posOffset?: number;
    };
}
/**
 * Image transformation
 */
interface ImageTransform {
    /** Rotation in degrees */
    rotation?: number;
    /** Flip horizontal */
    flipH?: boolean;
    /** Flip vertical */
    flipV?: boolean;
}
/**
 * Image padding/margins
 */
interface ImagePadding {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
}
/**
 * Image crop, expressed as fractions of the source image to trim from each
 * edge. OOXML's `<a:srcRect l="10000" t="0" r="5000" b="0"/>` uses units of
 * 1/100000 (so 10000 → 0.1 → 10% trimmed from the left). We store the
 * normalised fraction so both the renderer and the saver can read it
 * directly without re-parsing units.
 */
interface ImageCrop {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
}
/**
 * Embedded image (`w:drawing` with an inline or anchored picture). Carries
 * the relationship-id pointer to the binary in `word/media/`, its
 * resolved data URL (`src`), display dimensions, optional crop /
 * transform / wrap behaviors, and anchor positioning for floating
 * images.
 *
 * See ECMA-376 §20.4 (DrawingML wordprocessingDrawing).
 */
interface Image {
    type: 'image';
    /** Unique ID */
    id?: string;
    /** Relationship ID for the image data */
    rId: string;
    /** Resolved image data (base64 or blob URL) */
    src?: string;
    /** Image MIME type */
    mimeType?: string;
    /** Original filename */
    filename?: string;
    /** Alt text for accessibility */
    alt?: string;
    /** Title/description */
    title?: string;
    /** Image size */
    size: ImageSize;
    /** Original size before any transforms */
    originalSize?: ImageSize;
    /** Wrap settings */
    wrap: ImageWrap;
    /** Position for floating images */
    position?: ImagePosition;
    /** Image transformations */
    transform?: ImageTransform;
    /** Padding around image */
    padding?: ImagePadding;
    /** Source-image crop (fractional, OOXML `a:srcRect`). */
    crop?: ImageCrop;
    /** Opacity in [0, 1] (OOXML `a:alphaModFix amt`). Undefined = fully opaque. */
    opacity?: number;
    /** Whether this is a decorative image */
    decorative?: boolean;
    /**
     * `wp:anchor layoutInCell` — when true (default), an anchored image inside
     * a table cell is constrained to the cell. When false, the image escapes
     * the cell into the page area. Round-tripped on save.
     */
    layoutInCell?: boolean;
    /**
     * `wp:anchor allowOverlap` — when true (default), anchored objects may
     * overlap; when false, Word repositions them to avoid collisions. We
     * don't currently reposition; we round-trip the flag so saving preserves
     * the author's intent.
     */
    allowOverlap?: boolean;
    /** Hyperlink URL for clickable image */
    hlinkHref?: string;
    /** Image outline/border */
    outline?: ShapeOutline;
    /** Image effects */
    effects?: {
        brightness?: number;
        contrast?: number;
        saturation?: number;
    };
}

/**
 * Run content (`w:r`) and the inline pieces that live inside a run —
 * text, tab, break, symbol, footnote/endnote references, field chars,
 * instruction text, soft/no-break hyphens, drawings, shapes.
 */

/**
 * Plain text run content (`w:t`). `preserveSpace` mirrors the
 * `xml:space="preserve"` attribute and matters for runs that begin or end
 * with whitespace — without it, Word collapses leading/trailing spaces.
 */
interface TextContent {
    type: 'text';
    /** The text string */
    text: string;
    /** Preserve whitespace (xml:space="preserve") */
    preserveSpace?: boolean;
}
/**
 * Tab character
 */
interface TabContent {
    type: 'tab';
}
/**
 * Line break
 */
interface BreakContent {
    type: 'break';
    /** Break type */
    breakType?: 'page' | 'column' | 'textWrapping';
    /** Clear type for text wrapping break */
    clear?: 'none' | 'left' | 'right' | 'all';
}
/**
 * Symbol character (special font character)
 */
interface SymbolContent {
    type: 'symbol';
    /** Font name */
    font: string;
    /** Character code */
    char: string;
}
/**
 * Footnote or endnote reference
 */
interface NoteReferenceContent {
    type: 'footnoteRef' | 'endnoteRef';
    /** Note ID */
    id: number;
}
/**
 * Footnote/endnote auto-number mark (`w:footnoteRef` / `w:endnoteRef`).
 *
 * Distinct from {@link NoteReferenceContent}: that is the *reference* placed in
 * the document body (`w:footnoteReference`), whereas this is the numbering
 * placeholder that lives *inside* the note body — the run carrying it is what
 * Word renders as the note's leading superscript number. Preserving it keeps
 * the note's own number visible on round-trip.
 */
interface NoteRefMarkContent {
    type: 'footnoteRefMark' | 'endnoteRefMark';
}
/**
 * Footnote/endnote separator mark (`w:separator` / `w:continuationSeparator`).
 *
 * These appear inside the special separator notes (`w:type="separator"` and
 * `w:type="continuationSeparator"`) and draw the horizontal rule Word places
 * between the body and its notes. They carry no content; Word rejects a notes
 * part whose separator notes have lost these markers, so they must round-trip.
 */
interface SeparatorContent {
    type: 'separator' | 'continuationSeparator';
}
/**
 * Field character (begin/separate/end)
 */
interface FieldCharContent {
    type: 'fieldChar';
    /** Field character type */
    charType: 'begin' | 'separate' | 'end';
    /** Field is locked */
    fldLock?: boolean;
    /** Field is dirty (needs update) */
    dirty?: boolean;
}
/**
 * Field instruction text
 */
interface InstrTextContent {
    type: 'instrText';
    /** Field instruction */
    text: string;
}
/**
 * Soft hyphen
 */
interface SoftHyphenContent {
    type: 'softHyphen';
}
/**
 * Non-breaking hyphen
 */
interface NoBreakHyphenContent {
    type: 'noBreakHyphen';
}
/**
 * Drawing/image reference
 */
interface DrawingContent {
    type: 'drawing';
    /** Image data */
    image: Image;
}
/**
 * Shape reference
 */
interface ShapeContent {
    type: 'shape';
    /** Shape data */
    shape: Shape;
}
/**
 * All possible run content types
 */
type RunContent = TextContent | TabContent | BreakContent | SymbolContent | NoteReferenceContent | NoteRefMarkContent | SeparatorContent | FieldCharContent | InstrTextContent | SoftHyphenContent | NoBreakHyphenContent | DrawingContent | ShapeContent;
/**
 * A run (`w:r`) — a contiguous span of inline content sharing one set of
 * character properties (bold, italic, font, color, etc.). Runs are the
 * atomic unit of character formatting; toggling bold on a selection that
 * spans different formatting creates new runs.
 *
 * See ECMA-376 §17.3.2.
 *
 * @example
 * ```ts
 * const run: Run = {
 *   type: 'run',
 *   formatting: { bold: true },
 *   content: [{ type: 'text', text: 'Hello' }],
 * };
 * ```
 */
interface Run {
    type: 'run';
    /** Text formatting properties */
    formatting?: TextFormatting;
    /** Run-level tracked property changes (w:rPrChange) */
    propertyChanges?: RunPropertyChange[];
    /** Run content (text, tabs, breaks, etc.) */
    content: RunContent[];
}

/**
 * Styles, Theme, Font Table, Relationships & Media Types
 *
 * Types for document-level definitions that don't form the content tree.
 */

/**
 * Style type
 */
type StyleType = 'paragraph' | 'character' | 'numbering' | 'table';
/**
 * Style definition from `styles.xml` — a named, reusable bundle of
 * paragraph and/or character formatting. Word's "Heading 1", "Normal",
 * "Title", and "List Bullet" are styles; user-defined styles look the
 * same. `basedOn` chains styles for inheritance; `link` pairs a paragraph
 * style with a matching character style.
 *
 * See ECMA-376 §17.7.4.
 */
interface Style {
    /** Style ID */
    styleId: string;
    /** Style type */
    type: StyleType;
    /** Display name */
    name?: string;
    /** Based on style ID */
    basedOn?: string;
    /** Next style after Enter (for paragraph styles) */
    next?: string;
    /** Linked style (paragraph/character pair) */
    link?: string;
    /** UI sort priority */
    uiPriority?: number;
    /** Hidden from UI */
    hidden?: boolean;
    /** Semi-hidden from UI */
    semiHidden?: boolean;
    /** Unhide when used */
    unhideWhenUsed?: boolean;
    /** Quick format in gallery */
    qFormat?: boolean;
    /** Is default style */
    default?: boolean;
    /** Personal style (custom) */
    personal?: boolean;
    /** Paragraph properties (for paragraph/table styles) */
    pPr?: ParagraphFormatting;
    /** Run properties */
    rPr?: TextFormatting;
    /** Table properties (for table styles) */
    tblPr?: TableFormatting;
    /** Table row properties */
    trPr?: TableRowFormatting;
    /** Table cell properties */
    tcPr?: TableCellFormatting;
    /** Conditional table style parts */
    tblStylePr?: Array<{
        type: 'band1Horz' | 'band1Vert' | 'band2Horz' | 'band2Vert' | 'firstCol' | 'firstRow' | 'lastCol' | 'lastRow' | 'neCell' | 'nwCell' | 'seCell' | 'swCell';
        pPr?: ParagraphFormatting;
        rPr?: TextFormatting;
        tblPr?: TableFormatting;
        trPr?: TableRowFormatting;
        tcPr?: TableCellFormatting;
    }>;
}
/**
 * Document defaults (w:docDefaults)
 */
interface DocDefaults {
    /** Default run properties */
    rPr?: TextFormatting;
    /** Default paragraph properties */
    pPr?: ParagraphFormatting;
}
/**
 * Style definitions from styles.xml
 */
interface StyleDefinitions {
    /** Document defaults */
    docDefaults?: DocDefaults;
    /** Latent styles */
    latentStyles?: {
        defLockedState?: boolean;
        defUIPriority?: number;
        defSemiHidden?: boolean;
        defUnhideWhenUsed?: boolean;
        defQFormat?: boolean;
        count?: number;
    };
    /** Style definitions */
    styles: Style[];
}
/**
 * Theme color scheme (a:clrScheme)
 */
interface ThemeColorScheme {
    /** Dark 1 color (usually black) */
    dk1?: string;
    /** Light 1 color (usually white) */
    lt1?: string;
    /** Dark 2 color */
    dk2?: string;
    /** Light 2 color */
    lt2?: string;
    /** Accent colors 1-6 */
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
    /** Hyperlink color */
    hlink?: string;
    /** Followed hyperlink color */
    folHlink?: string;
}
/**
 * Theme font (with script variants)
 */
interface ThemeFont {
    /** Latin font */
    latin?: string;
    /** East Asian font */
    ea?: string;
    /** Complex script font */
    cs?: string;
    /** Script-specific fonts */
    fonts?: Record<string, string>;
}
/**
 * Theme font scheme (a:fontScheme)
 */
interface ThemeFontScheme {
    /** Major font (headings) */
    majorFont?: ThemeFont;
    /** Minor font (body text) */
    minorFont?: ThemeFont;
}
/**
 * Theme (from theme1.xml)
 */
interface Theme {
    /** Theme name */
    name?: string;
    /** Color scheme */
    colorScheme?: ThemeColorScheme;
    /** Font scheme */
    fontScheme?: ThemeFontScheme;
    /** Format scheme (fills, lines, effects) - simplified */
    formatScheme?: {
        name?: string;
    };
}
/**
 * Font info from fontTable.xml
 */
interface FontInfo {
    /** Font name */
    name: string;
    /** Alternate names */
    altName?: string;
    /** Panose-1 classification */
    panose1?: string;
    /** Character set */
    charset?: string;
    /** Font family type */
    family?: 'decorative' | 'modern' | 'roman' | 'script' | 'swiss' | 'auto';
    /** Pitch (fixed or variable) */
    pitch?: 'default' | 'fixed' | 'variable';
    /** Signature */
    sig?: {
        usb0?: string;
        usb1?: string;
        usb2?: string;
        usb3?: string;
        csb0?: string;
        csb1?: string;
    };
    /** Embedded regular face (`w:embedRegular`) */
    embedRegular?: FontEmbed;
    /** Embedded bold face (`w:embedBold`) */
    embedBold?: FontEmbed;
    /** Embedded italic face (`w:embedItalic`) */
    embedItalic?: FontEmbed;
    /** Embedded bold-italic face (`w:embedBoldItalic`) */
    embedBoldItalic?: FontEmbed;
}
/**
 * A single embedded font face referenced from `fontTable.xml`. The font binary
 * lives at the relationship target (an obfuscated `.odttf` under `word/fonts/`)
 * and is descrambled with `fontKey`. See `utils/fontDeobfuscation.ts`.
 */
interface FontEmbed {
    /** Relationship ID (`r:id`) resolved via `word/_rels/fontTable.xml.rels`. */
    relId: string;
    /** Obfuscation key GUID (`w:fontKey`), e.g. `{XXXXXXXX-....}`. */
    fontKey?: string;
    /** Whether the embedded face is subsetted (`w:subsetted`). */
    subsetted?: boolean;
}
/**
 * Font table from fontTable.xml
 */
interface FontTable {
    fonts: FontInfo[];
}
/**
 * Relationship type
 */
type RelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart' | 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData' | string;
/**
 * Relationship entry
 */
interface Relationship {
    /** Relationship ID (e.g., "rId1") */
    id: string;
    /** Relationship type URI */
    type: RelationshipType;
    /** Target path or URL */
    target: string;
    /** Target mode */
    targetMode?: 'External' | 'Internal';
}
/**
 * Relationship map (keyed by rId)
 */
type RelationshipMap = Map<string, Relationship>;
/**
 * Media file from word/media/
 */
interface MediaFile {
    /** File path in ZIP */
    path: string;
    /** Original filename */
    filename?: string;
    /** MIME type */
    mimeType: string;
    /** Binary data */
    data: ArrayBuffer;
    /** Base64 encoded data for rendering */
    base64?: string;
    /** Data URL for direct use in src attributes */
    dataUrl?: string;
}

/**
 * settings.xml parser
 *
 * Extracts document-wide settings the layout pipeline needs at render time.
 * We only read what's currently consumed; most of settings.xml (compatibility
 * flags, view state, autoformat) is irrelevant to layout.
 */
/** Document-wide settings parsed from `word/settings.xml`. */
interface DocumentSettings {
    /**
     * `w:defaultTabStop` (§17.6.13) — interval in twips between default tab
     * stops applied when a paragraph has no custom `w:tabs`. Word's default
     * if unspecified is 720 twips (0.5 inch).
     */
    defaultTabStop: number;
    /**
     * `w:defaultTableStyle` (§17.15.1.44) — the styleId applied to every
     * newly created table in this document. Distinct from the type-default
     * table style (`w:default="1"`), which is what tables inherit from when
     * they carry no `w:tblStyle`. Absent in most documents; a template author
     * sets it so inserted tables pick up the template's table look.
     */
    defaultTableStyle?: string;
    /**
     * `w:themeFontLang` (§17.15.1.88) — the language Word uses to pick the
     * concrete typeface for the *EastAsian* (`w:eastAsia`) and *complex-script*
     * (`w:bidi`) theme font slots when the theme's `<a:ea>`/`<a:cs>` typeface is
     * empty (the common case for Office's default theme). The script-specific
     * `<a:font script="…">` entries are selected by this language. Without it a
     * CJK document whose runs reference `minorEastAsia` resolves to an empty
     * font, which breaks both measurement and painting (text overflows the
     * right margin).
     */
    themeFontLang?: {
        eastAsia?: string;
        bidi?: string;
    };
}

/**
 * Comprehensive TypeScript types for full DOCX document representation
 *
 * This barrel file re-exports all types from the split modules.
 * Existing imports from './types/document' continue to work unchanged.
 *
 * Module structure:
 * - colors.ts      — Color primitives, borders, shading
 * - formatting.ts  — Text, paragraph, and table formatting properties
 * - lists.ts       — Numbering and list definitions
 * - content.ts     — Content model (runs, images, shapes, tables, paragraphs, sections)
 * - styles.ts      — Styles, theme, fonts, relationships, media
 * @packageDocumentation
 * @public
 */

/**
 * Complete DOCX package structure
 */
interface DocxPackage {
    /** Document body */
    document: DocumentBody;
    /** Style definitions */
    styles?: StyleDefinitions;
    /** Theme */
    theme?: Theme;
    /** Numbering definitions */
    numbering?: NumberingDefinitions;
    /** Document-wide settings from `word/settings.xml` */
    settings?: DocumentSettings;
    /** Font table */
    fontTable?: FontTable;
    /** Footnotes (normal notes only — separators live in `footnoteSeparators`) */
    footnotes?: Footnote[];
    /** Endnotes (normal notes only — separators live in `endnoteSeparators`) */
    endnotes?: Endnote[];
    /**
     * Separator footnotes (`w:type="separator"` / `"continuationSeparator"` /
     * `"continuationNotice"`) kept out of `footnotes` so rendering/layout only
     * sees real notes. Retained for round-trip: Word rejects a footnotes part
     * whose separator notes are missing, so the serializer re-emits these ahead
     * of the normal notes.
     */
    footnoteSeparators?: Footnote[];
    /** Separator endnotes — see `footnoteSeparators`. */
    endnoteSeparators?: Endnote[];
    /** Headers by relationship ID */
    headers?: Map<string, HeaderFooter>;
    /** Footers by relationship ID */
    footers?: Map<string, HeaderFooter>;
    /** Document relationships */
    relationships?: RelationshipMap;
    /** Media files */
    media?: Map<string, MediaFile>;
    /** Document properties */
    properties?: {
        title?: string;
        subject?: string;
        creator?: string;
        keywords?: string;
        description?: string;
        lastModifiedBy?: string;
        revision?: number;
        created?: Date;
        modified?: Date;
    };
}
/**
 * Top-level parsed DOCX document — the result of `parseDocx(buffer)`.
 *
 * Wraps the unzipped DOCX package (`document.xml`, `styles.xml`, etc.),
 * the original buffer for round-trip saves, and any template variables /
 * parse warnings detected during ingestion.
 *
 * @example
 * ```ts
 * import { parseDocx } from '@eigenpal/docx-editor-core/headless';
 * const doc = await parseDocx(buffer);
 * console.log(doc.package.document.content.length);
 * ```
 */
interface Document {
    /** Parsed DOCX package — body, styles, numbering, theme, media, headers/footers. */
    package: DocxPackage;
    /** Original DOCX buffer. Kept for round-trip saves that preserve untouched parts. */
    originalBuffer?: ArrayBuffer;
    /** Detected docxtemplater variables (e.g. `{name}`, `{address}`). Populated when the document is recognized as a template. */
    templateVariables?: string[];
    /** Non-fatal parser diagnostics — malformed parts, unsupported features, fallbacks. */
    warnings?: string[];
}

interface HeadingBlock {
    type: 'heading';
    index: number;
    /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
    paraId?: string;
    level: number;
    text: string;
}
interface ParagraphBlock {
    type: 'paragraph';
    index: number;
    /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
    paraId?: string;
    text: string;
}
interface TableBlock {
    type: 'table';
    index: number;
    rows: string[][];
    /** Per-cell paraIds, parallel to `rows`. cells[r][c] is the first paraId in that cell. */
    cellParaIds?: (string | undefined)[][];
}
interface ListItemBlock {
    type: 'list-item';
    index: number;
    /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
    paraId?: string;
    text: string;
    listLevel: number;
    listType: 'bullet' | 'number';
}
type ContentBlock = HeadingBlock | ParagraphBlock | TableBlock | ListItemBlock;
interface GetContentOptions {
    fromIndex?: number;
    toIndex?: number;
    /** Annotate tracked changes inline. Default: true */
    includeTrackedChanges?: boolean;
    /** Annotate comments inline. Default: true */
    includeCommentAnchors?: boolean;
}
interface ReviewChange {
    id: number;
    type: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo';
    author: string;
    date: string | null;
    text: string;
    context: string;
    /**
     * Index of the containing paragraph. For body changes this is the
     * document-wide paragraph index; for note changes it is the paragraph index
     * *within that note* (note bodies have their own numbering), so pair it with
     * `noteId` / `noteType` rather than reading it as a body index.
     */
    paragraphIndex: number;
    /**
     * Set when the change lives inside a footnote or endnote. The raw `id` is not
     * namespaced across parts, so pair it with `noteId` / `noteType` to identify
     * the change, and pass the whole `ReviewChange` back to `acceptChange` /
     * `rejectChange` to resolve it inside that note.
     */
    noteId?: number;
    /** Which note store the change came from. Absent for body changes. */
    noteType?: 'footnote' | 'endnote';
}
interface ReviewCommentReply {
    id: number;
    author: string;
    date: string | null;
    text: string;
}
interface ReviewComment {
    id: number;
    author: string;
    date: string | null;
    text: string;
    anchoredText: string;
    paragraphIndex: number;
    replies: ReviewCommentReply[];
    done: boolean;
}
interface ChangeFilter {
    author?: string;
    type?: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo';
    /** Also report tracked changes inside footnote bodies. Default: false. */
    includeFootnotes?: boolean;
    /** Also report tracked changes inside endnote bodies. Default: false. */
    includeEndnotes?: boolean;
}
/**
 * Options for `acceptAll` / `rejectAll`. Opt in to also resolve tracked changes
 * inside footnote/endnote bodies; mirrors the discovery flags on {@link ChangeFilter}.
 */
type AcceptChangesOptions = Pick<ChangeFilter, 'includeFootnotes' | 'includeEndnotes'>;
interface CommentFilter {
    author?: string;
    done?: boolean;
}
/**
 * Live-editor (bridge) action options — anchored by Word `w14:paraId`.
 * Stable across edits; the agent gets paraIds from `read_document` / `find_text`.
 */
interface AddCommentByParaIdOptions {
    paraId: string;
    text: string;
    author?: string;
    /** Optional: anchor to a specific phrase within the paragraph (must be unique). */
    search?: string;
}
/**
 * Headless / DocxReviewer action options — anchored by ordinal paragraph index.
 * Used by the static-document review pipeline.
 */
interface AddCommentOptions {
    paragraphIndex: number;
    text: string;
    author?: string;
    /** Optional: anchor to specific text. Omit to anchor whole paragraph. */
    search?: string;
}
interface ReplyOptions {
    text: string;
    author?: string;
}
/**
 * Live-editor change. Pass `replaceWith: ''` to delete; pass `search: ''` to
 * insert at end of paragraph; pass both non-empty for a replacement.
 */
interface ProposeChangeOptions {
    paraId: string;
    search: string;
    replaceWith: string;
    author?: string;
}
interface ProposeReplacementOptions {
    paragraphIndex: number;
    search: string;
    replaceWith: string;
    author?: string;
}
interface ProposeInsertionOptions {
    paragraphIndex: number;
    insertText: string;
    author?: string;
    position?: 'before' | 'after';
    search?: string;
}
interface ProposeDeletionOptions {
    paragraphIndex: number;
    search: string;
    author?: string;
}
/** Per-match handle returned by `findText` — pass `paraId` + `match` back as `search`. */
interface FoundMatch {
    paraId: string;
    match: string;
    before: string;
    after: string;
}
/** @public */
interface SelectionInfo {
    paraId: string | null;
    selectedText: string;
    paragraphText: string;
    before: string;
    after: string;
}
/**
 * Character formatting marks the agent can apply.
 *
 * Mirrors Word JS API `Range.font.*`. A `false` value clears that mark in the
 * range; a missing key leaves it untouched. `color.themeColor` follows ECMA-376
 * theme color values (e.g. `'accent1'`, `'text1'`) and resolves at render time.
 */
interface CharacterFormatting {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | {
        style?: string;
    };
    strike?: boolean;
    color?: {
        rgb?: string;
        themeColor?: string;
    };
    highlight?: string;
    fontSize?: number;
    fontFamily?: {
        ascii?: string;
        hAnsi?: string;
    };
}
/**
 * Apply character formatting to a range. `paraId` is required; if `search`
 * is provided, the formatting only applies to that phrase within the
 * paragraph (must match exactly once). Otherwise it applies to the whole
 * paragraph's text.
 */
interface ApplyFormattingOptions {
    paraId: string;
    search?: string;
    marks: CharacterFormatting;
}
/**
 * Apply a paragraph style by `styleId` (e.g. `'Heading1'`, `'Title'`,
 * `'Quote'`). The styleId must exist in the document's style definitions
 * — unknown ids are no-ops.
 */
interface SetParagraphStyleOptions {
    paraId: string;
    styleId: string;
}
/** Kind of break `insertBreak` inserts after a paragraph. */
type BreakType = 'page' | 'sectionNextPage' | 'sectionContinuous';
/**
 * Insert a page or section break after the paragraph identified by `paraId`.
 * `page` adds a page break; `sectionNextPage` starts a new section on a new
 * page; `sectionContinuous` starts a new section on the same page.
 */
interface InsertBreakOptions {
    paraId: string;
    type: BreakType;
}
/** A single paragraph anchored on a page (returned by `getPage` / `getPages`). */
interface PageParagraph {
    paraId: string;
    text: string;
    /** True for headings, list items, and other styled blocks. */
    styleId?: string;
}
/** What the agent sees when reading one or more pages. */
interface PageContent {
    /** 1-indexed page number. */
    pageNumber: number;
    /** Plain text of the page, formatted as `[paraId] text` lines. */
    text: string;
    /** Paragraphs on the page, in document order. */
    paragraphs: PageParagraph[];
}
/**
 * Snapshot of what the user is looking at — pass this to your agent's system
 * prompt so it knows the current selection / page without an extra
 * `read_selection` round-trip.
 *
 * @public
 */
interface AgentContextSnapshot {
    /** User's current selection or cursor (null if editor isn't focused). */
    selection: SelectionInfo | null;
    /** 1-indexed page the cursor / selection is on. 0 if unknown. */
    currentPage: number;
    /** Total number of pages currently rendered. */
    totalPages: number;
}
interface BatchReviewOptions {
    accept?: number[];
    reject?: number[];
    comments?: AddCommentOptions[];
    replies?: (ReplyOptions & {
        commentId: number;
    })[];
    proposals?: ProposeReplacementOptions[];
}
interface BatchError {
    operation: string;
    id?: number;
    search?: string;
    error: string;
}
interface BatchResult {
    accepted: number;
    rejected: number;
    commentsAdded: number;
    repliesAdded: number;
    proposalsAdded: number;
    errors: BatchError[];
}

/**
 * Headless DOCX reviewer — parse a file, read/comment/track changes
 * against the document model, write the modified DOCX back out. No DOM,
 * no editor instance. Pair with `createReviewerBridge()` to drive the
 * built-in agent tools against a file on disk.
 *
 * @example
 * ```ts
 * const reviewer = await DocxReviewer.fromBuffer(buffer, 'AI Reviewer');
 * reviewer.addComment(5, 'Fix this paragraph.');
 * reviewer.replace(5, '$50k', '$500k');
 * const output = await reviewer.toBuffer();
 * ```
 *
 * @public
 */
declare class DocxReviewer {
    private doc;
    /** Default author for comments and tracked changes. Set once, used everywhere. */
    readonly author: string;
    /**
     * Create a reviewer from a parsed Document.
     * @param document - Parsed Document from the core package
     * @param author - Default author name for comments and changes. (default: 'AI')
     * @param originalBuffer - Original DOCX buffer, needed for toBuffer()
     */
    constructor(document: Document, author?: string, originalBuffer?: ArrayBuffer);
    /**
     * Create a reviewer from a DOCX file buffer.
     * @param buffer - ArrayBuffer of the DOCX file
     * @param author - Default author name for comments and changes. (default: 'AI')
     */
    static fromBuffer(buffer: ArrayBuffer, author?: string): Promise<DocxReviewer>;
    private get body();
    private resolveAuthor;
    /** Get document content as structured blocks (headings, paragraphs, tables, lists). */
    getContent(options?: GetContentOptions): ContentBlock[];
    /**
     * Get document content as plain text for LLM prompts.
     * Each paragraph is prefixed with its index: `[0] text`, `[1] text`, etc.
     * Table cells include position: `[5] (table, row 1, col 2) cell text`.
     * Avoids JSON quote-escaping issues — LLMs can copy text verbatim.
     */
    getContentAsText(options?: GetContentOptions): string;
    /**
     * Get all tracked changes in the document. Pass `includeFootnotes` /
     * `includeEndnotes` in the filter to also report changes inside note bodies
     * (each such change carries `noteId` / `noteType`).
     *
     * The reported `id` is the raw `w:id`, which is unique only within its part
     * (document.xml / footnotes.xml / endnotes.xml) — it is NOT namespaced across
     * parts, so the same `id` can appear on a body change and a note change. Use
     * `noteType` / `noteId` to disambiguate.
     *
     * A returned change with `noteId`/`noteType` set can be accepted or rejected
     * by passing the whole {@link ReviewChange} back to {@link acceptChange} /
     * {@link rejectChange} (which resolves it inside its footnote/endnote), or in
     * bulk via {@link acceptAll} / {@link rejectAll} with the matching `include*`
     * option; the result persists on {@link toBuffer}.
     */
    getChanges(filter?: ChangeFilter): ReviewChange[];
    /** Get all comments with their replies. */
    getComments(filter?: CommentFilter): ReviewComment[];
    /**
     * Add a comment on a paragraph.
     * @param paragraphIndex - Index of the paragraph to comment on
     * @param text - Comment text
     * @returns The new comment ID
     */
    addComment(paragraphIndex: number, text: string): number;
    /**
     * Add a comment with full options (custom author, anchored to specific text).
     * @param options - Comment options
     * @returns The new comment ID
     */
    addComment(options: AddCommentOptions): number;
    /**
     * Reply to an existing comment.
     * @param commentId - ID of the comment to reply to
     * @param text - Reply text
     * @returns The new reply comment ID
     */
    replyTo(commentId: number, text: string): number;
    /** Reply to an existing comment with full options. */
    replyTo(commentId: number, options: ReplyOptions): number;
    /**
     * Remove a comment by ID. Removing a top-level comment also removes its
     * replies and the anchored range markers. Removing a reply only removes
     * that reply.
     * @param commentId - ID of the comment to remove
     */
    removeComment(commentId: number): void;
    /**
     * Replace text in a paragraph. Creates a tracked change (deletion + insertion).
     * @param paragraphIndex - Index of the paragraph
     * @param search - Short phrase to find within the paragraph
     * @param replaceWith - Replacement text
     */
    replace(paragraphIndex: number, search: string, replaceWith: string): void;
    /** Replace text with full options. */
    replace(options: ProposeReplacementOptions): void;
    /** @deprecated Use replace() instead. */
    proposeReplacement(options: ProposeReplacementOptions): void;
    /** Insert text as a tracked change. */
    proposeInsertion(options: ProposeInsertionOptions): void;
    /** Delete text as a tracked change. */
    proposeDeletion(options: ProposeDeletionOptions): void;
    /**
     * Accept a tracked change. Pass a revision id to accept a change in the
     * document body, or pass a {@link ReviewChange} from {@link getChanges} to
     * accept it wherever it lives — a change carrying `noteId`/`noteType` is
     * resolved inside that footnote/endnote and persists on {@link toBuffer}.
     *
     * A bare numeric id targets the body only: a `w:id` is unique only within its
     * part, so the same id can appear on a body change and a note change. To
     * resolve a note change pass the whole {@link ReviewChange} (its
     * `noteId`/`noteType` locate it); a bare id resolves to the body change, if any.
     */
    acceptChange(target: number | ReviewChange): void;
    /** Reject a tracked change. See {@link acceptChange} for body-vs-note targeting. */
    rejectChange(target: number | ReviewChange): void;
    /**
     * Accept all tracked changes in the body. Pass `{ includeFootnotes,
     * includeEndnotes }` to also accept changes inside note bodies. Returns count.
     */
    acceptAll(opts?: AcceptChangesOptions): number;
    /** Reject all tracked changes. See {@link acceptAll} for the note opt-in. */
    rejectAll(opts?: AcceptChangesOptions): number;
    /** The package's note stores, passed to change ops so note changes resolve. */
    private changeNotes;
    /**
     * Apply multiple review operations in one call.
     * Uses the reviewer's default author. Individual failures are collected, not thrown.
     */
    applyReview(ops: BatchReviewOptions): BatchResult;
    /** Get the modified Document model. */
    toDocument(): Document;
    /** Serialize back to a DOCX buffer. Requires the original buffer. */
    toBuffer(): Promise<ArrayBuffer>;
}

/**
 * Reviewer bridge — wraps a `DocxReviewer` (static document) in the same
 * `EditorBridge` interface the live editor exposes. Lets the same MCP server
 * / agent tools operate on a parsed-from-disk DOCX without a running editor.
 *
 * Trade-offs vs. the live bridge:
 *  - `getSelection()` always returns `null` (no user, no selection).
 *  - `scrollTo()` is a no-op that returns `true` (the doc isn't being viewed).
 *  - `onSelectionChange` listeners never fire (returned unsubscribers are no-ops).
 *  - `onContentChange` fires after every successful mutation through this bridge,
 *    so MCP clients still get notifications when the agent is the only writer.
 *  - paraId resolution maps to `paragraphIndex` by walking the document body
 *    once. None of the reviewer's mutators (`addComment`, `proposeChange`,
 *    `replyTo`, `resolveComment`) shift top-level indices, so the map is
 *    cached for the bridge's lifetime.
 *
 * After the agent finishes mutating, call `reviewer.toBuffer()` to serialize
 * back to DOCX. The bridge does NOT do that automatically — the host decides
 * when to flush.
 */

/**
 * Create an EditorBridge backed by a DocxReviewer. The agent (or MCP client)
 * can read, comment, propose changes, etc., against a parsed DOCX file on
 * disk. Call `reviewer.toBuffer()` afterwards to get the modified DOCX.
 *
 * @param reviewer - A DocxReviewer instance. The bridge mutates it in place.
 *
 * @public
 */
declare function createReviewerBridge(reviewer: DocxReviewer): EditorBridge;

/**
 * Option shapes for `scrollToParaId(paraId, { highlight })`.
 *
 * Kept in a DOM-free module (no imports, no browser globals) so non-browser
 * consumers like `@eigenpal/docx-editor-agents` can type-import them without
 * pulling `paragraphFlash`'s DOM code into their type-check surface.
 */
/**
 * Customization for the transient paragraph flash applied by
 * `scrollToParaId(paraId, { highlight })`.
 *
 * @public
 */
interface ParagraphHighlightOptions {
    /** CSS color used for the transient paragraph flash. Defaults to yellow. */
    color?: string;
    /** How long the flash remains visible before it is removed. Defaults to 1200ms. */
    durationMs?: number;
}
/**
 * Optional reveal behavior for `scrollToParaId`.
 *
 * @public
 */
interface ScrollToParaIdOptions {
    /** Flash rendered paragraph fragments after scrolling to the paragraph. */
    highlight?: ParagraphHighlightOptions;
}

/**
 * Agent-bridge contract every editor adapter (React, Vue, future) MUST satisfy.
 * Versioning: additions are coordinated minor bumps across the fixed group;
 * signature changes / removals are major. See
 * `openspec/changes/vue-editor-robust-implementation/design.md` Decision 18.
 */
interface EditorRefLike {
    getDocument(): unknown | null;
    getEditorRef(): {
        getDocument(): unknown | null;
    } | null;
    addComment(options: {
        paraId: string;
        text: string;
        author: string;
        search?: string;
    }): number | null;
    replyToComment(commentId: number, text: string, author: string): number | null;
    resolveComment(commentId: number): void;
    proposeChange(options: {
        paraId: string;
        search: string;
        replaceWith: string;
        author: string;
    }): boolean;
    scrollToParaId(paraId: string, options?: ScrollToParaIdOptions): boolean;
    findInDocument(query: string, options?: {
        caseSensitive?: boolean;
        limit?: number;
    }): FoundMatch[];
    getSelectionInfo(): SelectionInfo | null;
    getComments(): Array<{
        id: number;
        author: string;
        date?: string;
        parentId?: number;
        content: unknown[];
        done?: boolean;
    }>;
    /** Apply character formatting to a paragraph or sub-range. Returns false on missing paraId / ambiguous search. */
    applyFormatting(options: {
        paraId: string;
        search?: string;
        marks: CharacterFormatting;
    }): boolean;
    /** Apply a paragraph style by styleId. Returns false if paraId is unknown. */
    setParagraphStyle(options: {
        paraId: string;
        styleId: string;
    }): boolean;
    /** Insert a page or section break after the paragraph. Returns false if paraId is unknown. */
    insertBreak(options: {
        paraId: string;
        type: 'page' | 'sectionNextPage' | 'sectionContinuous';
    }): boolean;
    /** Read a single page's paragraphs (1-indexed). Returns null if the page does not exist. */
    getPageContent(pageNumber: number): PageContent | null;
    /** Total number of pages currently rendered. */
    getTotalPages(): number;
    /** 1-indexed page the user's cursor / selection is on. 0 if unknown. */
    getCurrentPage(): number;
    onContentChange(listener: (doc: unknown) => void): () => void;
    onSelectionChange(listener: (selection: unknown) => void): () => void;
}
/**
 * High-level agent surface over a live editor (or a headless reviewer).
 * Every built-in tool calls into this contract — implement it once and
 * the agent toolkit works against your editor.
 *
 * @public
 */
interface EditorBridge {
    /** Get document content as paraId-tagged text lines for LLM prompts. */
    getContentAsText(options?: GetContentOptions): string;
    /** Get document content as structured blocks (each paragraph carries its `paraId`). */
    getContent(options?: GetContentOptions): ContentBlock[];
    /** Get all comments in the document. */
    getComments(filter?: CommentFilter): ReviewComment[];
    /** Get all tracked changes in the document. */
    getChanges(filter?: ChangeFilter): ReviewChange[];
    /** Locate text in the document. Returns one handle per matching paragraph. */
    findText(query: string, options?: {
        caseSensitive?: boolean;
        limit?: number;
    }): FoundMatch[];
    /** Read the user's current cursor / selection. */
    getSelection(): SelectionInfo | null;
    /** Add a comment, anchored to a paragraph by paraId. */
    addComment(options: AddCommentByParaIdOptions): number | null;
    /** Reply to an existing comment. Returns the reply ID or null. */
    replyTo(commentId: number, options: ReplyOptions): number | null;
    /** Resolve a comment (mark as done). */
    resolveComment(commentId: number): void;
    /** Suggest a tracked change. `replaceWith=''` deletes; `search=''` inserts at paragraph end. */
    proposeChange(options: ProposeChangeOptions): boolean;
    /**
     * Apply character formatting (bold / italic / color / size / font / etc.)
     * to a paragraph, or to a unique phrase within it. This is a direct edit —
     * not a tracked change.
     */
    applyFormatting(options: ApplyFormattingOptions): boolean;
    /**
     * Apply a paragraph style by styleId (e.g. `'Heading1'`, `'Quote'`).
     * Direct edit, not a tracked change.
     */
    setParagraphStyle(options: SetParagraphStyleOptions): boolean;
    /**
     * Insert a page or section break after a paragraph, by paraId. `page` adds a
     * page break; `sectionNextPage` / `sectionContinuous` start a new section on
     * a new page / the same page. Direct edit, not a tracked change.
     */
    insertBreak(options: InsertBreakOptions): boolean;
    /** Read a single page (1-indexed). Returns null if the page does not exist. */
    getPage(pageNumber: number): PageContent | null;
    /** Read a range of pages (1-indexed, inclusive). Out-of-range pages are skipped. */
    getPages(options: {
        from: number;
        to: number;
    }): PageContent[];
    /** Total number of pages currently rendered in the editor. */
    getTotalPages(): number;
    /** 1-indexed page the user's cursor / selection is on. 0 if unknown. */
    getCurrentPage(): number;
    /** Scroll the editor to a paragraph by paraId, optionally flashing it. */
    scrollTo(paraId: string, options?: ScrollToParaIdOptions): boolean;
    /** Subscribe to document content changes. Returns an unsubscribe function. */
    onContentChange(listener: (event: ContentChangeEvent) => void): () => void;
    /** Subscribe to selection changes (cursor moves / selection changes). Returns an unsubscribe function. */
    onSelectionChange(listener: (event: SelectionChangeEvent) => void): () => void;
}
/** Event payload for `onContentChange`. */
interface ContentChangeEvent {
    /** Total comments in the document after the change. */
    commentCount: number;
    /** Total tracked changes after the change. */
    changeCount: number;
    /** Snapshot of all current comments. */
    comments: ReviewComment[];
    /** Snapshot of all current tracked changes. */
    changes: ReviewChange[];
}
/** Event payload for `onSelectionChange`. */
type SelectionChangeEvent = SelectionInfo | null;
/**
 * Create an EditorBridge from a DocxEditorRef.
 *
 * @param editorRef - A DocxEditorRef (or anything matching EditorRefLike)
 * @param author - Default author name for comments and changes. (default: 'AI')
 */
declare function createEditorBridge(editorRef: EditorRefLike, author?: string): EditorBridge;

/**
 * Agent tool type definitions.
 */

/**
 * Definition of an agent tool — name, JSON-schema input, handler.
 * Use this to build custom tools alongside the built-in `agentTools`.
 *
 * @public
 */
interface AgentToolDefinition<TInput = Record<string, unknown>> {
    /** Tool name (used in tool_use blocks) */
    name: string;
    /**
     * Friendly UI label for the tool. Shown in the agent panel's timeline
     * (e.g. "Reading document"). Falls back to a sentence-case version of
     * `name` if omitted, so consumer-defined tools render readably without
     * specifying this.
     */
    displayName?: string;
    /** Human-readable description for the LLM */
    description: string;
    /** JSON Schema for the input parameters */
    inputSchema: Record<string, unknown>;
    /** Handler — receives parsed input + bridge, returns result */
    handler: (input: TInput, bridge: EditorBridge) => AgentToolResult;
}
/**
 * Result returned by a tool handler. `success: false` carries an `error`
 * message; `success: true` may carry tool-specific `data`.
 *
 * @public
 */
interface AgentToolResult {
    success: boolean;
    data?: unknown;
    error?: string;
}

/**
 * Agent tool definitions and execution.
 *
 * Tools use OpenAI function-calling format. The pattern mirrors Word's JS API:
 * locate first (`read_document` / `find_text` / `read_selection` return paraId
 * handles), then mutate (`comment` / `suggest_change` / `reply_comment` /
 * `resolve_comment` / `scroll`). paraId anchors are stable across edits.
 */

/**
 * All built-in agent tools — read/write document content, comments, and
 * tracked changes. Use `getToolSchemas()` to feed them to an LLM and
 * `executeToolCall()` to run the handlers against an `EditorBridge`.
 *
 * @public
 */
declare const agentTools: AgentToolDefinition<any>[];
/**
 * Execute a tool call against an EditorBridge.
 * Returns the result (never throws).
 *
 * @public
 */
declare function executeToolCall(toolName: string, input: Record<string, unknown>, bridge: EditorBridge): AgentToolResult;
/**
 * Friendly UI label for a tool — sourced from the registry's `displayName`,
 * falling back to a sentence-case version of the snake_case name. Used by
 * `<AgentTimeline>` and any other UI that lists running / completed tools.
 *
 * @example getToolDisplayName('add_comment') // → 'Adding comment'
 * @example getToolDisplayName('fetch_clause_template') // → 'Fetch clause template'
 *
 * @public
 */
declare function getToolDisplayName(name: string): string;
/**
 * Get tool schemas in OpenAI function-calling format. Works directly
 * with the OpenAI SDK and Anthropic's tools API. For Vercel AI SDK,
 * LangChain, or other agent runtimes, transform this output to that
 * runtime's required shape — see `examples/agent-chat-demo/` for a
 * Vercel AI SDK example. The package stays runtime-agnostic.
 *
 * @public
 */
declare function getToolSchemas(): {
    type: "function";
    function: {
        name: string;
        description: string;
        parameters: Record<string, unknown>;
    };
}[];

export { type AddCommentByParaIdOptions as A, type BatchError as B, type ContentBlock as C, DocxReviewer as D, type EditorBridge as E, type FoundMatch as F, type GetContentOptions as G, type ProposeChangeOptions as P, type ReviewComment as R, type SelectionInfo as S, type CommentFilter as a, type ChangeFilter as b, type ReviewChange as c, type ReplyOptions as d, type ApplyFormattingOptions as e, type SetParagraphStyleOptions as f, type PageContent as g, type ScrollToParaIdOptions as h, type ContentChangeEvent as i, type SelectionChangeEvent as j, type AgentToolDefinition as k, type AgentToolResult as l, type BatchResult as m, type BatchReviewOptions as n, agentTools as o, createReviewerBridge as p, executeToolCall as q, getToolSchemas as r, type EditorRefLike as s, type AgentContextSnapshot as t, getToolDisplayName as u, type ParagraphHighlightOptions as v, createEditorBridge as w };
