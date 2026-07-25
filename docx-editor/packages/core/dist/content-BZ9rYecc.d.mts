import { T as TextFormatting, P as ParagraphFormatting, a as TableFormatting, b as TableRowFormatting, c as TableCellFormatting } from './formatting-DFtuRFQY.mjs';
import { WrapType } from './docx/wrapTypes.mjs';
import { B as BorderSpec, C as ColorValue, T as ThemeColorSlot } from './colors-C3vA7HUU.mjs';
import { N as NumberFormat, c as ListRendering } from './lists-CyGxd5Y2.mjs';
import { W as Watermark } from './watermark-D90356ZM.mjs';

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
type Field = SimpleField | ComplexField;

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
 * Tracked-change attribute triple as it appears on PM node attrs
 * (`paragraph.pPrIns`, `tableRow.trIns`, etc). Mirrors `TrackedChangeInfo`
 * but with a `null` date (PM attr defaults) and a `revisionId` name that
 * matches OOXML's `w:id` more idiomatically on the editor side.
 *
 * Round-trip pairs with `TrackedChangeInfo` via
 * `{ id, author, date? } ↔ { revisionId, author, date | null }`.
 */
interface RevisionInfo {
    revisionId: number;
    author: string;
    date: string | null;
}
/**
 * Tracked-cell marker — the OOXML `<w:cellIns>` / `<w:cellDel>` /
 * `<w:cellMerge>` shape attached to a `TableCell` PM node and surfaced
 * to the layout model and painter for visual rendering.
 *
 * `kind` matches the OOXML element name (ins / del / merge).
 */
interface CellMarker {
    kind: 'ins' | 'del' | 'merge';
    info: RevisionInfo;
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
 * Run-level tracked wrappers represented in WordprocessingML.
 */
type TrackedRunChange = Insertion | Deletion | MoveFrom | MoveTo;
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
 * Text box (floating text container)
 */
interface TextBox {
    type: 'textBox';
    /** Unique ID */
    id?: string;
    /** Size */
    size: ImageSize;
    /** Position */
    position?: ImagePosition;
    /** Wrap settings */
    wrap?: ImageWrap;
    /** Fill */
    fill?: ShapeFill;
    /** Outline */
    outline?: ShapeOutline;
    /** Text content */
    content: Paragraph[];
    /** Internal margins */
    margins?: {
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
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

export type { ImagePadding as $, BookmarkStart as A, BlockContent as B, Comment as C, DocumentBody as D, Endnote as E, Footnote as F, Column as G, Hyperlink as H, Image as I, ComplexField as J, DrawingContent as K, EndnotePosition as L, MoveFrom as M, EndnoteProperties as N, Field as O, Paragraph as P, FieldCharContent as Q, RunPropertyChange as R, SectionProperties as S, Table as T, FieldType as U, FooterReference as V, FootnotePosition as W, FootnoteProperties as X, HeaderFooterType as Y, HeaderReference as Z, ImageCrop as _, RevisionInfo as a, ImagePosition as a0, ImageSize as a1, ImageTransform as a2, ImageWrap as a3, InstrTextContent as a4, LineNumberRestart as a5, MathEquation as a6, MoveFromRangeEnd as a7, MoveFromRangeStart as a8, MoveToRangeEnd as a9, MoveToRangeStart as aa, NoBreakHyphenContent as ab, NoteNumberRestart as ac, NoteRefMarkContent as ad, NoteReferenceContent as ae, PageOrientation as af, PropertyChangeInfo as ag, Section as ah, SectionStart as ai, SeparatorContent as aj, Shape as ak, ShapeContent as al, ShapeFill as am, ShapeOutline as an, ShapeTextBody as ao, ShapeType as ap, SimpleField as aq, SoftHyphenContent as ar, SymbolContent as as, TabContent as at, TableStructuralChangeInfo as au, TextBox as av, VerticalAlign as aw, CellMarker as ax, ParagraphPropertyChange as b, TablePropertyChange as c, TableCellPropertyChange as d, TableRowPropertyChange as e, CommentRangeEnd as f, CommentRangeStart as g, Deletion as h, Insertion as i, MoveTo as j, ParagraphContent as k, Run as l, RunContent as m, TableCell as n, TableRow as o, TextContent as p, TrackedChangeInfo as q, TrackedRunChange as r, SdtType as s, SdtProperties as t, SdtDataBinding as u, BlockSdt as v, InlineSdt as w, BreakContent as x, HeaderFooter as y, BookmarkEnd as z };
