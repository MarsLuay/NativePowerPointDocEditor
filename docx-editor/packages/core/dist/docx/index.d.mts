/**
 * DOCX I/O
 *
 * Parsing DOCX archives into the `Document` model and re-zipping a
 * model back into a DOCX file. Use `./docx/serializer` for the lower-level
 * Document → XML transforms.
 *
 * The named exports below are the public API contract. Adding a parser
 * helper to a source module does not automatically make it public — it
 * must be added to this barrel to be reachable from
 * `@eigenpal/docx-editor-core/docx`.
 * @packageDocumentation
 * @public
 */
export { parseDocx } from './parser.mjs';
import { F as FontTable, b as RelationshipMap, M as MediaFile, a as Style, T as Theme } from '../styles-Diw0MASy.mjs';
export { createDocx, default as repackDocx, updateMultipleFiles } from './rezip.mjs';
import { S as SelectiveSaveOptions } from '../selectiveSave-jinP_4xa.mjs';
export { a as attemptSelectiveSave } from '../selectiveSave-jinP_4xa.mjs';
import { Document } from '../types/document.mjs';
export { b as buildPatchedDocumentXml, v as validatePatchSafety } from '../selectiveXmlPatch-ypkxlTD_.mjs';
import { I as Image, E as Endnote, F as Footnote, N as EndnoteProperties, X as FootnoteProperties, l as Run, U as FieldType, O as Field, aq as SimpleField, H as Hyperlink, ak as Shape, T as Table, n as TableCell, P as Paragraph, av as TextBox, a1 as ImageSize, a0 as ImagePosition, a3 as ImageWrap, B as BlockContent, C as Comment } from '../content-BZ9rYecc.mjs';
import { Element } from 'xml-js';
import { N as NumberingMap } from '../numberingParser-nSPZlzKr.mjs';
export { c as computeListRendering, a as createNumberingMap, f as formatNumber, g as getBulletCharacter, b as getCachedNumberingMap, i as isBulletLevel, p as parseNumbering, r as renderListMarker } from '../numberingParser-nSPZlzKr.mjs';
import { i as TabStopAlignment, h as TabLeader, e as TabStop, j as TableBorders, C as CellMargins, F as FloatingTableProperties, f as TableLook, k as TableMeasurement, a as TableFormatting } from '../formatting-DFtuRFQY.mjs';
export { emuToPixels, pixelsToEmu } from '../utils/units.mjs';
import { B as BorderSpec, S as ShadingProperties } from '../colors-C3vA7HUU.mjs';
import '../docxInput-DTbCa48g.mjs';
import '../lists-CyGxd5Y2.mjs';
import './wrapTypes.mjs';
import '../watermark-D90356ZM.mjs';
import 'jszip';

/**
 * Font Table Parser
 *
 * Parses `word/fontTable.xml` into a {@link FontTable}. The main payload we
 * care about beyond font names is the `w:embed*` elements, which reference the
 * obfuscated font binaries under `word/fonts/` (resolved via the font table's
 * own `.rels`) and carry the `w:fontKey` GUID used to de-obfuscate them.
 *
 * Structure:
 * ```xml
 * <w:fonts>
 *   <w:font w:name="My Brand Sans">
 *     <w:family w:val="swiss"/>
 *     <w:pitch w:val="variable"/>
 *     <w:embedRegular r:id="rId1" w:fontKey="{GUID}" w:subsetted="true"/>
 *   </w:font>
 * </w:fonts>
 * ```
 */

/**
 * Parse `word/fontTable.xml` into a {@link FontTable}. Returns an empty table
 * for missing/empty/unparseable input.
 *
 * @public
 */
declare function parseFontTable(fontTableXml: string | null | undefined): FontTable;

/**
 * Export a Document model to a DOCX ArrayBuffer without DocumentAgent.
 *
 * Same semantics as the former DocumentAgent.toBuffer path: prefer selective
 * save against originalBuffer when options are provided, else full repack.
 */

interface ExportDocxBufferOptions {
    selective?: SelectiveSaveOptions;
}
/**
 * Pack `doc` to DOCX bytes. Mutates `doc.originalBuffer` on successful save
 * so subsequent selective saves patch against the latest baseline.
 */
declare function exportDocxBuffer(doc: Document, options?: ExportDocxBufferOptions): Promise<ArrayBuffer>;

/**
 * Image Parser - Parse embedded images from w:drawing elements
 *
 * DOCX images are contained in <w:drawing> elements with either:
 * - wp:inline - Inline images that flow with text
 * - wp:anchor - Floating/anchored images with text wrapping
 *
 * OOXML Structure:
 * w:drawing
 *   ├── wp:inline or wp:anchor
 *   │   ├── wp:extent (size: cx, cy in EMUs)
 *   │   ├── wp:effectExtent (effect margins)
 *   │   ├── wp:docPr (document properties: id, name, descr, title)
 *   │   ├── wp:positionH / wp:positionV (for anchor only)
 *   │   ├── wp:wrap* (wrapping mode for anchor: wrapNone, wrapSquare, etc.)
 *   │   └── a:graphic
 *   │       └── a:graphicData
 *   │           └── pic:pic
 *   │               ├── pic:nvPicPr (non-visual properties)
 *   │               ├── pic:blipFill
 *   │               │   └── a:blip (r:embed = rId)
 *   │               └── pic:spPr
 *   │                   └── a:xfrm (transform: rotation, flip)
 *
 * EMU (English Metric Units): 914400 EMU = 1 inch
 * Conversion: pixels = (emu * 96) / 914400
 */

/**
 * Parse a w:drawing element
 *
 * The drawing element contains either wp:inline or wp:anchor.
 *
 * @param drawingEl - The w:drawing element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object or null if not an image
 */
declare function parseDrawing(drawingEl: Element, rels: RelationshipMap | undefined, media: Map<string, MediaFile> | undefined): Image | null;
/**
 * Parse an image from a w:drawing element
 *
 * This is the main entry point for image parsing.
 *
 * @param node - The w:drawing XML element
 * @param rels - Relationship map for resolving rId
 * @param media - Media files map
 * @returns Parsed Image object or null if parsing fails
 */
declare function parseImage(node: Element, rels: RelationshipMap | undefined, media: Map<string, MediaFile> | undefined): Image | null;
/**
 * Check if an image is inline (not floating)
 */
declare function isInlineImage(image: Image): boolean;
/**
 * Check if an image is floating (anchored)
 */
declare function isFloatingImage(image: Image): boolean;
/**
 * Check if an image is behind text
 */
declare function isBehindText(image: Image): boolean;
/**
 * Check if an image is in front of text
 */
declare function isInFrontOfText(image: Image): boolean;
/**
 * Get image width in pixels
 */
declare function getImageWidthPx(image: Image): number;
/**
 * Get image height in pixels
 */
declare function getImageHeightPx(image: Image): number;
/**
 * Check if image is decorative (should be ignored by screen readers)
 */
declare function isDecorativeImage(image: Image): boolean;
/**
 * Get wrap distances in pixels
 */
declare function getWrapDistancesPx(image: Image): {
    top: number;
    bottom: number;
    left: number;
    right: number;
};

/**
 * Style Parser - Parse styles.xml with full inheritance resolution
 *
 * Parses all style types (paragraph, character, table, list) with
 * complete basedOn inheritance chain resolution.
 *
 * OOXML Reference:
 * - Style file is at: word/styles.xml
 * - Uses WordprocessingML namespace (w:)
 *
 * Style Cascade (lowest to highest priority):
 * 1. Document defaults (w:docDefaults)
 * 2. Parent style properties (w:basedOn chain)
 * 3. Current style properties
 * 4. Direct formatting in document
 *
 * This file owns the style element itself, docDefaults, inheritance
 * resolution, and the top-level entry points (parseStyles,
 * parseStyleDefinitions, getResolved*). Property-level parsers live under
 * ./styleParser/{runProperties,paragraphProperties,tableProperties}.ts.
 */

/**
 * Style map keyed by styleId
 */
type StyleMap = Map<string, Style>;

/**
 * Footnote/Endnote Parser - Parse footnotes.xml and endnotes.xml
 *
 * Footnotes and endnotes are stored in separate XML files within the DOCX package:
 * - word/footnotes.xml - Contains all footnote definitions
 * - word/endnotes.xml - Contains all endnote definitions
 *
 * Each note contains:
 * - An ID that matches references in document.xml (w:footnoteReference, w:endnoteReference)
 * - A type (normal, separator, continuationSeparator, continuationNotice)
 * - Content (paragraphs)
 *
 * The references in the document body are parsed by runParser as NoteReferenceContent.
 *
 * OOXML Reference:
 * - Footnote: w:footnote[@w:id][@w:type]
 * - Endnote: w:endnote[@w:id][@w:type]
 * - Content: w:p (paragraphs)
 */

/**
 * Footnote map returned by parseFootnotes
 */
interface FootnoteMap {
    /** All footnotes indexed by ID */
    byId: Map<number, Footnote>;
    /** Array of all footnotes in document order */
    footnotes: Footnote[];
    /** Get footnote by ID */
    getFootnote(id: number): Footnote | undefined;
    /** Check if footnote exists */
    hasFootnote(id: number): boolean;
    /** Get all normal (non-separator) footnotes */
    getNormalFootnotes(): Footnote[];
    /** Get separator footnote if exists */
    getSeparator(): Footnote | undefined;
    /** Get continuation separator if exists */
    getContinuationSeparator(): Footnote | undefined;
}
/**
 * Endnote map returned by parseEndnotes
 */
interface EndnoteMap {
    /** All endnotes indexed by ID */
    byId: Map<number, Endnote>;
    /** Array of all endnotes in document order */
    endnotes: Endnote[];
    /** Get endnote by ID */
    getEndnote(id: number): Endnote | undefined;
    /** Check if endnote exists */
    hasEndnote(id: number): boolean;
    /** Get all normal (non-separator) endnotes */
    getNormalEndnotes(): Endnote[];
    /** Get separator endnote if exists */
    getSeparator(): Endnote | undefined;
    /** Get continuation separator if exists */
    getContinuationSeparator(): Endnote | undefined;
}
/**
 * Parse footnotes.xml
 *
 * @param footnotesXml - The raw XML content of word/footnotes.xml
 * @param styles - Parsed style map for applying styles
 * @param theme - Parsed theme for color resolution
 * @param numbering - Parsed numbering definitions for lists
 * @param rels - Relationships for resolving hyperlinks
 * @param media - Media files for images
 * @returns FootnoteMap with all footnotes
 */
declare function parseFootnotes(footnotesXml: string | null, styles?: StyleMap | null, theme?: Theme | null, numbering?: NumberingMap | null, rels?: RelationshipMap | null, media?: Map<string, MediaFile> | null): FootnoteMap;
/**
 * Parse endnotes.xml
 *
 * @param endnotesXml - The raw XML content of word/endnotes.xml
 * @param styles - Parsed style map for applying styles
 * @param theme - Parsed theme for color resolution
 * @param numbering - Parsed numbering definitions for lists
 * @param rels - Relationships for resolving hyperlinks
 * @param media - Media files for images
 * @returns EndnoteMap with all endnotes
 */
declare function parseEndnotes(endnotesXml: string | null, styles?: StyleMap | null, theme?: Theme | null, numbering?: NumberingMap | null, rels?: RelationshipMap | null, media?: Map<string, MediaFile> | null): EndnoteMap;
/**
 * Parse footnote properties from w:footnotePr element
 * (Can appear in w:sectPr or w:settings)
 */
declare function parseFootnoteProperties(element: Element | null): FootnoteProperties;
/**
 * Parse endnote properties from w:endnotePr element
 * (Can appear in w:sectPr or w:settings)
 */
declare function parseEndnoteProperties(element: Element | null): EndnoteProperties;
/**
 * Get plain text content of a footnote
 */
declare function getFootnoteText(footnote: Footnote): string;
/**
 * Get plain text content of an endnote
 */
declare function getEndnoteText(endnote: Endnote): string;
/**
 * Check if a footnote is a separator (not regular content)
 */
declare function isSeparatorFootnote(footnote: Footnote): boolean;
/**
 * Check if an endnote is a separator (not regular content)
 */
declare function isSeparatorEndnote(endnote: Endnote): boolean;

/**
 * Field Parser - Parse field codes in DOCX documents
 *
 * OOXML supports two types of fields:
 * 1. Simple fields (w:fldSimple) - Single element with instruction attribute
 * 2. Complex fields (w:fldChar + w:instrText) - Multi-element spanning runs
 *
 * Fields provide dynamic content like:
 * - Page numbers (PAGE, NUMPAGES)
 * - Dates and times (DATE, TIME, CREATEDATE)
 * - Document properties (AUTHOR, TITLE, FILENAME)
 * - Cross-references (REF, PAGEREF, NOTEREF)
 * - Tables of contents (TOC, INDEX)
 * - Mail merge fields (MERGEFIELD)
 *
 * OOXML Reference:
 * - Simple field: <w:fldSimple w:instr="FIELD INSTRUCTION">content</w:fldSimple>
 * - Complex field:
 *   <w:r><w:fldChar w:fldCharType="begin"/></w:r>
 *   <w:r><w:instrText>FIELD INSTRUCTION</w:instrText></w:r>
 *   <w:r><w:fldChar w:fldCharType="separate"/></w:r>
 *   <w:r><w:t>display result</w:t></w:r>
 *   <w:r><w:fldChar w:fldCharType="end"/></w:r>
 */

/**
 * All known field types from OOXML specification
 */
declare const KNOWN_FIELD_TYPES: FieldType[];
/**
 * Parse field type from instruction string
 *
 * Field instructions follow the format: FIELDNAME [arguments] [switches]
 * Examples:
 * - "PAGE \\* MERGEFORMAT"
 * - "DATE \\@ \"MMMM d, yyyy\""
 * - "MERGEFIELD client_name \\* Upper"
 * - "REF _Ref123456 \\h"
 *
 * @param instruction - The field instruction string
 * @returns The detected field type
 */
declare function parseFieldType(instruction: string): FieldType;
/**
 * Check if a field type is a known type
 *
 * @param type - Field type string to check
 * @returns true if it's a known field type
 */
declare function isKnownFieldType(type: string): type is FieldType;
/**
 * Parsed field instruction with arguments and switches
 */
interface ParsedFieldInstruction {
    /** Field type */
    type: FieldType;
    /** Raw instruction string */
    raw: string;
    /** Field argument (e.g., property name for DOCPROPERTY, bookmark name for REF) */
    argument?: string;
    /** Field switches (e.g., \* MERGEFORMAT, \@ "date format") */
    switches: FieldSwitch[];
}
/**
 * Field switch parsed from instruction
 */
interface FieldSwitch {
    /** Switch character (e.g., '*', '@', '#', 'h', 'p') */
    switch: string;
    /** Switch value if any */
    value?: string;
}
/**
 * Parse a complete field instruction into structured data
 *
 * @param instruction - Raw instruction string
 * @returns Parsed instruction object
 */
declare function parseFieldInstruction(instruction: string): ParsedFieldInstruction;
/**
 * Get the format switch value (\* or \@)
 *
 * @param instruction - Parsed instruction
 * @returns Format string or undefined
 */
declare function getFormatSwitch(instruction: ParsedFieldInstruction): string | undefined;
/**
 * Check if field has MERGEFORMAT switch (preserve formatting)
 *
 * @param instruction - Parsed instruction
 * @returns true if MERGEFORMAT is present
 */
declare function hasMergeFormat(instruction: ParsedFieldInstruction): boolean;
/**
 * Parse a simple field element (w:fldSimple)
 *
 * @param node - The w:fldSimple XML element
 * @param styles - Style definitions for parsing content runs
 * @param theme - Theme for color/font resolution
 * @returns Parsed SimpleField object
 */
declare function parseSimpleField(node: Element, styles: StyleMap | null, theme: Theme | null): SimpleField;
/**
 * State machine for tracking complex field parsing
 */
type ComplexFieldState = 'outside' | 'code' | 'result';
/**
 * Complex field parsing context
 */
interface ComplexFieldContext {
    /** Current state */
    state: ComplexFieldState;
    /** Accumulated instruction text */
    instruction: string;
    /** Runs in the field code section */
    codeRuns: Run[];
    /** Runs in the result section */
    resultRuns: Run[];
    /** Whether field is locked */
    fldLock: boolean;
    /** Whether field needs update */
    dirty: boolean;
    /** Nesting level (for nested fields) */
    nestingLevel: number;
}
/**
 * Create a new complex field context
 */
declare function createComplexFieldContext(): ComplexFieldContext;
/**
 * Get the current display value of a field
 *
 * @param field - The field (simple or complex)
 * @returns The display text
 */
declare function getFieldDisplayValue(field: Field): string;
/**
 * Check if field represents a page number
 *
 * @param field - The field to check
 * @returns true if this is a page number field
 */
declare function isPageNumberField(field: Field): boolean;
/**
 * Check if field represents total page count
 *
 * @param field - The field to check
 * @returns true if this is a total pages field
 */
declare function isTotalPagesField(field: Field): boolean;
/**
 * Check if field is a date/time field
 *
 * @param field - The field to check
 * @returns true if this is a date/time field
 */
declare function isDateTimeField(field: Field): boolean;
/**
 * Check if field is a document property field
 *
 * @param field - The field to check
 * @returns true if this is a document property field
 */
declare function isDocPropertyField(field: Field): boolean;
/**
 * Check if field is a cross-reference field
 *
 * @param field - The field to check
 * @returns true if this is a cross-reference field
 */
declare function isReferenceField(field: Field): boolean;
/**
 * Check if field is a mail merge field
 *
 * @param field - The field to check
 * @returns true if this is a mail merge field
 */
declare function isMergeField(field: Field): boolean;
/**
 * Check if field is a TOC/Index field
 *
 * @param field - The field to check
 * @returns true if this is a TOC or index field
 */
declare function isTocField(field: Field): boolean;

/**
 * Hyperlink Parser - Parse hyperlinks (w:hyperlink) with URL resolution
 *
 * OOXML Reference:
 * - Hyperlink element: w:hyperlink
 * - Attributes:
 *   - r:id - Relationship ID for external link (resolves via .rels)
 *   - w:anchor - Internal bookmark name
 *   - w:tooltip - Tooltip/title text
 *   - w:tgtFrame - Target frame (_blank, _self, etc.)
 *   - w:history - Whether to add to history
 *   - w:docLocation - Location within a document
 *
 * External links use r:id to reference a relationship in document.xml.rels
 * Internal links use w:anchor to reference a bookmark in the same document
 */

/**
 * Parse a hyperlink element (w:hyperlink)
 *
 * Handles both external links (via r:id relationship) and internal
 * links (via w:anchor bookmark reference).
 *
 * @param node - The w:hyperlink XML element
 * @param rels - Relationship map to resolve r:id references
 * @param styles - Style map for resolving run styles
 * @param theme - Theme for resolving colors/fonts
 * @param media - Media files map for image data
 * @returns Parsed Hyperlink object
 */
declare function parseHyperlink(node: Element, rels: RelationshipMap | null, styles?: StyleMap | null, theme?: Theme | null, media?: Map<string, MediaFile> | null): Hyperlink;
/**
 * Get the display text of a hyperlink
 *
 * Concatenates text from all child runs.
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Display text string
 */
declare function getHyperlinkText(hyperlink: Hyperlink): string;
/**
 * Check if a hyperlink is an external link
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if this links to an external URL
 */
declare function isExternalLink(hyperlink: Hyperlink): boolean;
/**
 * Check if a hyperlink is an internal bookmark link
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if this links to an internal bookmark
 */
declare function isInternalLink(hyperlink: Hyperlink): boolean;
/**
 * Get the resolved URL of a hyperlink
 *
 * For external links, returns the full URL.
 * For internal links, returns the anchor prefixed with #.
 * Returns undefined if the link couldn't be resolved.
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Resolved URL or undefined
 */
declare function getHyperlinkUrl(hyperlink: Hyperlink): string | undefined;
/**
 * Check if a hyperlink has any content (runs)
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns true if hyperlink has child runs
 */
declare function hasContent(hyperlink: Hyperlink): boolean;
/**
 * Get all runs from a hyperlink
 *
 * @param hyperlink - Parsed Hyperlink object
 * @returns Array of Run objects
 */
declare function getHyperlinkRuns(hyperlink: Hyperlink): Run[];
/**
 * Resolve a hyperlink's rId to a URL using a relationship map
 *
 * This is useful when you have a hyperlink that was parsed without
 * relationship context and need to resolve it later.
 *
 * @param hyperlink - Parsed Hyperlink object (will be modified)
 * @param rels - Relationship map to resolve against
 * @returns The resolved URL or undefined
 */
declare function resolveHyperlinkUrl(hyperlink: Hyperlink, rels: RelationshipMap): string | undefined;

/**
 * Shape Parser - Parse shapes and drawings from wps:wsp elements
 *
 * DOCX shapes are contained in drawings with wps:wsp (Word Processing Shape) elements.
 * Shapes can be standalone or inside groups (wpg:wgp).
 *
 * OOXML Structure:
 * w:drawing
 *   └── wp:inline or wp:anchor
 *       └── a:graphic
 *           └── a:graphicData
 *               └── wps:wsp (shape)
 *                   ├── wps:cNvSpPr (non-visual properties)
 *                   ├── wps:spPr (shape properties)
 *                   │   ├── a:xfrm (transform: position, size, rotation)
 *                   │   ├── a:prstGeom (preset geometry/shape type)
 *                   │   ├── a:solidFill / a:noFill / a:gradFill (fill)
 *                   │   └── a:ln (line/outline properties)
 *                   ├── wps:style (style reference)
 *                   ├── wps:txbx (text box container)
 *                   │   └── w:txbxContent (text content)
 *                   └── wps:bodyPr (body/text properties)
 *
 * EMU (English Metric Units): 914400 EMU = 1 inch
 */

/**
 * Parse a wps:wsp (Word Processing Shape) element
 *
 * @param node - The wps:wsp XML element
 * @returns Parsed Shape object
 */
declare function parseShape(node: Element): Shape;
/**
 * Parse shape from a w:drawing element that contains a shape (not an image)
 *
 * @param drawingEl - The w:drawing element
 * @returns Parsed Shape object or null if not a shape
 */
declare function parseShapeFromDrawing(drawingEl: Element): Shape | null;
/**
 * Check if a drawing element contains a shape (not an image)
 */
declare function isShapeDrawing(drawingEl: Element): boolean;
/**
 * Check if a shape is a line (connector)
 */
declare function isLineShape(shape: Shape): boolean;
/**
 * Check if a shape is a text box
 */
declare function isTextBoxShape(shape: Shape): boolean;
/**
 * Check if a shape has text content
 */
declare function hasTextContent(shape: Shape): boolean;
/**
 * Get shape width in pixels
 */
declare function getShapeWidthPx(shape: Shape): number;
/**
 * Get shape height in pixels
 */
declare function getShapeHeightPx(shape: Shape): number;
/**
 * Get shape dimensions in pixels
 */
declare function getShapeDimensionsPx(shape: Shape): {
    width: number;
    height: number;
};
/**
 * Check if shape is floating (anchored)
 */
declare function isFloatingShape(shape: Shape): boolean;
/**
 * Check if shape has fill
 */
declare function hasFill(shape: Shape): boolean;
/**
 * Check if shape has outline
 */
declare function hasOutline(shape: Shape): boolean;
/**
 * Get outline width in pixels
 */
declare function getOutlineWidthPx(shape: Shape): number;
/**
 * Resolve fill color to CSS color string
 */
declare function resolveFillColor(shape: Shape): string | undefined;
/**
 * Resolve outline color to CSS color string
 */
declare function resolveOutlineColor(shape: Shape): string | undefined;

/**
 * Tab Parser - Parse and handle tab stops in DOCX documents
 *
 * Tab stops define positions where the cursor jumps when the user presses Tab.
 * They can have different alignments (left, center, right, decimal) and
 * leader characters (dots, dashes, underscores).
 *
 * OOXML Reference:
 * - Tab stops container: w:tabs
 * - Individual tab stop: w:tab
 * - Tab character in runs: w:tab (different from tab stop definition)
 *
 * Attributes of w:tab in w:tabs:
 * - w:val - alignment type (left, center, right, decimal, bar, clear, num)
 * - w:pos - position in twips from left margin
 * - w:leader - leader character (none, dot, hyphen, underscore, heavy, middleDot)
 */

/**
 * Default tab stop interval in twips (0.5 inches = 720 twips at 1440 twips/inch)
 * Word uses this when no explicit tab stops are defined
 */
declare const DEFAULT_TAB_INTERVAL_TWIPS = 720;
/**
 * Default tab alignment
 */
declare const DEFAULT_TAB_ALIGNMENT: TabStopAlignment;
/**
 * Default tab leader
 */
declare const DEFAULT_TAB_LEADER: TabLeader;
/**
 * Parse a single tab stop element (w:tab within w:tabs)
 *
 * @param tab - The w:tab XML element
 * @returns Parsed TabStop or null if invalid
 */
declare function parseTabStop(tab: Element): TabStop | null;
/**
 * Parse tab stops container (w:tabs)
 *
 * @param tabs - The w:tabs XML element
 * @returns Array of TabStop objects, sorted by position
 */
declare function parseTabStops(tabs: Element | null): TabStop[];
/**
 * Parse tab stops from paragraph properties element
 *
 * @param pPr - The w:pPr XML element
 * @returns Array of TabStop objects or undefined if none
 */
declare function parseTabStopsFromParagraphProperties(pPr: Element | null): TabStop[] | undefined;
/**
 * Merge tab stops from different sources (style, direct formatting)
 *
 * Direct formatting tab stops override style tab stops at the same position.
 * "clear" alignment removes a tab stop from the style.
 *
 * @param styleTabs - Tab stops from style
 * @param directTabs - Tab stops from direct formatting (w:pPr in paragraph)
 * @returns Merged and filtered tab stops
 */
declare function mergeTabStops(styleTabs: TabStop[] | undefined, directTabs: TabStop[] | undefined): TabStop[];
/**
 * Get the next tab stop position for a given current position
 *
 * @param currentPosition - Current position in twips from left margin
 * @param tabStops - Defined tab stops
 * @param pageWidth - Page content width in twips (for boundary)
 * @returns The next tab stop or a default position
 */
declare function getNextTabStop(currentPosition: number, tabStops: TabStop[], pageWidth: number): TabStop;
/**
 * Calculate the width needed for a tab at a given position
 *
 * @param currentPosition - Current position in twips
 * @param tabStops - Defined tab stops
 * @param pageWidth - Page content width in twips
 * @returns Width in twips that the tab should span
 */
declare function calculateTabWidth(currentPosition: number, tabStops: TabStop[], pageWidth: number): number;
/**
 * Calculate tab width considering alignment
 *
 * For non-left alignments (center, right, decimal), the width depends on
 * the content that follows the tab.
 *
 * @param currentPosition - Current position in twips
 * @param tabStops - Defined tab stops
 * @param pageWidth - Page content width in twips
 * @param followingContentWidth - Width of content after the tab (for alignment)
 * @returns Width in twips
 */
declare function calculateTabWidthWithAlignment(currentPosition: number, tabStops: TabStop[], pageWidth: number, followingContentWidth?: number): {
    width: number;
    alignment: TabStopAlignment;
};
/**
 * Get the character used for a tab leader
 *
 * @param leader - Tab leader type
 * @returns The character to use for filling
 */
declare function getLeaderCharacter(leader: TabLeader | undefined): string;
/**
 * Check if a leader type requires visible filling
 *
 * @param leader - Tab leader type
 * @returns true if the leader needs visible characters
 */
declare function hasVisibleLeader(leader: TabLeader | undefined): boolean;

/**
 * Shared border parsing for every `CT_Border` field — table borders
 * (`w:tblBorders`/`w:tcBorders`), paragraph borders (`w:pBdr`), and page borders
 * (`w:pgBorders`). Mirrors {@link ../serializer/borderSerializer} on the parse
 * side: the table/paragraph/section/style parsers all delegate here so a border
 * is read identically everywhere (one place to fix, one place to harden).
 */

/**
 * Parse a single border element (`<w:top>`, `<w:left>`, ...) into a
 * {@link BorderSpec}.
 *
 * A missing `w:val` defaults to `none` (the attribute is required by the schema,
 * so this only guards malformed input) and `w:color="auto"` maps to
 * `{ auto: true }` via {@link parseColorValue} rather than being stored as a
 * literal `rgb`. Returns `undefined` only when the element itself is absent.
 */
declare function parseBorderSpec(element: Element | null): BorderSpec | undefined;
/**
 * Parse a borders container (`w:tblBorders` or `w:tcBorders`) into
 * {@link TableBorders}. `left`/`right` fall back to the RTL `start`/`end`
 * aliases. Returns `undefined` when no side parsed.
 */
declare function parseTableBorders(bordersElement: Element | null): TableBorders | undefined;

/**
 * Table Atom-level Property Parsers
 *
 * Parsers for the small, leaf-style table property elements:
 * - measurements (w:tblW, w:tcW, w:trHeight twip values + width type)
 * - borders (w:tblBorders, w:tcBorders, individual side specs)
 * - cell margins (w:tblCellMar, w:tcMar)
 * - shading (w:shd background + theme tint/shade)
 * - table look flags (w:tblLook firstRow/lastRow/etc.)
 * - floating table positioning (w:tblpPr anchors + offsets)
 *
 * These are leaf parsers — they don't recurse into rows/cells. The composite
 * row/cell/table parsers (in `../tableParser.ts`) compose these.
 */

/**
 * Parse a table measurement (width, height, etc.)
 *
 * @param element - Element with w:w and w:type attributes
 * @returns Parsed measurement or undefined
 */
declare function parseTableMeasurement(element: Element | null): TableMeasurement | undefined;
/**
 * Parse cell margins (w:tblCellMar or w:tcMar)
 *
 * @param marginsElement - The margins container element
 * @returns Parsed margins or undefined
 */
declare function parseCellMargins(marginsElement: Element | null): CellMargins | undefined;
/**
 * Parse shading properties (w:shd)
 *
 * @param shdElement - The w:shd element
 * @returns Parsed shading or undefined
 */
declare function parseShading(shdElement: Element | null): ShadingProperties | undefined;
/**
 * Parse table look flags (w:tblLook)
 *
 * @param lookElement - The w:tblLook element
 * @returns Parsed table look or undefined
 */
declare function parseTableLook(lookElement: Element | null): TableLook | undefined;
/**
 * Parse floating table properties (w:tblpPr)
 *
 * @param tblpPrElement - The w:tblpPr element
 * @returns Parsed floating properties or undefined
 */
declare function parseFloatingTableProperties(tblpPrElement: Element | null): FloatingTableProperties | undefined;

/**
 * Table Query Helpers
 *
 * Read-only utilities for inspecting a parsed Table model — column/row
 * counts, cell merge state, plain-text extraction, header detection,
 * floating-table detection. No XML access; these operate on the parsed
 * `Table` shape.
 */

/**
 * Get the number of columns in a table
 *
 * Uses the table grid if available, otherwise counts cells in first row.
 *
 * @param table - The table to measure
 * @returns Number of columns
 */
declare function getTableColumnCount(table: Table): number;
/**
 * Get the number of rows in a table
 *
 * @param table - The table to measure
 * @returns Number of rows
 */
declare function getTableRowCount(table: Table): number;
/**
 * Check if a cell is part of a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell continues a vertical merge
 */
declare function isCellMergeContinuation(cell: TableCell): boolean;
/**
 * Check if a cell starts a vertical merge
 *
 * @param cell - The cell to check
 * @returns true if cell starts a vertical merge
 */
declare function isCellMergeStart(cell: TableCell): boolean;
/**
 * Check if table has header row
 *
 * @param table - The table to check
 * @returns true if first row is marked as header
 */
declare function hasHeaderRow(table: Table): boolean;

/**
 * Table Parser - Parse tables with full OOXML structure
 *
 * OOXML tables consist of:
 * - w:tbl - Table element
 * - w:tblPr - Table properties (width, borders, style)
 * - w:tblGrid - Column width definitions
 * - w:tr - Table rows
 * - w:trPr - Row properties (height, header)
 * - w:tc - Table cells
 * - w:tcPr - Cell properties (width, borders, merge)
 *
 * Cell merging:
 * - Horizontal: w:gridSpan (how many grid columns this cell spans)
 * - Vertical: w:vMerge (restart = start of merge, continue = continuation)
 *
 * OOXML Reference:
 * - w:tbl contains w:tblPr, w:tblGrid, and w:tr elements
 * - w:tr contains w:trPr and w:tc elements
 * - w:tc contains w:tcPr and content (paragraphs, tables)
 *
 * Composite parsers (parseTable, parseTableRow, parseTableCell, plus the
 * three property parsers and their tracked-change variants) live here.
 * Leaf property parsers (measurements, borders, margins, shading, look,
 * floating) are in ./tableParser/properties.ts. Query helpers (counts, merge
 * checks, text extraction) are in ./tableParser/queries.ts.
 */

/**
 * Parse table properties (w:tblPr)
 *
 * @param tblPrElement - The w:tblPr element
 * @returns Parsed table formatting
 */
declare function parseTableProperties(tblPrElement: Element | null): TableFormatting | undefined;

/**
 * Text Box Parser - Parse floating text box containers
 *
 * Text boxes in DOCX are implemented as shapes (wps:wsp) with text body content (wps:txbx).
 * The text body contains w:txbxContent which holds paragraphs and tables like the main document.
 *
 * OOXML Structure:
 * w:drawing
 *   └── wp:inline or wp:anchor
 *       └── a:graphic
 *           └── a:graphicData
 *               └── wps:wsp (shape)
 *                   ├── wps:cNvSpPr (non-visual properties)
 *                   ├── wps:spPr (shape properties)
 *                   │   ├── a:xfrm (transform: position, size)
 *                   │   ├── a:prstGeom (preset geometry - typically "rect" for text boxes)
 *                   │   ├── a:solidFill / a:noFill (fill)
 *                   │   └── a:ln (outline)
 *                   ├── wps:txbx (text box container)
 *                   │   └── w:txbxContent (text content)
 *                   │       ├── w:p (paragraphs)
 *                   │       └── w:tbl (tables)
 *                   └── wps:bodyPr (body properties - margins, text direction, etc.)
 *
 * EMU (English Metric Units): 914400 EMU = 1 inch
 */

/**
 * Extract raw paragraph elements from w:txbxContent
 * Actual parsing happens via document parser to avoid circular dependencies
 */
declare function extractTextBoxContentElements(txbxContent: Element | null): {
    paragraphElements: Element[];
    tableElements: Element[];
};
/**
 * Type for the paragraph parser function to avoid circular imports
 */
type ParagraphParserFn = (node: Element, styles: StyleMap | null, theme: Theme | null, numbering: NumberingMap | null, rels?: RelationshipMap | null) => Paragraph;
/**
 * Type for the table parser function to avoid circular imports
 */
type TableParserFn = (node: Element, styles: StyleMap | null, theme: Theme | null, numbering: NumberingMap | null, rels?: RelationshipMap | null, media?: Map<string, MediaFile>) => Table;
/**
 * Parse text box content with provided parser functions
 * This avoids circular dependencies by accepting parser functions as parameters
 */
declare function parseTextBoxContent(txbxContent: Element | null, parseParagraph: ParagraphParserFn, parseTable: TableParserFn | null, styles: StyleMap | null, theme: Theme | null, numbering: NumberingMap | null, rels?: RelationshipMap | null, _media?: Map<string, MediaFile>): Paragraph[];
/**
 * Check if a drawing element contains a text box
 * Text boxes are shapes with wps:txbx content
 */
declare function isTextBoxDrawing(drawingEl: Element): boolean;
/**
 * Check if a wps:wsp element is a text box
 */
declare function isShapeTextBox(wsp: Element): boolean;
/**
 * Parse a text box from a w:drawing element
 *
 * This creates a TextBox object with placeholder content.
 * The actual content parsing requires paragraph/table parsers which
 * creates a circular dependency. The document parser should call
 * parseTextBoxContent() separately with the required parsers.
 *
 * @param drawingEl - The w:drawing XML element
 * @returns TextBox object with placeholder content, or null if not a text box
 */
declare function parseTextBox(drawingEl: Element): TextBox | null;
/**
 * Parse text box content XML element
 * @param wsp - The wps:wsp element containing the text box
 * @returns The w:txbxContent element or null
 */
declare function getTextBoxContentElement(wsp: Element): Element | null;
/**
 * Parse text box from a wps:wsp element directly
 * Useful when you already have the shape element
 */
declare function parseTextBoxFromShape(wsp: Element, size: ImageSize, position?: ImagePosition, wrap?: ImageWrap): TextBox | null;
/**
 * Get text box width in pixels
 */
declare function getTextBoxWidthPx(textBox: TextBox): number;
/**
 * Get text box height in pixels
 */
declare function getTextBoxHeightPx(textBox: TextBox): number;
/**
 * Get text box dimensions in pixels
 */
declare function getTextBoxDimensionsPx(textBox: TextBox): {
    width: number;
    height: number;
};
/**
 * Get text box margins in pixels
 */
declare function getTextBoxMarginsPx(textBox: TextBox): {
    top: number;
    bottom: number;
    left: number;
    right: number;
};
/**
 * Check if text box is floating (anchored)
 */
declare function isFloatingTextBox(textBox: TextBox): boolean;
/**
 * Check if text box has fill
 */
declare function hasTextBoxFill(textBox: TextBox): boolean;
/**
 * Check if text box has outline
 */
declare function hasTextBoxOutline(textBox: TextBox): boolean;
/**
 * Check if text box has content
 */
declare function hasTextBoxContent(textBox: TextBox): boolean;
/**
 * Get plain text from text box (helper for search/indexing)
 */
declare function getTextBoxText(textBox: TextBox): string;
/**
 * Resolve fill color to CSS color string
 */
declare function resolveTextBoxFillColor(textBox: TextBox): string | undefined;
/**
 * Resolve outline color to CSS color string
 */
declare function resolveTextBoxOutlineColor(textBox: TextBox): string | undefined;
/**
 * Get outline width in pixels
 */
declare function getTextBoxOutlineWidthPx(textBox: TextBox): number;

/**
 * Reply-range marker injection for serialization.
 *
 * Word / Pages / LibreOffice expect every comment in `comments.xml`
 * (including REPLY threads) to have matching `commentRangeStart` /
 * `commentRangeEnd` / `commentReference` markers in `document.xml`.
 * The PM document only stamps marks for the parent comment because
 * replies don't have their own visible range — they share the parent
 * thread's text. So before serialization we walk the body content and
 * synthesize parallel range markers for every reply.
 *
 * Two helpers, one per parent shape:
 * - `injectReplyRangeMarkers` — replies whose parent is another
 *   comment (regular threaded discussion). Finds the parent's
 *   `commentRangeStart`/`End` and adds parallel markers next to them.
 * - `injectTCReplyRangeMarkers` — replies whose parent is a tracked
 *   change (insertion/deletion). Wraps the TC content with
 *   commentRange markers.
 *
 * Pre-#... this code lived inside React's DocxEditor.tsx; Vue had no
 * equivalent and so silently lost reply markers when saving collab
 * documents. Living in core means both adapters get it for free.
 */

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for reply comments
 * that share their parent comment's text range.
 */
declare function injectReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void;
/**
 * Inject `commentRangeStart`/`commentRangeEnd` for comments whose
 * parent is a tracked-change revision (insertion/deletion). The TC
 * content nodes don't carry the comment's range, so we wrap them.
 */
declare function injectTCReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void;

export { type ComplexFieldContext, type ComplexFieldState, DEFAULT_TAB_ALIGNMENT, DEFAULT_TAB_INTERVAL_TWIPS, DEFAULT_TAB_LEADER, type EndnoteMap, type ExportDocxBufferOptions, type FieldSwitch, type FootnoteMap, KNOWN_FIELD_TYPES, NumberingMap, type ParagraphParserFn, type ParsedFieldInstruction, SelectiveSaveOptions, type TableParserFn, calculateTabWidth, calculateTabWidthWithAlignment, createComplexFieldContext, exportDocxBuffer, extractTextBoxContentElements, getEndnoteText, getFieldDisplayValue, getFootnoteText, getFormatSwitch, getHyperlinkRuns, getHyperlinkText, getHyperlinkUrl, getImageHeightPx, getImageWidthPx, getLeaderCharacter, getNextTabStop, getOutlineWidthPx, getShapeDimensionsPx, getShapeHeightPx, getShapeWidthPx, getTableColumnCount, getTableRowCount, getTextBoxContentElement, getTextBoxDimensionsPx, getTextBoxHeightPx, getTextBoxMarginsPx, getTextBoxOutlineWidthPx, getTextBoxText, getTextBoxWidthPx, getWrapDistancesPx, hasContent, hasFill, hasHeaderRow, hasMergeFormat, hasOutline, hasTextBoxContent, hasTextBoxFill, hasTextBoxOutline, hasTextContent, hasVisibleLeader, injectReplyRangeMarkers, injectTCReplyRangeMarkers, isBehindText, isCellMergeContinuation, isCellMergeStart, isDateTimeField, isDecorativeImage, isDocPropertyField, isExternalLink, isFloatingImage, isFloatingShape, isFloatingTextBox, isInFrontOfText, isInlineImage, isInternalLink, isKnownFieldType, isLineShape, isMergeField, isPageNumberField, isReferenceField, isSeparatorEndnote, isSeparatorFootnote, isShapeDrawing, isShapeTextBox, isTextBoxDrawing, isTextBoxShape, isTocField, isTotalPagesField, mergeTabStops, parseBorderSpec, parseCellMargins, parseDrawing, parseEndnoteProperties, parseEndnotes, parseFieldInstruction, parseFieldType, parseFloatingTableProperties, parseFontTable, parseFootnoteProperties, parseFootnotes, parseHyperlink, parseImage, parseShading, parseShape, parseShapeFromDrawing, parseSimpleField, parseTabStop, parseTabStops, parseTabStopsFromParagraphProperties, parseTableBorders, parseTableLook, parseTableMeasurement, parseTableProperties, parseTextBox, parseTextBoxContent, parseTextBoxFromShape, resolveFillColor, resolveHyperlinkUrl, resolveOutlineColor, resolveTextBoxFillColor, resolveTextBoxOutlineColor };
