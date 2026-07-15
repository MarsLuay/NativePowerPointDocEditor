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
export { B as BorderSpec, C as ColorValue, S as ShadingProperties, T as ThemeColorSlot } from '../colors-C3vA7HUU.mjs';
export { C as CellMargins, g as ConditionalFormatStyle, E as EmphasisMark, F as FloatingTableProperties, L as LineSpacingRule, d as ParagraphAlignment, P as ParagraphFormatting, h as TabLeader, e as TabStop, i as TabStopAlignment, j as TableBorders, c as TableCellFormatting, a as TableFormatting, f as TableLook, k as TableMeasurement, b as TableRowFormatting, l as TableWidthType, m as TextEffect, T as TextFormatting, U as UnderlineStyle } from '../formatting-DFtuRFQY.mjs';
import { a as NumberingDefinitions } from '../lists-CyGxd5Y2.mjs';
export { A as AbstractNumbering, d as LevelSuffix, L as ListLevel, c as ListRendering, N as NumberFormat, b as NumberingInstance } from '../lists-CyGxd5Y2.mjs';
import { D as DocumentBody, F as Footnote, E as Endnote, y as HeaderFooter } from '../content-BZ9rYecc.mjs';
export { B as BlockContent, v as BlockSdt, z as BookmarkEnd, A as BookmarkStart, x as BreakContent, G as Column, C as Comment, f as CommentRangeEnd, g as CommentRangeStart, J as ComplexField, h as Deletion, K as DrawingContent, L as EndnotePosition, N as EndnoteProperties, O as Field, Q as FieldCharContent, U as FieldType, V as FooterReference, W as FootnotePosition, X as FootnoteProperties, Y as HeaderFooterType, Z as HeaderReference, H as Hyperlink, I as Image, _ as ImageCrop, $ as ImagePadding, a0 as ImagePosition, a1 as ImageSize, a2 as ImageTransform, a3 as ImageWrap, w as InlineSdt, i as Insertion, a4 as InstrTextContent, a5 as LineNumberRestart, a6 as MathEquation, M as MoveFrom, a7 as MoveFromRangeEnd, a8 as MoveFromRangeStart, j as MoveTo, a9 as MoveToRangeEnd, aa as MoveToRangeStart, ab as NoBreakHyphenContent, ac as NoteNumberRestart, ad as NoteRefMarkContent, ae as NoteReferenceContent, af as PageOrientation, P as Paragraph, k as ParagraphContent, b as ParagraphPropertyChange, ag as PropertyChangeInfo, l as Run, m as RunContent, R as RunPropertyChange, u as SdtDataBinding, t as SdtProperties, s as SdtType, ah as Section, S as SectionProperties, ai as SectionStart, aj as SeparatorContent, ak as Shape, al as ShapeContent, am as ShapeFill, an as ShapeOutline, ao as ShapeTextBody, ap as ShapeType, aq as SimpleField, ar as SoftHyphenContent, as as SymbolContent, at as TabContent, T as Table, n as TableCell, d as TableCellPropertyChange, c as TablePropertyChange, o as TableRow, e as TableRowPropertyChange, au as TableStructuralChangeInfo, av as TextBox, p as TextContent, q as TrackedChangeInfo, r as TrackedRunChange, aw as VerticalAlign } from '../content-BZ9rYecc.mjs';
export { D as DEFAULT_WATERMARK_PRESETS, P as PictureWatermark, T as TextWatermark, W as Watermark, p as pictureWatermarkDisplayEmu } from '../watermark-D90356ZM.mjs';
import { S as StyleDefinitions, T as Theme, F as FontTable, b as RelationshipMap, M as MediaFile } from '../styles-Diw0MASy.mjs';
export { D as DocDefaults, c as FontEmbed, d as FontInfo, R as Relationship, e as RelationshipType, a as Style, f as StyleType, g as ThemeColorScheme, h as ThemeFont, i as ThemeFontScheme } from '../styles-Diw0MASy.mjs';
import '../docx/wrapTypes.mjs';

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
    ['document']: DocumentBody;
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

export { type Document, DocumentBody, type DocumentSettings, type DocxPackage, Endnote, FontTable, Footnote, HeaderFooter, MediaFile, NumberingDefinitions, RelationshipMap, StyleDefinitions, Theme };
