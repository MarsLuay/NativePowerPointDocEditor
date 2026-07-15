/**
 * DOCX XML Serializers
 *
 * Lower-level Document → OOXML transforms. For the round-trip "model in,
 * `.docx` archive out" path, use `./docx` instead.
 * @packageDocumentation
 * @public
 */
export { c as serializeBlockContent, a as serializeDocument, s as serializeDocumentBody, b as serializeSectionProperties } from '../../sectionPropertiesSerializer-Bp3ooY2t.mjs';
import { P as Paragraph, l as Run, T as Table, y as HeaderFooter, C as Comment, E as Endnote, F as Footnote } from '../../content-BZ9rYecc.mjs';
import '../../types/document.mjs';
import '../../colors-C3vA7HUU.mjs';
import '../../formatting-DFtuRFQY.mjs';
import '../../lists-CyGxd5Y2.mjs';
import '../../watermark-D90356ZM.mjs';
import '../../styles-BNjUANte.mjs';
import '../wrapTypes.mjs';

/**
 * Paragraph Serializer - Serialize paragraphs to OOXML XML
 *
 * Converts Paragraph objects back to <w:p> XML format for DOCX files.
 * Handles all paragraph properties and child content (runs, hyperlinks, fields, bookmarks).
 *
 * pPr property serializers (borders/shading/tabs/spacing/indentation/
 * numbering/frame) live in `paragraphSerializer/properties.ts`; child
 * content serializers (hyperlinks/fields/SDT/tracked-change wrappers)
 * live in `paragraphSerializer/content.ts`. This file orchestrates
 * paragraph-level serialization and re-exports the public API consumed
 * by sibling serializers.
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Runs, hyperlinks, bookmarks, fields as child elements
 */

/**
 * Serialize a paragraph to OOXML XML (w:p)
 *
 * @param paragraph - The paragraph to serialize
 * @returns XML string for the paragraph
 */
declare function serializeParagraph(paragraph: Paragraph): string;

/**
 * Run Serializer - Serialize runs to OOXML XML
 *
 * Converts Run objects back to <w:r> XML format for DOCX files.
 * Handles all formatting properties and content types.
 *
 * OOXML Reference:
 * - Run: w:r
 * - Run properties: w:rPr
 * - Text content: w:t
 */

/**
 * Serialize a run to OOXML XML (w:r)
 *
 * @param run - The run to serialize
 * @returns XML string for the run
 */
declare function serializeRun(run: Run): string;

/**
 * Table Serializer - Serialize tables to OOXML XML
 *
 * Converts Table objects back to <w:tbl> XML format for DOCX files.
 * Handles all table, row, and cell properties including merged cells.
 *
 * OOXML Reference:
 * - Table: w:tbl
 * - Table properties: w:tblPr
 * - Table grid: w:tblGrid
 * - Table row: w:tr
 * - Row properties: w:trPr
 * - Table cell: w:tc
 * - Cell properties: w:tcPr
 */

/**
 * Serialize a table to OOXML XML (w:tbl)
 *
 * @param table - The table to serialize
 * @returns XML string for the table
 */
declare function serializeTable(table: Table): string;

/**
 * Header/Footer Serializer - Serialize headers/footers to OOXML XML
 *
 * Converts HeaderFooter objects back to valid header*.xml / footer*.xml format.
 * Reuses paragraph and table serializers for content.
 *
 * OOXML Reference:
 * - Header root: w:hdr
 * - Footer root: w:ftr
 * - Content: w:p, w:tbl (same as document body)
 */

/**
 * Serialize a HeaderFooter object to valid OOXML XML
 *
 * @param hf - HeaderFooter object to serialize
 * @returns Complete XML string for header*.xml or footer*.xml
 */
declare function serializeHeaderFooter(hf: HeaderFooter): string;

/**
 * Comment Serializer
 *
 * Serializes Comment[] to OOXML comments.xml and commentsExtended.xml.
 *
 * comments.xml: the main comment content (w:comment elements)
 * commentsExtended.xml: reply threading via w15:commentEx with paraId/paraIdParent
 *
 * Each comment paragraph gets a w14:paraId. The last paragraph's paraId is used
 * in commentsExtended.xml to link replies to parents via w15:paraIdParent.
 */

/**
 * Serialize comments array to comments.xml content (backwards-compatible wrapper)
 */
declare function serializeComments(comments: Comment[]): string;

/**
 * Footnote / Endnote Serializer
 *
 * Serializes Footnote[] → word/footnotes.xml and Endnote[] → word/endnotes.xml.
 *
 * Unlike the comment serializer (which reimplements a minimal paragraph/run
 * emitter), note bodies are serialized with the SAME `serializeBlockContent`
 * the document body uses. That is deliberate: note bodies can carry the full
 * block model — paragraphs, tables, tracked-change wrappers (`w:ins`/`w:del`),
 * fields, run/paragraph properties — and reusing the body serializer preserves
 * all of it on round-trip rather than silently flattening it.
 *
 * Separator notes (`w:type="separator"` / `w:type="continuationSeparator"`) and
 * the in-body auto-number marks (`w:footnoteRef` / `w:endnoteRef`) survive
 * because the run model now carries them (see SeparatorContent / NoteRefMark-
 * Content in types/content/run.ts); no special-casing is needed here.
 *
 * OOXML Reference:
 * - Footnotes root: w:footnotes; each note: w:footnote[@w:id][@w:type]
 * - Endnotes root:  w:endnotes;  each note: w:endnote[@w:id][@w:type]
 */

/**
 * Serialize footnotes to word/footnotes.xml content.
 *
 * @param footnotes - All footnotes in document order, including separator notes.
 * @returns Complete footnotes.xml string.
 */
declare function serializeFootnotes(footnotes: Footnote[]): string;
/**
 * Serialize endnotes to word/endnotes.xml content.
 *
 * @param endnotes - All endnotes in document order, including separator notes.
 * @returns Complete endnotes.xml string.
 */
declare function serializeEndnotes(endnotes: Endnote[]): string;

export { serializeComments, serializeEndnotes, serializeFootnotes, serializeHeaderFooter, serializeParagraph, serializeRun, serializeTable };
