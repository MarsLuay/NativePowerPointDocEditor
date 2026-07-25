import { Document } from './types/document.mjs';
import { D as DocumentBody, B as BlockContent, S as SectionProperties } from './content-BZ9rYecc.mjs';

/**
 * Document Serializer - Serialize complete document.xml
 *
 * Converts Document objects back to valid document.xml OOXML format.
 * Combines all content (paragraphs, tables) with section properties
 * and proper namespace declarations.
 *
 * OOXML Reference:
 * - Document root: w:document
 * - Document body: w:body
 * - Section properties: w:sectPr
 */

/**
 * Serialize a single block content item (paragraph, table, or block SDT).
 *
 * Exported so the footnote/endnote serializer can reuse the exact same
 * block-level emission the document body uses — preserving tracked-change
 * wrappers (`w:ins`/`w:del`), paragraph/run properties, fields, and tables
 * inside note bodies instead of reimplementing a minimal serializer.
 */
declare function serializeBlockContent(block: BlockContent): string;
/**
 * Serialize a DocumentBody to document.xml body content
 *
 * @param body - The document body to serialize
 * @returns XML string for the body element (without body tags)
 */
declare function serializeDocumentBody(body: DocumentBody): string;
/**
 * Serialize a complete Document to valid document.xml
 *
 * @param doc - The document to serialize
 * @returns Complete XML string for document.xml
 */
declare function serializeDocument(doc: Document): string;

/**
 * Section Properties Serializer - Serialize w:sectPr
 *
 * Converts SectionProperties back to OOXML `<w:sectPr>`. Used both for the
 * final `w:body/w:sectPr` (last section) and for mid-body section breaks
 * carried on a paragraph via `w:pPr/w:sectPr`.
 *
 * OOXML Reference:
 * - Section properties: w:sectPr (ECMA-376 §17.6.17)
 * - A paragraph-level sectPr (`w:pPr/w:sectPr`) marks the end of a section;
 *   the trailing `w:body/w:sectPr` describes the final section.
 */

/**
 * Serialize section properties (w:sectPr)
 */
declare function serializeSectionProperties(props: SectionProperties | undefined): string;

export { serializeDocument as a, serializeSectionProperties as b, serializeBlockContent as c, serializeDocumentBody as s };
