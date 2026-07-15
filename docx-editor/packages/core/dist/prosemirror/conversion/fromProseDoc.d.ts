/**
 * ProseMirror to Document Conversion
 *
 * Converts a ProseMirror document back to our Document type.
 * This enables round-trip editing: DOCX -> Document -> PM -> Document -> DOCX
 *
 * Key responsibilities:
 * - Coalesce consecutive text with same marks into single Runs
 * - Preserve paragraph attributes (paraId, textId, formatting)
 * - Handle marks -> TextFormatting conversion
 *
 * This file owns the top-level orchestrator (`fromProseDoc`) plus block
 * extraction and the page-break paragraph factory. Per-domain converters
 * live under ./fromProseDoc/ (marks, runs, paragraph, tables, textbox).
 * The deep import `@eigenpal/.../prosemirror/conversion/fromProseDoc` is
 * a tsup entry consumed by the Vue adapter — the barrel re-exports
 * preserve that surface.
 * @packageDocumentation
 * @public
 */
import { Node } from 'prosemirror-model';
import { Document } from '../../types/document.js';
import { B as BlockContent } from '../../content-B8ScSBzC.js';
import '../../colors-C3vA7HUU.js';
import '../../formatting-JhqWT_XM.js';
import '../../lists-Bn29SzeS.js';
import '../../watermark-D90356ZM.js';
import '../../styles-2J4U-Lgk.js';
import '../../docx/wrapTypes.js';

/**
 * ProseMirror to Document Conversion
 *
 * Converts a ProseMirror document back to our Document type.
 * This enables round-trip editing: DOCX -> Document -> PM -> Document -> DOCX
 *
 * Key responsibilities:
 * - Coalesce consecutive text with same marks into single Runs
 * - Preserve paragraph attributes (paraId, textId, formatting)
 * - Handle marks -> TextFormatting conversion
 *
 * This file owns the top-level orchestrator (`fromProseDoc`) plus block
 * extraction and the page-break paragraph factory. Per-domain converters
 * live under ./fromProseDoc/ (marks, runs, paragraph, tables, textbox).
 * The deep import `@eigenpal/.../prosemirror/conversion/fromProseDoc` is
 * a tsup entry consumed by the Vue adapter — the barrel re-exports
 * preserve that surface.
 * @packageDocumentation
 * @public
 */

/**
 * Convert a ProseMirror document to our Document type
 */
declare function fromProseDoc(pmDoc: Node, baseDocument?: Document): Document;
/**
 * Update a Document with content from a ProseMirror document
 * Preserves all non-content parts of the original document
 */
declare function updateDocumentContent(originalDocument: Document, pmDoc: Node): Document;
/**
 * Convert a ProseMirror document back to an array of Paragraph/Table blocks.
 * Used for converting edited header/footer PM content back to the document model.
 */
declare function proseDocToBlocks(pmDoc: Node): BlockContent[];

export { fromProseDoc, proseDocToBlocks, updateDocumentContent };
