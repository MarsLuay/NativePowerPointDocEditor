/**
 * Document Conversion Utilities
 *
 * Bidirectional conversion between Document (DOCX) and ProseMirror document.
 * @packageDocumentation
 * @public
 */
import { Node } from 'prosemirror-model';
import { Document } from '../../types/document.js';
import { S as StyleDefinitions, T as Theme } from '../../styles-2J4U-Lgk.js';
import { B as BlockContent } from '../../content-B8ScSBzC.js';
export { fromProseDoc, proseDocToBlocks, updateDocumentContent } from './fromProseDoc.js';
import '../../colors-C3vA7HUU.js';
import '../../formatting-JhqWT_XM.js';
import '../../lists-Bn29SzeS.js';
import '../../watermark-D90356ZM.js';
import '../../docx/wrapTypes.js';

/**
 * Document to ProseMirror Conversion
 *
 * Converts our Document type (from DOCX parsing) to a ProseMirror document.
 * Preserves all formatting attributes for round-trip fidelity.
 *
 * Style Resolution:
 * When styles are provided, paragraph properties are resolved from the style chain:
 * - Document defaults (docDefaults)
 * - Normal style (if no explicit styleId)
 * - Style chain (basedOn inheritance)
 * - Inline properties (highest priority)
 *
 * This file owns the top-level entry points (toProseDoc, headerFooterToProseDoc,
 * footnoteToProseDoc, createEmptyDoc). Per-domain converters live under
 * ./toProseDoc/ (marks, runs, paragraph, tables, textbox) — symmetric to
 * the fromProseDoc/ split.
 */

/**
 * Options for document conversion
 */
interface ToProseDocOptions {
    /** Style definitions for resolving paragraph styles */
    styles?: StyleDefinitions;
    /**
     * Doc-level `w:defaultTabStop` (§17.6.13) in twips, stamped onto the PM
     * doc node so `toFlowBlocks` picks it up. The body entry point reads
     * this from the parsed package; HF/footnote callers must pass it
     * through explicitly since their input is a content array, not a full
     * `Document`. Falls back to the OOXML default (720 twips) when null.
     */
    defaultTabStopTwips?: number | null;
}
/**
 * Convert a Document to a ProseMirror document
 *
 * @param document - The Document to convert
 * @param options - Conversion options including style definitions
 */
declare function toProseDoc(document: Document, options?: ToProseDocOptions): Node;
/**
 * Convert HeaderFooter content (array of Paragraph/Table blocks) to a ProseMirror document.
 * Used for editing headers/footers in their own ProseMirror editor and for the
 * unified header/footer render pipeline. `theme` must be threaded for themeColor
 * resolution in cell shading (`<w:shd w:themeFill=...>`) — without it, themed
 * fills in HF tables fall back to the unresolved theme key.
 */
declare function headerFooterToProseDoc(content: BlockContent[], options?: ToProseDocOptions & {
    theme?: Theme | null;
}): Node;
/**
 * Convert footnote/endnote content (array of Paragraph/Table blocks) to a
 * ProseMirror document. Mirrors `headerFooterToProseDoc` so footnotes flow
 * through the same body pipeline (toFlowBlocks → measureBlocks →
 * renderFragment) and inherit its block support — paragraph + table + image
 * + textBox + fields. Pre-PR, footnoteLayout's `convertFootnoteToContent`
 * re-implemented run/paragraph conversion by hand and silently dropped
 * tables, images, and fields nested inside a footnote.
 */
declare function footnoteToProseDoc(content: BlockContent[], options?: ToProseDocOptions & {
    theme?: Theme | null;
}): Node;
/**
 * Create an empty ProseMirror document
 */
declare function createEmptyDoc(): Node;

export { type ToProseDocOptions, createEmptyDoc, footnoteToProseDoc, headerFooterToProseDoc, toProseDoc };
