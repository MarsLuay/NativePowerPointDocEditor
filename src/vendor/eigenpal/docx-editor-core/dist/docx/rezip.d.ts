/**
 * DOCX Repacker - Repack modified document into valid DOCX
 *
 * Takes a Document with modified content and creates a new DOCX file
 * by updating document.xml while preserving all other files from
 * the original ZIP archive.
 *
 * This ensures round-trip fidelity:
 * - styles.xml, theme1.xml, fontTable.xml remain untouched
 * - Media files preserved
 * - Relationships preserved
 * - Only document.xml is updated with new content
 *
 * OOXML Package Structure:
 * - [Content_Types].xml - Content type declarations
 * - _rels/.rels - Package relationships
 * - word/document.xml - Main document (modified)
 * - word/styles.xml - Styles (preserved)
 * - word/theme/theme1.xml - Theme (preserved)
 * - word/numbering.xml - Numbering (preserved)
 * - word/fontTable.xml - Font table (preserved)
 * - word/settings.xml - Settings (preserved)
 * - word/header*.xml - Headers (preserved)
 * - word/footer*.xml - Footers (preserved)
 * - word/footnotes.xml - Footnotes (preserved)
 * - word/endnotes.xml - Endnotes (preserved)
 * - word/media/* - Media files (preserved)
 * - word/_rels/document.xml.rels - Document relationships (preserved)
 * - docProps/* - Document properties (preserved)
 *
 * Orchestrators (repackDocx, selective updates, validation, create-empty)
 * live here. Per-domain helpers — part enumeration, new-image registration,
 * new-hyperlink registration, header/footer & comment packaging, and the
 * empty-DOCX template — live under ./rezip/.
 * @packageDocumentation
 * @public
 */
import JSZip from 'jszip';
import { Document } from '../types/document.js';
import '../colors-C3vA7HUU.js';
import '../formatting-JhqWT_XM.js';
import '../lists-Bn29SzeS.js';
import '../content-B8ScSBzC.js';
import './wrapTypes.js';
import '../watermark-D90356ZM.js';
import '../styles-BGGjYqnf.js';

/**
 * DOCX Unzipper
 *
 * Extracts all files from a DOCX ZIP archive and organizes them
 * into a structured format for further processing.
 *
 * A DOCX file is a ZIP archive containing:
 * - [Content_Types].xml - Content type declarations
 * - word/document.xml - Main document content
 * - word/styles.xml - Style definitions
 * - word/theme/theme1.xml - Theme colors and fonts
 * - word/numbering.xml - List/numbering definitions
 * - word/fontTable.xml - Font declarations
 * - word/settings.xml - Document settings
 * - word/webSettings.xml - Web settings
 * - word/header*.xml - Header content
 * - word/footer*.xml - Footer content
 * - word/footnotes.xml - Footnotes
 * - word/endnotes.xml - Endnotes
 * - word/media/* - Embedded images and media
 * - word/_rels/document.xml.rels - Relationships
 * - _rels/.rels - Package relationships
 * - docProps/core.xml - Core properties
 * - docProps/app.xml - Application properties
 */

/**
 * Raw extracted content from a DOCX file
 */
interface RawDocxContent {
    documentXml: string | null;
    stylesXml: string | null;
    themeXml: string | null;
    numberingXml: string | null;
    fontTableXml: string | null;
    settingsXml: string | null;
    webSettingsXml: string | null;
    headers: Map<string, string>;
    footers: Map<string, string>;
    footnotesXml: string | null;
    endnotesXml: string | null;
    commentsXml: string | null;
    commentsExtensibleXml: string | null;
    commentsExtendedXml: string | null;
    documentRels: string | null;
    packageRels: string | null;
    contentTypesXml: string | null;
    corePropsXml: string | null;
    appPropsXml: string | null;
    customPropsXml: string | null;
    media: Map<string, ArrayBuffer>;
    fonts: Map<string, ArrayBuffer>;
    allXml: Map<string, string>;
    originalZip: JSZip;
    originalBuffer: ArrayBuffer;
}

/**
 * Part Enumeration
 *
 * Shared helpers for walking the parts of a DOCX package (body, headers,
 * footers, footnotes, endnotes) when registering newly inserted images and
 * hyperlinks against each part's rels file.
 */

/**
 * Find the highest rId number in a relationships XML string.
 */
declare function findMaxRId(relsXml: string): number;

/**
 * Header/Footer & Comment Packaging
 *
 * Serialize modified headers, footers, and comments back into the ZIP and
 * ensure each new part is registered in `[Content_Types].xml` and the
 * document's rels file. Without this step Word silently drops parts that
 * the editor inserted into a previously-blank document (#274).
 */

declare const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
declare const COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
declare const COMMENTS_IDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
declare const COMMENTS_EXTENSIBLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";
/**
 * Collect serialized header/footer XML updates from the document model.
 * Uses the relationship map to resolve rId → filename.
 */
declare function collectHeaderFooterUpdates(doc: Document): Map<string, string>;

/**
 * Empty DOCX Template
 *
 * Produces a minimal valid DOCX package — Content Types, package + document
 * rels, an empty body with the standard letter-size section, default
 * Calibri 11 styles, and core/app properties. Used as the base for
 * `createDocx()` when the caller has no existing buffer.
 */
/**
 * Create a new empty DOCX file.
 *
 * @returns Promise resolving to minimal DOCX as ArrayBuffer
 */
declare function createEmptyDocx(): Promise<ArrayBuffer>;

/**
 * DOCX Repacker - Repack modified document into valid DOCX
 *
 * Takes a Document with modified content and creates a new DOCX file
 * by updating document.xml while preserving all other files from
 * the original ZIP archive.
 *
 * This ensures round-trip fidelity:
 * - styles.xml, theme1.xml, fontTable.xml remain untouched
 * - Media files preserved
 * - Relationships preserved
 * - Only document.xml is updated with new content
 *
 * OOXML Package Structure:
 * - [Content_Types].xml - Content type declarations
 * - _rels/.rels - Package relationships
 * - word/document.xml - Main document (modified)
 * - word/styles.xml - Styles (preserved)
 * - word/theme/theme1.xml - Theme (preserved)
 * - word/numbering.xml - Numbering (preserved)
 * - word/fontTable.xml - Font table (preserved)
 * - word/settings.xml - Settings (preserved)
 * - word/header*.xml - Headers (preserved)
 * - word/footer*.xml - Footers (preserved)
 * - word/footnotes.xml - Footnotes (preserved)
 * - word/endnotes.xml - Endnotes (preserved)
 * - word/media/* - Media files (preserved)
 * - word/_rels/document.xml.rels - Document relationships (preserved)
 * - docProps/* - Document properties (preserved)
 *
 * Orchestrators (repackDocx, selective updates, validation, create-empty)
 * live here. Per-domain helpers — part enumeration, new-image registration,
 * new-hyperlink registration, header/footer & comment packaging, and the
 * empty-DOCX template — live under ./rezip/.
 * @packageDocumentation
 * @public
 */

/**
 * Options for repacking DOCX
 */
interface RepackOptions {
    /** Compression level (0-9, default: 6) */
    compressionLevel?: number;
    /** Whether to update modification date in docProps/core.xml */
    updateModifiedDate?: boolean;
    /** Custom modifier name for lastModifiedBy */
    modifiedBy?: string;
}
/**
 * Repack a Document into a valid DOCX file
 *
 * @param doc - Document with modified content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 * @throws Error if document has no original buffer for round-trip
 */
declare function repackDocx(doc: Document, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Repack a Document using raw content for more control
 *
 * @param doc - Document with modified content
 * @param rawContent - Original raw content from unzipDocx
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
declare function repackDocxFromRaw(doc: Document, rawContent: RawDocxContent, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Update only document.xml in a DOCX buffer (minimal changes)
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param newDocumentXml - New document.xml content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
declare function updateDocumentXml(originalBuffer: ArrayBuffer, newDocumentXml: string, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Update a specific XML file in a DOCX buffer
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param path - Path within the ZIP (e.g., "word/styles.xml")
 * @param content - New XML content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
declare function updateXmlFile(originalBuffer: ArrayBuffer, path: string, content: string, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Update multiple files in a DOCX buffer
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param updates - Map of path -> content for files to update
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
declare function updateMultipleFiles(originalBuffer: ArrayBuffer, updates: Map<string, string | ArrayBuffer>, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Apply file updates to an already-loaded JSZip instance and generate the output.
 * Use this when the zip is already loaded to avoid a redundant decompression pass.
 */
declare function applyUpdatesToZip(zip: JSZip, updates: Map<string, string | ArrayBuffer>, options?: RepackOptions): Promise<ArrayBuffer>;
/**
 * Add a new relationship to document.xml.rels
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param relationship - New relationship to add
 * @returns Promise resolving to { buffer: ArrayBuffer, rId: string }
 */
declare function addRelationship(originalBuffer: ArrayBuffer, relationship: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
}): Promise<{
    buffer: ArrayBuffer;
    rId: string;
}>;
/**
 * Add a media file to the DOCX
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param filename - Filename for the media (e.g., "image1.png")
 * @param data - Binary data for the media file
 * @param mimeType - MIME type (e.g., "image/png")
 * @returns Promise resolving to { buffer: ArrayBuffer, rId: string, path: string }
 */
declare function addMedia(originalBuffer: ArrayBuffer, filename: string, data: ArrayBuffer, mimeType: string): Promise<{
    buffer: ArrayBuffer;
    rId: string;
    path: string;
}>;
/**
 * Update core properties XML with new modification date
 */
declare function updateCoreProperties(corePropsXml: string, options: {
    updateModifiedDate?: boolean;
    modifiedBy?: string;
}): string;
/**
 * Validate that a buffer is a valid DOCX file
 *
 * @param buffer - Buffer to validate
 * @returns Promise resolving to validation result
 */
declare function validateDocx(buffer: ArrayBuffer): Promise<{
    valid: boolean;
    errors: string[];
    warnings: string[];
}>;
/**
 * Check if buffer looks like a DOCX file (quick check)
 *
 * @param buffer - Buffer to check
 * @returns true if buffer starts with ZIP signature
 */
declare function isDocxBuffer(buffer: ArrayBuffer): boolean;
/**
 * Create a new DOCX from a Document (without requiring original buffer)
 *
 * @param doc - Document to serialize
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
declare function createDocx(doc: Document): Promise<ArrayBuffer>;

export { COMMENTS_CONTENT_TYPE, COMMENTS_EXTENDED_CONTENT_TYPE, COMMENTS_EXTENSIBLE_CONTENT_TYPE, COMMENTS_IDS_CONTENT_TYPE, type RepackOptions, addMedia, addRelationship, applyUpdatesToZip, collectHeaderFooterUpdates, createDocx, createEmptyDocx, repackDocx as default, findMaxRId, isDocxBuffer, repackDocx, repackDocxFromRaw, updateCoreProperties, updateDocumentXml, updateMultipleFiles, updateXmlFile, validateDocx };
