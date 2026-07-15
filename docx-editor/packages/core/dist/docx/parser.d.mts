/**
 * Main Parser Orchestrator - Unified parseDocx function
 *
 * Coordinates all sub-parsers to produce a complete Document model.
 * Handles loading order, dependency resolution, and font preloading.
 *
 * Parsing order:
 * 1. Unzip DOCX package
 * 2. Parse relationships
 * 3. Parse theme (needed for style color/font resolution)
 * 4. Parse styles (depends on theme)
 * 5. Parse numbering
 * 6. Parse document body (depends on styles, theme, numbering, rels)
 * 7. Parse headers/footers (depends on styles, theme, numbering, rels)
 * 8. Parse footnotes/endnotes (depends on styles, theme, numbering, rels)
 * 9. Extract and load fonts
 * 10. Build media file map
 * 11. Assemble final Document
 * @packageDocumentation
 * @public
 */
import { Document } from '../types/document.mjs';
import { D as DocxInput } from '../docxInput-DTbCa48g.mjs';
import '../colors-C3vA7HUU.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../lists-CyGxd5Y2.mjs';
import '../content-BZ9rYecc.mjs';
import './wrapTypes.mjs';
import '../watermark-D90356ZM.mjs';
import '../styles-BNjUANte.mjs';

/**
 * Main Parser Orchestrator - Unified parseDocx function
 *
 * Coordinates all sub-parsers to produce a complete Document model.
 * Handles loading order, dependency resolution, and font preloading.
 *
 * Parsing order:
 * 1. Unzip DOCX package
 * 2. Parse relationships
 * 3. Parse theme (needed for style color/font resolution)
 * 4. Parse styles (depends on theme)
 * 5. Parse numbering
 * 6. Parse document body (depends on styles, theme, numbering, rels)
 * 7. Parse headers/footers (depends on styles, theme, numbering, rels)
 * 8. Parse footnotes/endnotes (depends on styles, theme, numbering, rels)
 * 9. Extract and load fonts
 * 10. Build media file map
 * 11. Assemble final Document
 * @packageDocumentation
 * @public
 */

/**
 * Progress callback for tracking parsing stages
 */
type ProgressCallback = (stage: string, percent: number) => void;
/**
 * Parsing options
 */
interface ParseOptions {
    /** Progress callback for tracking parsing stages */
    onProgress?: ProgressCallback;
    /** Whether to preload fonts (default: true) */
    preloadFonts?: boolean;
    /** Whether to parse headers/footers (default: true) */
    parseHeadersFooters?: boolean;
    /** Whether to parse footnotes/endnotes (default: true) */
    parseNotes?: boolean;
    /** Whether to detect template variables (default: true) */
    detectVariables?: boolean;
}
/**
 * Parse a DOCX file into a complete Document model
 *
 * @param input - DOCX file as ArrayBuffer, Uint8Array, Blob, or File
 * @param options - Parsing options
 * @returns Promise resolving to Document
 * @throws Error if parsing fails
 */
declare function parseDocx(input: DocxInput, options?: ParseOptions): Promise<Document>;
/**
 * Quick parse - parse a DOCX without font loading
 * Useful for quick content extraction or when fonts aren't needed
 */
declare function quickParseDocx(buffer: ArrayBuffer): Promise<Document>;
/**
 * Full parse - parse everything including fonts
 */
declare function fullParseDocx(buffer: ArrayBuffer, onProgress?: ProgressCallback): Promise<Document>;
/**
 * Get template variables from a DOCX without full parsing
 * Faster than full parse when you only need variables
 */
declare function getDocxVariables(buffer: ArrayBuffer): Promise<string[]>;
/**
 * Get document summary without full parsing
 */
declare function getDocxSummary(buffer: ArrayBuffer): Promise<{
    hasDocument: boolean;
    hasStyles: boolean;
    hasTheme: boolean;
    hasNumbering: boolean;
    headerCount: number;
    footerCount: number;
    mediaCount: number;
    variableCount: number;
}>;

export { type ParseOptions, type ProgressCallback, fullParseDocx, getDocxSummary, getDocxVariables, parseDocx, quickParseDocx };
