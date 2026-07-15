import { Document } from './types/document.mjs';
import { x as BreakContent, P as Paragraph, l as Run, m as RunContent } from './content-BZ9rYecc.mjs';

/**
 * Insert Operations Utility
 *
 * Utility functions for inserting content into the document.
 * Provides functions for inserting page breaks, horizontal rules, and other elements.
 */

/**
 * Insert position in the document
 */
interface InsertPosition {
    /** Paragraph index in the document body */
    paragraphIndex: number;
    /** Run index within the paragraph (optional) */
    runIndex?: number;
    /** Character offset within the run (optional) */
    offset?: number;
}
/**
 * Create a page break content element
 */
declare function createPageBreak(): BreakContent;
/**
 * Create a column break content element
 */
declare function createColumnBreak(): BreakContent;
/**
 * Create a text wrapping break (line break)
 */
declare function createLineBreak(clear?: 'none' | 'left' | 'right' | 'all'): BreakContent;
/**
 * Create a run containing a page break
 */
declare function createPageBreakRun(): Run;
/**
 * Create an empty paragraph with a page break before it
 */
declare function createPageBreakParagraph(): Paragraph;
/**
 * Insert a page break at a position in the document
 * This inserts a new paragraph with pageBreakBefore: true
 */
declare function insertPageBreak(doc: Document, position: InsertPosition): Document;
/**
 * Create a horizontal rule paragraph
 * Uses a paragraph with bottom border to simulate horizontal rule
 */
declare function createHorizontalRule(): Paragraph;
/**
 * Insert a horizontal rule at a position in the document
 */
declare function insertHorizontalRule(doc: Document, position: InsertPosition): Document;
/**
 * Check if content is a page break
 */
declare function isPageBreak(content: RunContent): boolean;
/**
 * Check if content is a column break
 */
declare function isColumnBreak(content: RunContent): boolean;
/**
 * Check if content is a line break
 */
declare function isLineBreak(content: RunContent): boolean;
/**
 * Check if content is any type of break
 */
declare function isBreakContent(content: RunContent): content is BreakContent;
/**
 * Check if a paragraph has pageBreakBefore
 */
declare function hasPageBreakBefore(paragraph: Paragraph): boolean;
/**
 * Count page breaks in a document
 */
declare function countPageBreaks(doc: Document): number;
/**
 * Find all page break positions in a document
 */
declare function findPageBreaks(doc: Document): InsertPosition[];
/**
 * Remove a page break at a specific position
 */
declare function removePageBreak(doc: Document, position: InsertPosition): Document;

/**
 * Google Fonts Loader
 *
 * Dynamically loads fonts from Google Fonts API with:
 * - Loading state tracking
 * - Duplicate prevention
 * - Callback notifications
 * - Font availability detection
 */
/**
 * Load a font from Google Fonts
 *
 * @param fontFamily - The font family name to load
 * @param options - Optional configuration
 * @returns Promise resolving to true if font loaded successfully, false otherwise
 */
declare function loadFont(fontFamily: string, options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
}): Promise<boolean>;
/**
 * Load multiple fonts from Google Fonts
 *
 * @param families - Array of font family names to load
 * @param options - Optional configuration
 * @returns Promise resolving when all fonts are loaded (or failed)
 */
declare function loadFonts(families: string[], options?: {
    weights?: number[];
    styles?: ('normal' | 'italic')[];
}): Promise<void>;
/**
 * Check if a font is loaded
 *
 * @param fontFamily - The font family name to check
 * @returns true if the font is loaded, false otherwise
 */
declare function isFontLoaded(fontFamily: string): boolean;
/**
 * Enable or disable the editor's automatic Google Fonts lookup.
 *
 * Defaults to enabled. Set to `false` in no-egress embedders (strict CSP,
 * offline) so `loadFont` / `loadFontWithMapping` never inject a
 * `fonts.googleapis.com` stylesheet `<link>`. This gates ONLY the implicit
 * Google Fonts lookup — font URLs you register yourself (the `fonts` prop /
 * `loadFontFromUrl`) are still fetched by the browser, and embedded blobs
 * (`loadFontFromBuffer`) and system fonts still resolve.
 *
 * The flag is page-global (module-level), not per-editor: with multiple
 * editors on one page the last caller wins. Call it before loading documents;
 * disabling does not cancel fetches already in flight. A document font that
 * is neither local nor registered then resolves `loadFont` to `false`
 * silently (no `onFontError`) and renders via its CSS fallback stack.
 *
 * @see isGoogleFontsEnabled
 * @see loadFontFromBuffer
 * @see loadFontFromUrl
 *
 * @public
 */
declare function setGoogleFontsEnabled(enabled: boolean): void;
/**
 * Whether the automatic Google Fonts lookup is currently enabled.
 *
 * @see setGoogleFontsEnabled
 *
 * @public
 */
declare function isGoogleFontsEnabled(): boolean;
/**
 * Check if any fonts are currently loading
 *
 * @returns true if any fonts are loading, false otherwise
 */
declare function isLoading(): boolean;
/**
 * Get list of all loaded fonts
 *
 * @returns Array of loaded font family names
 */
declare function getLoadedFonts(): string[];
/**
 * Register a callback to be notified when fonts are loaded
 *
 * @param callback - Function to call when fonts are loaded
 * @returns Cleanup function to remove the callback
 */
declare function onFontsLoaded(callback: (fonts: string[]) => void): () => void;
/**
 * Register a callback to be notified when a font fails to load.
 *
 * Adapters subscribe and forward to their `onError` prop. Returns the unsub.
 *
 * @public
 */
declare function onFontError(callback: (error: Error) => void): () => void;
/**
 * Check if a font is available on the system using canvas measurement
 *
 * Compares text width with the target font vs a known fallback font (and the
 * opposite fallback, for names that collide with the browser defaults). If
 * the widths differ, the font is available. Reuses one shared canvas; the
 * fallback widths are session constants and memoized.
 *
 * @param fontFamily - The font family name to check
 * @param fallbackFont - Fallback font to compare against
 * @returns true if font is available, false otherwise
 */
declare function canRenderFont(fontFamily: string, fallbackFont?: string): boolean;
/**
 * Load a font from a raw buffer (e.g., embedded in DOCX)
 *
 * Call before loading the document when possible: registration marks the
 * family as ours, which is what keeps the metric-compatible Google fallback
 * in play for subsetted faces during the document's font resolution pass.
 * Late registration still self-heals on the next pass. Under
 * `setGoogleFontsEnabled(false)` that fallback fetch is suppressed and
 * glyphs missing from the registered faces render via the CSS stack.
 *
 * @param fontFamily - The font family name
 * @param buffer - Font file buffer (TTF, OTF, WOFF, WOFF2)
 * @param options - Font options
 * @returns Promise resolving when font is loaded
 */
declare function loadFontFromBuffer(fontFamily: string, buffer: ArrayBuffer, options?: {
    weight?: number | string;
    style?: 'normal' | 'italic';
}): Promise<boolean>;
/**
 * Load a font face from a URL (woff2, woff, ttf, otf).
 *
 * Injects an `@font-face` rule pointing at the URL. Multiple weights of the
 * same family can be registered independently. Families registered here are
 * treated as potentially subsetted: the editor still fetches their
 * metric-compatible Google equivalent as a glyph-coverage fallback, unless
 * disabled via `setGoogleFontsEnabled(false)`.
 *
 * @param fontFamily - CSS font-family name to expose
 * @param src - URL to the font file
 * @param options - Optional weight
 * @returns Promise resolving to true if the face became available
 *
 * @public
 */
declare function loadFontFromUrl(fontFamily: string, src: string, options?: {
    weight?: number | string;
}): Promise<boolean>;
/**
 * Declarative description of a single font face to register with the editor.
 *
 * Each entry injects one `@font-face` rule pointing at a URL. Multiple
 * entries can share `family` to register distinct weights as separate faces.
 *
 * For Google Fonts, call `loadFont(family)` directly — the `fonts` prop is
 * for fonts the consumer hosts themselves. For raw bytes already in memory
 * (DOCX-embedded fonts, user uploads), call `loadFontFromBuffer(family, buf)`.
 *
 * @public
 */
interface FontDefinition {
    /**
     * CSS `font-family` name to expose. Match the family name your documents
     * reference; the browser uses this to look up glyphs when text is rendered.
     */
    family: string;
    /**
     * URL to the font file (woff2, woff, ttf, or otf). The loader injects an
     * `@font-face` rule and lets the browser fetch on demand.
     */
    src: string;
    /**
     * CSS `font-weight` for this face. Defaults to `'normal'` (≈400). Pass a
     * number (`400`, `700`) or a CSS keyword (`'bold'`). Required when one
     * `family` registers multiple weights as separate entries.
     */
    weight?: number | string;
}
/**
 * Register a list of custom font faces. Used by the `fonts` prop on
 * `<DocxEditor>` (React + Vue). Idempotent — safe to call on every render.
 *
 * @public
 */
declare function loadFontDefinitions(defs: ReadonlyArray<FontDefinition> | undefined): Promise<void>;
/**
 * Mapping from common Office/system fonts to Google Fonts equivalents
 *
 * Google Fonts doesn't have exact matches for many Microsoft fonts,
 * but these are close alternatives that work well for document rendering.
 */
declare const FONT_MAPPING: Record<string, string>;
/**
 * Get the Google Fonts equivalent for a font name
 *
 * @param fontName - The original font name from the document
 * @returns The Google Fonts equivalent, or the original name if no mapping exists
 */
declare function getGoogleFontEquivalent(fontName: string): string;
/**
 * Load a font, automatically mapping to Google Fonts equivalent if needed.
 * If the font needs mapping, also creates a CSS alias so the original font
 * name works in stylesheets.
 *
 * @param fontFamily - The font family name (may be an Office font)
 * @returns Promise resolving to true if font loaded
 */
declare function loadFontWithMapping(fontFamily: string): Promise<boolean>;
/**
 * Load multiple fonts with automatic mapping to Google Fonts equivalents
 *
 * @param families - Array of font family names
 * @returns Promise resolving when all fonts are loaded
 */
declare function loadFontsWithMapping(families: string[]): Promise<void>;
/**
 * Preload a list of common document fonts
 *
 * This preloads fonts commonly used in DOCX documents that have
 * Google Fonts equivalents.
 */
declare function preloadCommonFonts(): Promise<void>;
/**
 * Extract all font families used in a document
 *
 * Uses loose typing to handle any document-like structure.
 *
 * @param document - The parsed document
 * @returns Set of unique font family names
 */
declare function extractFontsFromDocument(document: unknown): Set<string>;
/**
 * Extract fonts from a document and load them from Google Fonts
 *
 * @param document - The parsed document
 * @returns Promise resolving when fonts are loaded
 */
declare function loadDocumentFonts(document: unknown): Promise<void>;

export { setGoogleFontsEnabled as A, type FontDefinition as B, extractFontsFromDocument as C, getGoogleFontEquivalent as D, loadDocumentFonts as E, FONT_MAPPING as F, loadFontDefinitions as G, loadFontFromUrl as H, type InsertPosition as I, loadFontWithMapping as J, loadFontsWithMapping as K, onFontError as L, countPageBreaks as a, createColumnBreak as b, canRenderFont as c, createHorizontalRule as d, createLineBreak as e, createPageBreak as f, createPageBreakParagraph as g, createPageBreakRun as h, findPageBreaks as i, getLoadedFonts as j, hasPageBreakBefore as k, insertHorizontalRule as l, insertPageBreak as m, isBreakContent as n, isColumnBreak as o, isFontLoaded as p, isLoading as q, isGoogleFontsEnabled as r, isLineBreak as s, isPageBreak as t, loadFont as u, loadFontFromBuffer as v, loadFonts as w, onFontsLoaded as x, preloadCommonFonts as y, removePageBreak as z };
