/**
 * Editor utilities (curated public surface).
 *
 * The named exports below are the public API contract. Adding a helper
 * to a source module does not automatically make it public — it must
 * be added to this barrel to be reachable from `@npde/docx-editor-core/utils`.
 * @packageDocumentation
 * @public
 */
export { PIXELS_PER_INCH, TWIPS_PER_INCH, clamp, eighthsToPixels, emuToPixels, emuToTwips, formatPx, halfPointsToPixels, halfPointsToPoints, pixelsToEmu, pixelsToTwips, pointsToHalfPoints, pointsToPixels, roundPixels, twipsToEmu, twipsToPixels } from './units.js';
export { C as CreateEmptyDocumentOptions, P as ProcessTemplateOptions, a as ProcessTemplateResult, T as TemplateError, z as ThemeMatrixCell, b as blendColors, c as colorsEqual, d as createDocumentWithText, e as createEmptyDocument, f as createRgbColor, g as createTemplateProcessor, h as createThemeColor, i as darkenColor, A as ensureHexPrefix, B as generateThemeTintShadeMatrix, j as getContrastingColor, k as getMissingVariables, l as getTemplateTags, D as getThemeTintShadeHex, m as isBlack, n as isWhite, o as lightenColor, p as parseColorString, q as previewTemplate, r as processTemplate, s as processTemplateAdvanced, F as processTemplateAndDownload, t as processTemplateAsBlob, u as processTemplateDetailed, v as resolveColor, G as resolveColorToHex, w as resolveHighlightColor, E as resolveHighlightToCss, x as resolveShadingColor, y as validateTemplate } from '../colorResolver-D9T-sBKo.js';
import { l as Run, P as Paragraph } from '../content-B8ScSBzC.js';
import { T as Theme, F as FontTable } from '../styles-2J4U-Lgk.js';
export { D as DocxInput, t as toArrayBuffer } from '../docxInput-DTbCa48g.js';
export { F as FONT_MAPPING, B as FontDefinition, I as InsertPosition, c as canRenderFont, a as countPageBreaks, b as createColumnBreak, d as createHorizontalRule, e as createLineBreak, f as createPageBreak, g as createPageBreakParagraph, h as createPageBreakRun, C as extractFontsFromDocument, i as findPageBreaks, D as getGoogleFontEquivalent, j as getLoadedFonts, k as hasPageBreakBefore, l as insertHorizontalRule, m as insertPageBreak, n as isBreakContent, o as isColumnBreak, p as isFontLoaded, r as isGoogleFontsEnabled, s as isLineBreak, q as isLoading, t as isPageBreak, E as loadDocumentFonts, u as loadFont, G as loadFontDefinitions, v as loadFontFromBuffer, H as loadFontFromUrl, J as loadFontWithMapping, w as loadFonts, K as loadFontsWithMapping, L as onFontError, x as onFontsLoaded, y as preloadCommonFonts, z as removePageBreak, A as setGoogleFontsEnabled } from '../fontLoader-C2_QwHvP.js';
import { FontOption } from './fontOptions.js';
import { Document } from '../types/document.js';
import { B as BorderSpec, S as ShadingProperties } from '../colors-C3vA7HUU.js';
import { P as ParagraphFormatting, T as TextFormatting } from '../formatting-JhqWT_XM.js';
export { HeadingInfo, collectHeadings } from './headingCollector.js';
export { WordSelectionResult, createDoubleClickWordSelector, createTripleClickParagraphSelector, expandSelectionToWordBoundaries, findWordAt, findWordBoundaries, getWordAt, handleClickForMultiClick, selectParagraphAtCursor, selectWordAtCursor, selectWordInTextNode } from './textSelection.js';
export { MIN_CARD_GAP, SIDEBAR_DOCUMENT_SHIFT, SIDEBAR_PAGE_GAP, SIDEBAR_WIDTH } from './sidebarConstants.js';
import '../docx/wrapTypes.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';
import 'prosemirror-model';

/**
 * Clipboard utilities for copy/paste with formatting
 *
 * Handles:
 * - Copy: puts formatted HTML and plain text on clipboard
 * - Paste: reads HTML clipboard, converts to runs with formatting
 * - Handles paste from Word (cleans up Word HTML)
 * - Ctrl+C, Ctrl+V, Ctrl+X keyboard shortcuts
 */

/**
 * Clipboard content format
 */
interface ClipboardContent {
    /** Plain text representation */
    plainText: string;
    /** HTML representation */
    html: string;
    /** Internal format (JSON) for preserving full formatting */
    internal?: string;
}
/**
 * Parsed clipboard content
 */
interface ParsedClipboardContent {
    /** Runs parsed from clipboard */
    runs: Run[];
    /** Whether content came from Word */
    fromWord: boolean;
    /** Whether content came from our editor */
    fromEditor: boolean;
    /** Original plain text */
    plainText: string;
}
/**
 * Options for clipboard operations
 */
interface ClipboardOptions {
    /** Whether to include formatting in copy */
    includeFormatting?: boolean;
    /** Whether to clean Word-specific formatting */
    cleanWordFormatting?: boolean;
    /** Callback for handling errors */
    onError?: (error: Error) => void;
    /** Document theme — required to resolve themed text/shading colors in HTML. */
    theme?: Theme | null;
}
/**
 * Custom MIME type for internal clipboard format
 */
declare const INTERNAL_CLIPBOARD_TYPE = "application/x-docx-editor";
/**
 * Standard clipboard MIME types
 */
declare const CLIPBOARD_TYPES: {
    readonly HTML: "text/html";
    readonly PLAIN: "text/plain";
};
/**
 * Extract image files from clipboard data (if present).
 */
declare function getClipboardImageFiles(clipboardData: DataTransfer | null): File[];
/**
 * Copy runs to clipboard with formatting
 */
declare function copyRuns(runs: Run[], options?: ClipboardOptions): Promise<boolean>;
/**
 * Copy paragraphs to clipboard with formatting
 */
declare function copyParagraphs(paragraphs: Paragraph[], options?: ClipboardOptions): Promise<boolean>;
/**
 * Convert runs to clipboard content (HTML and plain text).
 *
 * @param theme - Optional document theme. Pass it so themed text color and
 *   shading resolve correctly in the HTML payload (matters when pasting into
 *   Gmail/Word/etc.). Omit for rgb-only content.
 */
declare function runsToClipboardContent(runs: Run[], includeFormatting?: boolean, theme?: Theme | null): ClipboardContent;
/**
 * Convert paragraphs to clipboard content.
 *
 * @param theme - See {@link runsToClipboardContent}.
 */
declare function paragraphsToClipboardContent(paragraphs: Paragraph[], includeFormatting?: boolean, theme?: Theme | null): ClipboardContent;
/**
 * Write content to clipboard
 */
declare function writeToClipboard(content: ClipboardContent): Promise<boolean>;
/**
 * Read content from clipboard
 */
declare function readFromClipboard(options?: ClipboardOptions): Promise<ParsedClipboardContent | null>;
/**
 * Handle paste event
 */
declare function handlePasteEvent(event: ClipboardEvent, options?: ClipboardOptions): ParsedClipboardContent | null;
/**
 * Parse HTML from clipboard
 */
declare function parseClipboardHtml(html: string, plainText: string, cleanWordFormatting?: boolean): ParsedClipboardContent;
/**
 * Check if HTML is from Microsoft Word
 */
declare function isWordHtml(html: string): boolean;
/**
 * Check if HTML is from our editor
 */
declare function isEditorHtml(html: string): boolean;
/**
 * Clean Microsoft Word HTML
 */
declare function cleanWordHtml(html: string): string;
/**
 * Convert HTML to runs
 */
declare function htmlToRuns(html: string, plainTextFallback: string): Run[];
/**
 * Create clipboard keyboard handlers for an editor
 */
declare function createClipboardHandlers(options: {
    onCopy?: () => {
        runs: Run[];
    } | null;
    onCut?: () => {
        runs: Run[];
    } | null;
    onPaste?: (content: ParsedClipboardContent) => void;
    clipboardOptions?: ClipboardOptions;
}): {
    handleCopy: (event: ClipboardEvent) => Promise<void>;
    handleCut: (event: ClipboardEvent) => Promise<void>;
    handlePaste: (event: ClipboardEvent) => void;
    handleKeyDown: (event: KeyboardEvent) => Promise<void>;
};

/**
 * Embedded font de-obfuscation (ECMA-376 Part 4 §2.8.1, "Font Embedding").
 *
 * Word stores embedded fonts as obfuscated OpenType (`.odttf`): the first 32
 * bytes of the font binary are XOR-scrambled with a 16-byte key, the rest of
 * the file is untouched. The key is the font's `w:fontKey` GUID with its byte
 * order REVERSED, applied once to bytes 0-15 and again to bytes 16-31.
 *
 * Example (from the spec): GUID `001B70DC-AA60-4AD5-90EC-18A0948E1EAE` yields
 * the key bytes `AE 1E 8E 94 A0 18 EC 90 D5 4A 60 AA DC 70 1B 00`.
 *
 * The scheme is a pure XOR, so the same operation obfuscates and de-obfuscates.
 */
/**
 * Whether a string is a usable embedded-font obfuscation key (a 128-bit GUID,
 * with or without braces/hyphens).
 *
 * @public
 */
declare function isValidFontKey(fontKey: string | undefined | null): boolean;
/**
 * De-obfuscate an embedded `.odttf` font into a usable OpenType/TrueType
 * binary by XOR-ing its first 32 bytes with the reversed `w:fontKey` GUID.
 *
 * Returns a new buffer; the input is not mutated. Throws if `fontKey` is not a
 * valid 128-bit GUID.
 *
 * @param data - Raw obfuscated font bytes from `word/fonts/*.odttf`.
 * @param fontKey - The `w:fontKey` GUID from the `w:embed*` element.
 * @public
 */
declare function deobfuscateFont(data: ArrayBuffer, fontKey: string): ArrayBuffer;

/**
 * Embedded font loading.
 *
 * Turns the obfuscated `word/fonts/*.odttf` binaries a DOCX carries into live
 * `@font-face` registrations so documents render in their authored fonts
 * instead of falling back to a metric-compatible substitute.
 *
 * The pure half ({@link getEmbeddedFontFaces}) resolves + de-obfuscates the
 * faces; {@link loadEmbeddedFonts} registers them with the browser (no-op
 * outside a DOM). Round-trip preservation is unaffected: the serializer copies
 * `word/fonts/*` and `fontTable.xml` from the original package untouched.
 */

/**
 * Names of the fonts a table embeds at least one face for. Used by the picker
 * to surface embedded fonts even when the canvas probe is unreliable for a
 * subsetted face.
 *
 * @public
 */
declare function getEmbeddedFontFamilies(fontTable: FontTable | undefined): Set<string>;
/** A single de-obfuscated embedded font face, ready for `loadFontFromBuffer`. */
interface EmbeddedFontFace {
    /** Word font name to register the face under. */
    family: string;
    /** CSS `font-weight` the face maps to (`embedBold*` → `'bold'`). */
    weight: 'normal' | 'bold';
    /** CSS `font-style` the face maps to (`embed*Italic` → `'italic'`). */
    style: 'normal' | 'italic';
    /** De-obfuscated OpenType/TrueType bytes. */
    data: ArrayBuffer;
    /** Whether the source face was subsetted (`w:subsetted`). */
    subsetted: boolean;
}
/**
 * Resolve and de-obfuscate every embedded font face declared in a font table.
 * Pure: does not touch the DOM. Faces whose relationship or binary is missing,
 * or whose key is unusable, are skipped.
 *
 * @param fontTable - Parsed `fontTable.xml` (`pkg.fontTable`).
 * @param rawFonts - Unzipped font binaries keyed by package path.
 * @param fontTableRelsXml - Raw `word/_rels/fontTable.xml.rels` XML.
 * @public
 */
declare function getEmbeddedFontFaces(fontTable: FontTable | undefined, rawFonts: ReadonlyMap<string, ArrayBuffer>, fontTableRelsXml: string | null | undefined): EmbeddedFontFace[];
/**
 * Register every embedded font face with the browser via `@font-face`. No-op
 * outside a DOM (headless/SSR). Resolves to the set of font family names that
 * were registered (deduped), so callers can surface them in the font picker.
 *
 * @public
 */
declare function loadEmbeddedFonts(fontTable: FontTable | undefined, rawFonts: ReadonlyMap<string, ArrayBuffer>, fontTableRelsXml: string | null | undefined): Promise<Set<string>>;

/**
 * Discover the fonts a document actually references and that the browser can
 * render, so they can be offered in the font picker under a "Document fonts"
 * group. A font qualifies when it is embedded in the file (already loaded via
 * {@link loadEmbeddedFonts}) or when the host system can render it
 * ({@link canRenderFont}). Fonts that would only fall back to a substitute are
 * left out, so the selector never lists a face it cannot actually show.
 */

interface RenderableFontOptions {
    /** Font families already loaded from the document (embedded faces). */
    embeddedFamilies?: ReadonlySet<string>;
    /** Names already present in the picker (built-in / configured) to skip. */
    exclude?: Iterable<string>;
    /** Override the system-font probe. Defaults to {@link canRenderFont}. */
    canRender?: (name: string) => boolean;
}
/**
 * Filter a list of referenced font names down to renderable {@link FontOption}s.
 * Pure and DOM-free when `canRender` is injected.
 *
 * @public
 */
declare function selectRenderableFonts(names: readonly string[], options?: RenderableFontOptions): FontOption[];
/**
 * Walk a parsed document for the fonts it references and return those the
 * browser can render (embedded or system-resolved) as picker options.
 *
 * @public
 */
declare function getRenderableDocumentFonts(doc: Document, options?: RenderableFontOptions): FontOption[];
/**
 * Drop fonts whose names already appear in `existingNames` (case-insensitive),
 * also deduping the input. Used by both adapters' pickers to render the
 * "Document fonts" group without repeating a font the built-in list covers.
 *
 * @public
 */
declare function excludeFontsByName(fonts: readonly FontOption[] | undefined, existingNames: Iterable<string>): FontOption[];

/**
 * Formatting to CSS Converter
 *
 * Converts OOXML formatting objects (TextFormatting, ParagraphFormatting)
 * to React CSSProperties for rendering.
 *
 * Handles ALL formatting properties:
 * - Font: family, size, weight, style
 * - Text: color, background, decoration (underline, strike, double-strike)
 * - Effects: superscript, subscript, small-caps, all-caps
 * - Spacing: letter-spacing
 * - Paragraph: alignment, line-height, margins, padding, borders, background
 */
/** Framework-agnostic CSS properties type (compatible with React.CSSProperties) */
type CSSProperties$1 = Record<string, any>;

/**
 * Convert TextFormatting to CSS properties for a run/span
 *
 * @param formatting - Text formatting from OOXML
 * @param theme - Theme for resolving colors and fonts
 * @returns React CSSProperties
 */
declare function textToStyle(formatting: TextFormatting | undefined | null, theme?: Theme | null): CSSProperties$1;
/**
 * Convert ParagraphFormatting to CSS properties
 *
 * @param formatting - Paragraph formatting from OOXML
 * @param theme - Theme for resolving colors
 * @returns React CSSProperties
 */
declare function paragraphToStyle(formatting: ParagraphFormatting | undefined | null, theme?: Theme | null): CSSProperties$1;
/**
 * Convert a BorderSpec to CSS border properties
 *
 * @param border - Border specification
 * @param side - 'Top' | 'Bottom' | 'Left' | 'Right' | '' for all
 * @param theme - Theme for color resolution
 * @returns Partial CSSProperties with border styles
 */
declare function borderToStyle(border: BorderSpec | undefined | null, side?: 'Top' | 'Bottom' | 'Left' | 'Right' | '', theme?: Theme | null): CSSProperties$1;
/**
 * Convert ShadingProperties to background color
 *
 * @param shading - Shading properties
 * @param theme - Theme for color resolution
 * @returns CSS color string or empty string
 */
declare function resolveShadingFill(shading: ShadingProperties | undefined | null, theme?: Theme | null): string;
/**
 * Merge multiple CSSProperties objects
 *
 * Later objects override earlier ones for conflicting properties.
 *
 * @param styles - Array of CSSProperties objects
 * @returns Merged CSSProperties
 */
declare function mergeStyles(...styles: (CSSProperties$1 | undefined | null)[]): CSSProperties$1;
/**
 * Get CSS for a table cell based on formatting
 *
 * @param formatting - Table cell formatting
 * @param theme - Theme for color resolution
 * @returns CSSProperties for the cell
 */
declare function tableCellToStyle(formatting: {
    verticalAlign?: 'top' | 'center' | 'bottom';
    textDirection?: string;
    shading?: ShadingProperties;
    borders?: {
        top?: BorderSpec;
        bottom?: BorderSpec;
        left?: BorderSpec;
        right?: BorderSpec;
    };
    margins?: {
        top?: {
            value: number;
            type: string;
        };
        bottom?: {
            value: number;
            type: string;
        };
        left?: {
            value: number;
            type: string;
        };
        right?: {
            value: number;
            type: string;
        };
    };
} | undefined | null, theme?: Theme | null): CSSProperties$1;
/**
 * Get CSS for page/section container
 *
 * @param sectionProps - Section properties
 * @returns CSSProperties for the page container
 */
declare function sectionToStyle(sectionProps: {
    pageWidth?: number;
    pageHeight?: number;
    marginTop?: number;
    marginBottom?: number;
    marginLeft?: number;
    marginRight?: number;
    background?: {
        color?: {
            rgb?: string;
            themeColor?: string;
        };
    };
} | undefined | null, theme?: Theme | null): CSSProperties$1;

/**
 * Keyboard Navigation Utilities
 *
 * Provides enhanced keyboard navigation for the editor:
 * - Ctrl+Left/Right: Move by word
 * - Home/End: Move to start/end of line
 * - Ctrl+Home/End: Move to start/end of document
 * - Ctrl+Shift+Left/Right: Select by word
 * - Shift+Home/End: Select to start/end of line
 */
/**
 * Navigation direction
 */
type NavigationDirection = 'left' | 'right' | 'up' | 'down';
/**
 * Navigation unit
 */
type NavigationUnit = 'character' | 'word' | 'line' | 'paragraph' | 'document';
/**
 * Keyboard navigation action
 */
interface NavigationAction {
    /** Direction to navigate */
    direction: NavigationDirection;
    /** Unit of movement */
    unit: NavigationUnit;
    /** Whether to extend selection */
    extend: boolean;
}
/**
 * Keyboard shortcut definition
 */
interface KeyboardShortcut {
    key: string;
    ctrlKey?: boolean;
    metaKey?: boolean;
    shiftKey?: boolean;
    altKey?: boolean;
}
/**
 * Check if a character is a word character (letter, digit, or underscore)
 */
declare function isWordCharacter(char: string): boolean;
/**
 * Check if a character is whitespace
 */
declare function isWhitespace(char: string): boolean;
/**
 * Check if a character is a punctuation character
 */
declare function isPunctuation(char: string): boolean;
/**
 * Find the start of the current or previous word
 */
declare function findWordStart(text: string, position: number): number;
/**
 * Find the end of the current or next word
 */
declare function findWordEnd(text: string, position: number): number;
/**
 * Find the next word start (for Ctrl+Right navigation)
 */
declare function findNextWordStart(text: string, position: number): number;
/**
 * Find the previous word start (for Ctrl+Left navigation)
 */
declare function findPreviousWordStart(text: string, position: number): number;
/**
 * Find the start of the current line in a text node
 * Uses visual line detection based on bounding rectangles
 */
declare function findVisualLineStart(container: Node, offset: number): {
    node: Node;
    offset: number;
} | null;
/**
 * Find the end of the current line in a text node
 * Uses visual line detection based on bounding rectangles
 */
declare function findVisualLineEnd(container: Node, offset: number): {
    node: Node;
    offset: number;
} | null;
/**
 * Get the current selection info
 */
declare function getSelectionInfo(): {
    node: Node;
    offset: number;
    anchorNode: Node | null;
    anchorOffset: number;
    focusNode: Node | null;
    focusOffset: number;
    isCollapsed: boolean;
    text: string;
} | null;
/**
 * Set the selection to a specific position
 */
declare function setSelectionPosition(node: Node, offset: number): void;
/**
 * Extend selection to a specific position
 */
declare function extendSelectionTo(node: Node, offset: number): void;
/**
 * Move selection by word in a text node
 */
declare function moveByWord(direction: 'left' | 'right', extend?: boolean): boolean;
/**
 * Move to start/end of line
 */
declare function moveToLineEdge(edge: 'start' | 'end', extend?: boolean): boolean;
/**
 * Parse a keyboard event into a navigation action
 */
declare function parseNavigationAction(event: KeyboardEvent): NavigationAction | null;
/**
 * Handle a keyboard navigation event
 * Returns true if the event was handled
 */
declare function handleNavigationKey(event: KeyboardEvent, options?: {
    onDocumentStart?: () => void;
    onDocumentEnd?: () => void;
}): boolean;
/**
 * Check if an event is a navigation key event
 */
declare function isNavigationKey(event: KeyboardEvent): boolean;
/**
 * Expand selection to word boundaries
 * Used for double-click word selection
 */
declare function expandSelectionToWord(): boolean;
/**
 * Get the word at the current cursor position
 */
declare function getWordAtCursor(): string | null;
/**
 * Check if a keyboard event matches a shortcut definition
 */
declare function matchesShortcut(event: KeyboardEvent, shortcut: KeyboardShortcut): boolean;
/**
 * Common navigation shortcuts
 */
declare const NAVIGATION_SHORTCUTS: {
    readonly wordLeft: KeyboardShortcut;
    readonly wordRight: KeyboardShortcut;
    readonly selectWordLeft: KeyboardShortcut;
    readonly selectWordRight: KeyboardShortcut;
    readonly lineStart: KeyboardShortcut;
    readonly lineEnd: KeyboardShortcut;
    readonly selectToLineStart: KeyboardShortcut;
    readonly selectToLineEnd: KeyboardShortcut;
    readonly documentStart: KeyboardShortcut;
    readonly documentEnd: KeyboardShortcut;
    readonly selectToDocumentStart: KeyboardShortcut;
    readonly selectToDocumentEnd: KeyboardShortcut;
};
/**
 * Get a human-readable description of a shortcut
 */
declare function describeShortcut(shortcut: KeyboardShortcut): string;
/**
 * Get all navigation shortcuts with descriptions
 */
declare function getNavigationShortcutDescriptions(): Array<{
    action: string;
    shortcut: string;
}>;

/**
 * Selection Highlight Utilities
 *
 * Provides visual highlighting for text selection across multiple runs.
 * Browsers handle ::selection pseudo-element differently, especially when
 * selection spans multiple elements with different backgrounds or styling.
 *
 * This module provides:
 * - Custom selection highlight rendering
 * - Programmatic selection range marking
 * - Visual feedback for selection across runs
 */
/** Framework-agnostic CSS properties type (compatible with React.CSSProperties) */
type CSSProperties = Record<string, any>;
/**
 * Highlight rectangle representing a selected region
 */
interface HighlightRect {
    /** Left position in pixels */
    left: number;
    /** Top position in pixels */
    top: number;
    /** Width in pixels */
    width: number;
    /** Height in pixels */
    height: number;
}
/**
 * Selection highlight configuration
 */
interface SelectionHighlightConfig {
    /** Background color for selection */
    backgroundColor: string;
    /** Optional border color for selection */
    borderColor?: string;
    /** Optional border radius */
    borderRadius?: number;
    /** Z-index for overlay */
    zIndex?: number;
    /** Opacity for highlight */
    opacity?: number;
    /** Mix blend mode */
    mixBlendMode?: CSSProperties['mixBlendMode'];
}
/**
 * Selection range in document coordinates
 */
interface SelectionRange {
    /** Start position */
    start: {
        paragraphIndex: number;
        contentIndex: number;
        offset: number;
    };
    /** End position */
    end: {
        paragraphIndex: number;
        contentIndex: number;
        offset: number;
    };
}
/**
 * Default selection highlight style (matches Word/Google Docs)
 */
declare const DEFAULT_SELECTION_STYLE: SelectionHighlightConfig;
/**
 * High contrast selection style
 */
declare const HIGH_CONTRAST_SELECTION_STYLE: SelectionHighlightConfig;
/**
 * Selection highlight CSS custom properties
 */
declare const SELECTION_CSS_VARS: {
    readonly backgroundColor: "--docx-selection-bg";
    readonly borderColor: "--docx-selection-border";
    readonly textColor: "--docx-selection-text";
};
/**
 * Get all selection rectangles from the current DOM selection
 *
 * Uses getClientRects() to get accurate rectangles even when
 * selection spans multiple inline elements.
 */
declare function getSelectionRects(containerElement?: HTMLElement | null): HighlightRect[];
/**
 * Merge adjacent or overlapping rectangles
 *
 * This reduces the number of highlight elements needed and creates
 * a cleaner visual appearance.
 */
declare function mergeAdjacentRects(rects: HighlightRect[], tolerance?: number): HighlightRect[];
/**
 * Get selection rectangles with merging applied
 */
declare function getMergedSelectionRects(containerElement?: HTMLElement | null): HighlightRect[];
/**
 * Generate CSS styles for a highlight rectangle
 */
declare function getHighlightRectStyle(rect: HighlightRect, config?: SelectionHighlightConfig): CSSProperties;
/**
 * Generate inline CSS for selection pseudo-elements
 *
 * This is used to inject consistent selection styling
 * across all editable elements.
 */
declare function generateSelectionCSS(selector: string, config?: SelectionHighlightConfig): string;
/**
 * Check if there is an active text selection (not collapsed)
 */
declare function hasActiveSelection(): boolean;
/**
 * Get the selected text
 */
declare function getSelectedText(): string;
/**
 * Check if selection is within a specific element
 */
declare function isSelectionWithin(element: HTMLElement): boolean;
/**
 * Get the bounding rect of the current selection
 */
declare function getSelectionBoundingRect(): DOMRect | null;
/**
 * Create a selection highlight for a specific text range
 *
 * This is useful for find/replace highlighting, AI action previews, etc.
 */
declare function highlightTextRange(_containerElement: HTMLElement, startNode: Node, startOffset: number, endNode: Node, endOffset: number): Range | null;
/**
 * Select a text range programmatically
 */
declare function selectRange(range: Range): void;
/**
 * Clear the current selection
 */
declare function clearSelection(): void;
/**
 * Check if selection is backwards (focus before anchor)
 */
declare function isSelectionBackwards(): boolean;
/**
 * Normalize selection to always be forward (start before end)
 */
declare function normalizeSelectionDirection(): void;
/**
 * Inject selection highlight CSS into document
 */
declare function injectSelectionStyles(config?: SelectionHighlightConfig): void;
/**
 * Remove injected selection styles
 */
declare function removeSelectionStyles(): void;
/**
 * Check if selection styles are injected
 */
declare function areSelectionStylesInjected(): boolean;
/**
 * Create a selection change handler that updates highlight rects
 */
declare function createSelectionChangeHandler(containerElement: HTMLElement | null, onRectsChange: (rects: HighlightRect[]) => void, merge?: boolean): () => void;

/**
 * Option shapes for `scrollToParaId(paraId, { highlight })`.
 *
 * Kept in a DOM-free module (no imports, no browser globals) so non-browser
 * consumers can type-import them without pulling `paragraphFlash`'s DOM code
 * into their type-check surface.
 */
/**
 * Customization for the transient paragraph flash applied by
 * `scrollToParaId(paraId, { highlight })`.
 *
 * @public
 */
interface ParagraphHighlightOptions {
    /** CSS color used for the transient paragraph flash. Defaults to yellow. */
    color?: string;
    /** How long the flash remains visible before it is removed. Defaults to 1200ms. */
    durationMs?: number;
}
/**
 * Optional reveal behavior for `scrollToParaId`.
 *
 * @public
 */
interface ScrollToParaIdOptions {
    /** Flash rendered paragraph fragments after scrolling to the paragraph. */
    highlight?: ParagraphHighlightOptions;
}

/**
 * DOM helpers for transient paragraph flashes on the painted layout surface.
 */

/** Default color used by paragraph flashes. */
declare const DEFAULT_PARAGRAPH_FLASH_COLOR = "rgba(255, 235, 59, 0.55)";
/** Default duration for paragraph flashes. */
declare const DEFAULT_PARAGRAPH_FLASH_DURATION_MS = 1200;
/** CSS class applied to paragraph fragments during a transient flash. */
declare const PARAGRAPH_FLASH_CLASS_NAME = "docx-paragraph-flash";
/**
 * Find all painted paragraph fragments with a stable `data-para-id`.
 *
 * @public
 */
declare function findParagraphFragmentsByParaId(root: ParentNode, paraId: string): HTMLElement[];
/**
 * Apply a transient flash to a collection of paragraph elements.
 *
 * @returns the number of elements flashed
 * @public
 */
declare function flashParagraphElements(elements: Iterable<HTMLElement>, options?: ParagraphHighlightOptions): number;
/**
 * Find paragraph fragments by `paraId` and flash them.
 *
 * @returns whether at least one rendered fragment was found
 * @public
 */
declare function flashParagraphFragmentsByParaId(root: ParentNode, paraId: string, options?: ParagraphHighlightOptions): boolean;

/**
 * Shared table split algorithm — model-agnostic core logic.
 *
 * This module contains the pure layout computation for splitting a table cell:
 * column width redistribution, neighbor span adjustment, and new-cell placement.
 * Both the ProseMirror path (tableSplit.ts) and the Document-model path
 * (TableToolbar.tsx) delegate to these functions.
 */
/** A cell's position and span within the logical grid. */
interface CellAnchor<T> {
    /** Opaque payload — the caller's cell type (PMNode, TableCell, etc.) */
    data: T;
    row: number;
    col: number;
    rowspan: number;
    colspan: number;
}
/** Parameters describing the split target. */
interface SplitTarget {
    row: number;
    col: number;
    rowspan: number;
    colspan: number;
}
/** Result of `computeSplitLayout`. */
interface SplitLayoutResult<T> {
    /** All anchors after the split (neighbors adjusted + new split cells). */
    anchors: CellAnchor<T>[];
    deltaRows: number;
    deltaCols: number;
    newRowCount: number;
}
declare function sumColumnWidths(widths: number[], start: number, span: number): number;
/**
 * Redistribute column widths when splitting a cell's column span.
 *
 * @param existing   Current column widths array for the whole table.
 * @param startCol   First column of the cell being split.
 * @param currentSpan Current column span of the cell.
 * @param targetSpan  Desired column span after split.
 * @returns New column widths array with the split applied.
 */
declare function redistributeColumnWidths(existing: number[], startCol: number, currentSpan: number, targetSpan: number): number[];
/**
 * Compute the new anchor layout after splitting a target cell.
 *
 * This is the core algorithm shared between ProseMirror and Document-model
 * paths. It adjusts neighbor spans, shifts positions for inserted rows/cols,
 * and creates placeholder anchors for the new split cells.
 *
 * @param anchors     All cell anchors in the current table.
 * @param target      The anchor being split.
 * @param rows        Number of rows the target should become.
 * @param cols        Number of columns the target should become.
 * @param totalRows   Current total row count.
 * @param createSplitCellData  Factory to create `data` for each new split cell.
 *   Called with `(isOriginal, rowOffset, colOffset)` — `isOriginal` is true for
 *   the top-left cell that retains the original content.
 */
declare function computeSplitLayout<T>(anchors: CellAnchor<T>[], target: CellAnchor<T>, rows: number, cols: number, totalRows: number, createSplitCellData: (isOriginal: boolean, rowOffset: number, colOffset: number) => T): SplitLayoutResult<T>;
/**
 * Build lookup maps from an anchor list — by start position and by covered slot.
 */
declare function buildAnchorMaps<T>(anchors: CellAnchor<T>[]): {
    byStart: Map<string, CellAnchor<T>>;
    byCoveredSlot: Map<string, CellAnchor<T>>;
};
/**
 * Compute the initial dialog values for a split-cell dialog.
 */
declare function computeSplitDialogDefaults(rowspan: number, colspan: number): {
    minRows: number;
    minCols: number;
    initialRows: number;
    initialCols: number;
};

/**
 * Shared file-input → docx-buffer reader.
 *
 * React (DocxEditor.tsx `handleDocxFileChange`) and Vue
 * (DocxEditor.vue `handleDocxFileChange`) had byte-equivalent
 * `await file.arrayBuffer()` + `name.replace(/\.docx$/i, '')` boilerplate
 * around their hidden file inputs. This helper folds the common steps
 * into one place so the two adapters can never drift on filename
 * normalization or on the "reset input.value so re-picking the same
 * file fires `change` again" detail.
 *
 * Returns `null` when the user cancelled the picker (no file chosen).
 */
interface ReadDocxFileResult {
    /** ArrayBuffer ready to feed into `loadBuffer` / `parseDocx`. */
    buffer: ArrayBuffer;
    /** File name with the trailing `.docx` extension stripped. */
    name: string;
}
/**
 * Read the first selected file out of an `<input type="file">` change
 * event, return its ArrayBuffer + a stem-form name. Always resets
 * `input.value` so re-selecting the same file in the same picker fires
 * the next `change` event.
 */
declare function readDocxFileFromInput(event: Event): Promise<ReadDocxFileResult | null>;

/**
 * Color-mode resolution shared by every adapter so the light/dark/auto logic
 * (and its SSR + prefers-color-scheme handling) can't drift between React, Vue,
 * or any future host. The adapters keep only their framework binding
 * (useState/useEffect vs ref/watchEffect); the rules live here.
 *
 * @packageDocumentation
 */
/** UI color theme for the editor chrome and canvas. */
type ColorMode = 'light' | 'dark' | 'system';
/**
 * Current OS dark-mode preference. Returns false when `matchMedia` is
 * unavailable (e.g. server-side render), so it is safe to call as a state seed.
 *
 * @public
 */
declare function prefersColorSchemeDark(): boolean;
/**
 * Resolve the effective dark flag from a {@link ColorMode} and the current OS
 * preference. `'system'` follows the OS; `'dark'`/`'light'` are explicit.
 *
 * @public
 */
declare function resolveIsDark(colorMode: ColorMode, systemDark: boolean): boolean;
/**
 * Subscribe to OS dark-mode changes. Invokes `onChange` immediately with the
 * current value (so a stale seed is corrected on entry), then on every change.
 * Returns an unsubscribe function. A no-op under SSR.
 *
 * @public
 */
declare function subscribeSystemDark(onChange: (dark: boolean) => void): () => void;

/**
 * Allowlist URL schemes on hrefs that originate from untrusted input
 * (DOCX relationship targets, pasted HTML). Fragments and relative paths
 * pass through; anything with a scheme outside the allowlist is dropped.
 */
declare function sanitizeHref(href: string | null | undefined): string | undefined;

export { CLIPBOARD_TYPES, type CellAnchor, type ClipboardContent, type ClipboardOptions, type ColorMode, DEFAULT_PARAGRAPH_FLASH_COLOR, DEFAULT_PARAGRAPH_FLASH_DURATION_MS, DEFAULT_SELECTION_STYLE, type EmbeddedFontFace, HIGH_CONTRAST_SELECTION_STYLE, type HighlightRect, INTERNAL_CLIPBOARD_TYPE, type KeyboardShortcut, NAVIGATION_SHORTCUTS, type NavigationAction, type NavigationDirection, type NavigationUnit, PARAGRAPH_FLASH_CLASS_NAME, type ParagraphHighlightOptions, type ParsedClipboardContent, type ReadDocxFileResult, type RenderableFontOptions, SELECTION_CSS_VARS, type ScrollToParaIdOptions, type SelectionHighlightConfig, type SelectionRange, type SplitLayoutResult, type SplitTarget, areSelectionStylesInjected, borderToStyle, buildAnchorMaps, cleanWordHtml, clearSelection, computeSplitDialogDefaults, computeSplitLayout, copyParagraphs, copyRuns, createClipboardHandlers, createSelectionChangeHandler, deobfuscateFont, describeShortcut, excludeFontsByName, expandSelectionToWord, extendSelectionTo, findNextWordStart, findParagraphFragmentsByParaId, findPreviousWordStart, findVisualLineEnd, findVisualLineStart, findWordEnd, findWordStart, flashParagraphElements, flashParagraphFragmentsByParaId, generateSelectionCSS, getClipboardImageFiles, getEmbeddedFontFaces, getEmbeddedFontFamilies, getHighlightRectStyle, getMergedSelectionRects, getNavigationShortcutDescriptions, getRenderableDocumentFonts, getSelectedText, getSelectionBoundingRect, getSelectionInfo, getSelectionRects, getWordAtCursor, handleNavigationKey, handlePasteEvent, hasActiveSelection, highlightTextRange, htmlToRuns, injectSelectionStyles, isEditorHtml, isNavigationKey, isPunctuation, isSelectionBackwards, isSelectionWithin, isValidFontKey, isWhitespace, isWordCharacter, isWordHtml, loadEmbeddedFonts, matchesShortcut, mergeAdjacentRects, mergeStyles, moveByWord, moveToLineEdge, normalizeSelectionDirection, paragraphToStyle, paragraphsToClipboardContent, parseClipboardHtml, parseNavigationAction, prefersColorSchemeDark, readDocxFileFromInput, readFromClipboard, redistributeColumnWidths, removeSelectionStyles, resolveIsDark, resolveShadingFill, runsToClipboardContent, sanitizeHref, sectionToStyle, selectRange, selectRenderableFonts, setSelectionPosition, subscribeSystemDark, sumColumnWidths, tableCellToStyle, textToStyle, writeToClipboard };
