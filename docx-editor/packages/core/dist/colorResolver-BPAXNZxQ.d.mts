import { Document } from './types/document.mjs';
import { T as ThemeColorSlot, C as ColorValue } from './colors-C3vA7HUU.mjs';
import { T as Theme, g as ThemeColorScheme } from './styles-Diw0MASy.mjs';

/**
 * Template Processing Utility
 *
 * Uses docxtemplater to substitute template variables in DOCX documents:
 * - Processes {variable_name} patterns (docxtemplater default syntax)
 * - Preserves all formatting (fonts, styles, colors, tables)
 * - Error handling with useful messages
 */
/**
 * Options for template processing
 */
interface ProcessTemplateOptions {
    /** How to handle undefined variables */
    nullGetter?: 'keep' | 'empty' | 'error';
    /** Custom parser for variable names */
    parser?: (tag: string) => {
        get: (scope: Record<string, unknown>) => unknown;
    };
    /** Line breaks: keep raw \n or convert to w:br */
    linebreaks?: boolean;
    /** Delimiter settings */
    delimiters?: {
        start?: string;
        end?: string;
    };
}
/**
 * Result of template processing
 */
interface ProcessTemplateResult {
    /** The processed document buffer */
    buffer: ArrayBuffer;
    /** Variables that were found and replaced */
    replacedVariables: string[];
    /** Variables that were not replaced (no value provided) */
    unreplacedVariables: string[];
    /** Any warnings during processing */
    warnings: string[];
}
/**
 * Error details from template processing
 */
interface TemplateError {
    /** Error message */
    message: string;
    /** Variable name that caused the error (if applicable) */
    variable?: string;
    /** Error type */
    type: 'parse' | 'render' | 'undefined' | 'unknown';
    /** Original error */
    originalError?: Error;
}
/**
 * Process a DOCX template with variable substitution
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param variables - Map of variable names to values
 * @param options - Processing options
 * @returns Processed DOCX as ArrayBuffer
 */
declare function processTemplate(buffer: ArrayBuffer, variables: Record<string, string>, options?: ProcessTemplateOptions): ArrayBuffer;
/**
 * Process template with detailed result
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param variables - Map of variable names to values
 * @param options - Processing options
 * @returns Detailed processing result
 */
declare function processTemplateDetailed(buffer: ArrayBuffer, variables: Record<string, string>, options?: ProcessTemplateOptions): ProcessTemplateResult;
/**
 * Process template and return as Blob
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param variables - Map of variable names to values
 * @param options - Processing options
 * @returns Processed DOCX as Blob
 */
declare function processTemplateAsBlob(buffer: ArrayBuffer, variables: Record<string, string>, options?: ProcessTemplateOptions): Blob;
/**
 * Process template and trigger download
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param variables - Map of variable names to values
 * @param filename - Output filename (without extension)
 * @param options - Processing options
 */
declare function processTemplateAndDownload(buffer: ArrayBuffer, variables: Record<string, string>, filename?: string, options?: ProcessTemplateOptions): void;
/**
 * Get all template tags in a document without processing
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @returns List of tag names found
 */
declare function getTemplateTags(buffer: ArrayBuffer): string[];
/**
 * Validate that a document is a valid docxtemplater template
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @returns Validation result
 */
declare function validateTemplate(buffer: ArrayBuffer): {
    valid: boolean;
    errors: TemplateError[];
    tags: string[];
};
/**
 * Check if all required variables have values
 *
 * @param tags - List of template tags
 * @param variables - Provided variable values
 * @returns Missing variable names
 */
declare function getMissingVariables(tags: string[], variables: Record<string, string>): string[];
/**
 * Preview what the document will look like after processing
 * Returns the document text with variables replaced (for preview purposes)
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param variables - Map of variable names to values
 * @returns Preview text
 */
declare function previewTemplate(buffer: ArrayBuffer, variables: Record<string, string>): string;
/**
 * Process template with conditional sections
 * Supports #if, #unless, #each loops
 *
 * @param buffer - The DOCX file as ArrayBuffer
 * @param data - Full data object (can include arrays, nested objects)
 * @param options - Processing options
 * @returns Processed DOCX as ArrayBuffer
 */
declare function processTemplateAdvanced(buffer: ArrayBuffer, data: Record<string, unknown>, options?: ProcessTemplateOptions): ArrayBuffer;
/**
 * Create a template processor with preset options
 */
declare function createTemplateProcessor(defaultOptions?: ProcessTemplateOptions): (buffer: ArrayBuffer, variables: Record<string, string>) => ArrayBuffer;

/**
 * Create Document Utility
 *
 * Provides functions to create new documents programmatically.
 */

/**
 * Options for creating an empty document
 */
interface CreateEmptyDocumentOptions {
    /** Page width in twips (default: 12240 = 8.5 inches) */
    pageWidth?: number;
    /** Page height in twips (default: 15840 = 11 inches) */
    pageHeight?: number;
    /** Page orientation (default: 'portrait') */
    orientation?: 'portrait' | 'landscape';
    /** Top margin in twips (default: 1440 = 1 inch) */
    marginTop?: number;
    /** Bottom margin in twips (default: 1440 = 1 inch) */
    marginBottom?: number;
    /** Left margin in twips (default: 1440 = 1 inch) */
    marginLeft?: number;
    /** Right margin in twips (default: 1440 = 1 inch) */
    marginRight?: number;
    /** Initial text content (default: empty string) */
    initialText?: string;
}
/**
 * Create an empty document with a single paragraph
 *
 * @param options - Optional configuration for the document
 * @returns A new empty Document object
 *
 * @example
 * ```ts
 * // Create a blank document
 * const doc = createEmptyDocument();
 *
 * // Create with custom margins
 * const doc = createEmptyDocument({
 *   marginTop: 720,  // 0.5 inch
 *   marginBottom: 720,
 * });
 *
 * // Create with initial text
 * const doc = createEmptyDocument({
 *   initialText: 'Hello, World!'
 * });
 * ```
 */
declare function createEmptyDocument(options?: CreateEmptyDocumentOptions): Document;
/**
 * Create a document with a single paragraph containing the given text
 *
 * @param text - The text content for the document
 * @param options - Optional configuration for the document
 * @returns A new Document object with the specified text
 */
declare function createDocumentWithText(text: string, options?: Omit<CreateEmptyDocumentOptions, 'initialText'>): Document;

/**
 * Color Resolver - Convert OOXML colors to CSS
 *
 * Handles:
 * - Theme color references (accent1, dk1, etc.)
 * - RGB hex values
 * - "auto" colors (context-dependent)
 * - Tint/shade modifications
 *
 * OOXML Color References:
 * - w:color/@w:val - RGB hex or "auto"
 * - w:color/@w:themeColor - Theme color slot
 * - w:color/@w:themeTint - Tint modifier (0-255, hex)
 * - w:color/@w:themeShade - Shade modifier (0-255, hex)
 *
 * Tint/Shade Calculations:
 * - Tint makes color lighter (blend with white)
 * - Shade makes color darker (blend with black)
 * - Value is in hex (00-FF), converted to 0-1 for calculation
 */

/**
 * Resolve a ColorValue to a CSS color string
 *
 * @param color - ColorValue object with rgb, themeColor, tint/shade, or auto
 * @param theme - Theme for resolving theme colors
 * @param defaultColor - Default color if auto or undefined (default: black)
 * @returns CSS color string (e.g., "#FF0000" or "inherit")
 */
declare function resolveColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, defaultColor?: string): string;
/**
 * Resolve any ColorValue (text, fill/shading, border, underline) to a 6-char
 * uppercase hex string — or `undefined` if transparent/unset/unresolvable.
 *
 * Shared display-side resolver. Prefer this over reading `.rgb` directly so
 * that `themeColor` + `themeTint`/`themeShade` are honored consistently across
 * all render paths (PM attrs, layout-bridge, clipboard HTML, toolbar swatches).
 *
 * When a themed color is present but `theme` is null/undefined, falls back to
 * `color.rgb` if Word wrote one for compat; otherwise returns `undefined`.
 *
 * @returns 6-char uppercase hex without `#`, or `undefined`.
 */
declare function resolveColorToHex(color: ColorValue | undefined | null, theme: Theme | null | undefined): string | undefined;
/**
 * Resolve a highlight color name to CSS
 *
 * @param highlight - Highlight color name (e.g., "yellow", "cyan")
 * @returns CSS color string or empty string for "none"
 */
declare function resolveHighlightColor(highlight: string | undefined): string;
/**
 * Resolve a shading fill or pattern color to CSS
 *
 * @param color - ColorValue for fill
 * @param theme - Theme for resolving theme colors
 * @returns CSS color string
 */
declare function resolveShadingColor(color: ColorValue | undefined | null, theme: Theme | null | undefined): string;
/**
 * Check if a color is effectively black
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to black or very dark
 */
declare function isBlack(color: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;
/**
 * Check if a color is effectively white
 *
 * @param color - ColorValue object
 * @param theme - Theme for resolving theme colors
 * @returns True if color resolves to white or very light
 */
declare function isWhite(color: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;
/**
 * Get contrasting text color for a background
 *
 * @param backgroundColor - Background ColorValue
 * @param theme - Theme for resolving theme colors
 * @returns Black or white hex color for best contrast
 */
declare function getContrastingColor(backgroundColor: ColorValue | undefined | null, theme: Theme | null | undefined): string;
/**
 * Parse a color string (various formats) to ColorValue
 *
 * @param colorString - Color string like "FF0000", "auto", or theme color name
 * @returns ColorValue object
 */
declare function parseColorString(colorString: string | undefined): ColorValue | undefined;
/**
 * Create a ColorValue from theme color reference
 *
 * @param themeColor - Theme color slot name
 * @param tint - Optional tint modifier
 * @param shade - Optional shade modifier
 * @returns ColorValue object
 */
declare function createThemeColor(themeColor: ThemeColorSlot, tint?: number, shade?: number): ColorValue;
/**
 * Create a ColorValue from RGB hex
 *
 * @param hex - 6-character hex color (no #)
 * @returns ColorValue object
 */
declare function createRgbColor(hex: string): ColorValue;
/**
 * Darken a color by a percentage
 *
 * @param color - ColorValue to darken
 * @param theme - Theme for resolving
 * @param percent - Percentage to darken (0-100)
 * @returns CSS color string
 */
declare function darkenColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, percent: number): string;
/**
 * Lighten a color by a percentage
 *
 * @param color - ColorValue to lighten
 * @param theme - Theme for resolving
 * @param percent - Percentage to lighten (0-100)
 * @returns CSS color string
 */
declare function lightenColor(color: ColorValue | undefined | null, theme: Theme | null | undefined, percent: number): string;
/**
 * Blend two colors together
 *
 * @param color1 - First color
 * @param color2 - Second color
 * @param ratio - Blend ratio (0 = all color1, 1 = all color2)
 * @param theme - Theme for resolving
 * @returns CSS color string
 */
declare function blendColors(color1: ColorValue | undefined | null, color2: ColorValue | undefined | null, ratio: number, theme: Theme | null | undefined): string;
/**
 * Ensure a hex color string has a '#' prefix.
 */
declare function ensureHexPrefix(hex: string): string;
/**
 * Resolve a highlight color value to a CSS-ready string.
 * Tries OOXML named highlight first, then ensures hex prefix.
 */
declare function resolveHighlightToCss(value: string): string;
/**
 * Theme color matrix cell
 */
interface ThemeMatrixCell {
    /** Resolved hex color (6 chars, no #) */
    hex: string;
    /** Theme color slot */
    themeSlot: ThemeColorSlot;
    /** Tint hex modifier if applicable (e.g., "CC") */
    tint?: string;
    /** Shade hex modifier if applicable (e.g., "BF") */
    shade?: string;
    /** Human-readable label (e.g., "Accent 1, Lighter 60%") */
    label: string;
}
/**
 * Compute a single tinted or shaded hex color from a base color.
 *
 * @param baseHex - 6-character hex color (no #)
 * @param type - 'tint' to lighten, 'shade' to darken
 * @param fraction - Amount (0-1). For tint: 0=no change, 1=white. For shade: 0=black, 1=no change.
 * @returns 6-character hex color (no #)
 */
declare function getThemeTintShadeHex(baseHex: string, type: 'tint' | 'shade', fraction: number): string;
/**
 * Generate the 10×6 theme color matrix for an advanced color picker.
 *
 * Columns: lt1, dk1, lt2, dk2, accent1-6 (matches Word's order)
 * Rows: base, 80% tint, 60% tint, 40% tint, 25% shade, 50% shade
 *
 * @param colorScheme - Theme color scheme (falls back to Office 2016 defaults)
 * @returns 6 rows × 10 columns of ThemeMatrixCell
 */
declare function generateThemeTintShadeMatrix(colorScheme?: ThemeColorScheme | null): ThemeMatrixCell[][];
/**
 * Check if two colors are equal
 *
 * @param color1 - First color
 * @param color2 - Second color
 * @param theme - Theme for resolving
 * @returns True if colors resolve to the same value
 */
declare function colorsEqual(color1: ColorValue | undefined | null, color2: ColorValue | undefined | null, theme: Theme | null | undefined): boolean;

export { ensureHexPrefix as A, generateThemeTintShadeMatrix as B, type CreateEmptyDocumentOptions as C, getThemeTintShadeHex as D, resolveHighlightToCss as E, processTemplateAndDownload as F, resolveColorToHex as G, type ProcessTemplateOptions as P, type TemplateError as T, type ProcessTemplateResult as a, blendColors as b, colorsEqual as c, createDocumentWithText as d, createEmptyDocument as e, createRgbColor as f, createTemplateProcessor as g, createThemeColor as h, darkenColor as i, getContrastingColor as j, getMissingVariables as k, getTemplateTags as l, isBlack as m, isWhite as n, lightenColor as o, parseColorString as p, previewTemplate as q, processTemplate as r, processTemplateAdvanced as s, processTemplateAsBlob as t, processTemplateDetailed as u, resolveColor as v, resolveHighlightColor as w, resolveShadingColor as x, validateTemplate as y, type ThemeMatrixCell as z };
