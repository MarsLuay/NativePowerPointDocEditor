/**
 * Text Measurement Module
 *
 * Provides text measurement utilities for the layout engine.
 * Uses Canvas API for accurate, cached measurements.
 * @packageDocumentation
 * @public
 */
import { TextRun, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types.js';
import { b as FloatingImageZone } from '../../measureBlocksPipeline-iYNzAH-o.js';
export { F as FloatPageGeometry, a as FloatingExclusionRect, c as FloatingLineSegmentZone, M as MeasureBlockFn, g as getFloatingMargins, m as measureBlocksWithFloats, r as rectsToFloatingZones } from '../../measureBlocksPipeline-iYNzAH-o.js';
import '../../content-B8ScSBzC.js';
import '../../formatting-JhqWT_XM.js';
import '../../colors-C3vA7HUU.js';
import '../../docx/wrapTypes.js';
import '../../lists-Bn29SzeS.js';
import '../../watermark-D90356ZM.js';
import '../../anchoredObjectPosition-CS-8BfhO.js';

/**
 * Measurement container for text layout
 *
 * Uses HTML5 Canvas API to measure text runs and calculate typography metrics.
 * Canvas-based measurement is more accurate and performant than DOM-based approaches.
 *
 * Typography conventions (matching Word behavior):
 * - ascent ≈ fontSize * 0.8 (baseline to top)
 * - descent ≈ fontSize * 0.2 (baseline to bottom)
 * - lineHeight from font metrics (fontBoundingBoxAscent + fontBoundingBoxDescent),
 *   falling back to fontSize * 1.0 (OOXML spec default single spacing)
 */
/**
 * Font styling properties for measurement
 */
interface FontStyle {
    fontFamily?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    letterSpacing?: number;
}
/**
 * Typography metrics for a font
 */
interface FontMetrics {
    fontSize: number;
    ascent: number;
    descent: number;
    lineHeight: number;
    fontFamily: string;
    /** OS/2 single-line ratio for OOXML line spacing calculation */
    singleLineRatio: number;
}
/**
 * Result of measuring a text string
 */
interface TextMeasurement {
    width: number;
    height: number;
    ascent: number;
    descent: number;
}
/**
 * Result of measuring a run of text
 */
interface RunMeasurement {
    width: number;
    charWidths: number[];
    metrics: FontMetrics;
}
/**
 * Get or create a canvas 2D context for text measurement
 */
declare function getCanvasContext(): CanvasRenderingContext2D;
/**
 * Reset the canvas context (useful for testing)
 */
declare function resetCanvasContext(): void;
/**
 * Build a CSS font string from styling properties
 *
 * Font sizes are in points and need to be converted to pixels for canvas.
 * 1pt = 96/72 px ≈ 1.333px at standard web DPI.
 *
 * Uses the font resolver to get category-appropriate fallback stacks
 * (serif fonts get serif fallbacks, sans-serif get sans-serif, etc.)
 * matching the same stacks used in rendering for consistent measurements.
 *
 * @example
 * buildFontString({ fontFamily: "Arial", fontSize: 12, bold: true })
 * // Returns: "bold 16px Arial, Arimo, Helvetica, sans-serif" (12pt = 16px)
 */
declare function buildFontString(style: FontStyle): string;
/**
 * Get typography metrics for a given font size and family
 *
 * Uses Canvas TextMetrics API when available for precise metrics,
 * falls back to ratio-based approximations.
 */
declare function getFontMetrics(style: FontStyle): FontMetrics;
/**
 * Measure the width of a text string with specific styling
 *
 * @param text - The text to measure
 * @param style - Font styling properties
 * @returns Width in pixels
 */
declare function measureTextWidth(text: string, style: FontStyle): number;
/**
 * Measure text and return full metrics
 */
declare function measureText(text: string, style: FontStyle): TextMeasurement;
/**
 * Measure a run of text and return per-character widths for click positioning
 *
 * @param text - The text to measure
 * @param style - Font styling properties
 * @returns Run measurement with width and per-character widths
 */
declare function measureRun(text: string, style: FontStyle): RunMeasurement;
/**
 * Find the character offset at a given X position within a text run
 *
 * @param x - X position relative to run start
 * @param charWidths - Per-character widths from measureRun
 * @returns Character offset (0-based index)
 */
declare function findCharacterAtX(x: number, charWidths: number[]): number;
/**
 * Get the X position of a character offset within a text run
 *
 * @param offset - Character offset (0-based index)
 * @param charWidths - Per-character widths from measureRun
 * @returns X position in pixels
 */
declare function getXForCharacter(offset: number, charWidths: number[]): number;
/**
 * Convert twips to pixels
 */
declare function twipsToPx(twips: number): number;
/**
 * Convert pixels to twips
 */
declare function pxToTwips(px: number): number;
/**
 * Convert points to pixels
 */
declare function ptToPx(pt: number): number;
/**
 * Convert pixels to points
 */
declare function pxToPt(px: number): number;
/**
 * Convert OOXML half-points to pixels
 * OOXML font sizes are in half-points (24 = 12pt)
 */
declare function halfPtToPx(halfPt: number): number;
/**
 * Convert pixels to OOXML half-points
 */
declare function pxToHalfPt(px: number): number;

/**
 * Paragraph measurement module
 *
 * Measures paragraph blocks and computes line breaking.
 * Converts runs into measured lines with typography metrics.
 */

/**
 * Options for paragraph measurement
 */
interface MeasureParagraphOptions {
    /** Floating image exclusion zones that affect line widths */
    floatingZones?: FloatingImageZone[];
    /** Y offset of this paragraph relative to the exclusion zones (default: 0) */
    paragraphYOffset?: number;
}
/**
 * When a float's wrap margins consume the entire content width (or more),
 * there is no horizontal strip beside it for body text. Word renders the
 * following lines at full content width instead of squeezing them into a
 * 1-pixel column. Unchecked margins from near-full-width tables/images can
 * exceed contentWidth and collapse every line to ~1 glyph (the "single
 * character per line after a wide floating table" bug).
 *
 * Returned margins are zeroed when:
 * - either side alone is >= contentWidth (no strip on that side at all), or
 * - their sum is >= contentWidth (no strip exists between the two sides).
 */
declare function clampFloatingWrapMargins(leftMargin: number, rightMargin: number, contentWidth: number): {
    leftMargin: number;
    rightMargin: number;
};
declare function measureParagraph(block: ParagraphBlock, maxWidth: number, options?: MeasureParagraphOptions): ParagraphMeasure;
/**
 * Measure multiple paragraph blocks
 *
 * @param blocks - Array of paragraph blocks to measure
 * @param maxWidth - Maximum available width
 * @returns Array of ParagraphMeasure results
 */
declare function measureParagraphs(blocks: ParagraphBlock[], maxWidth: number): ParagraphMeasure[];
/**
 * Get per-character widths for a text run (for click positioning)
 *
 * @param run - The text run to measure
 * @returns Array of character widths
 */
declare function getRunCharWidths(run: TextRun): number[];

/**
 * Measurement Cache
 *
 * LRU cache for text width measurements and paragraph layout results.
 * Improves performance by avoiding repeated measurements of identical content.
 */

/**
 * Get cached text width or return undefined
 */
declare function getCachedTextWidth(text: string, font: string, letterSpacing?: number): number | undefined;
/**
 * Store text width in cache
 */
declare function setCachedTextWidth(text: string, font: string, letterSpacing: number, width: number): void;
/**
 * Clear the text width cache
 */
declare function clearTextWidthCache(): void;
/**
 * Set the maximum size of the text width cache
 */
declare function setTextCacheSize(size: number): void;
/**
 * Get current text width cache size
 */
declare function getTextCacheSize(): number;
/**
 * Cached font metrics entry
 */
interface FontMetricsEntry {
    ascent: number;
    descent: number;
    lineHeight: number;
}
/**
 * Get cached font metrics or return undefined
 */
declare function getCachedFontMetrics(fontFamily: string, fontSize: number, bold?: boolean, italic?: boolean): FontMetricsEntry | undefined;
/**
 * Store font metrics in cache
 */
declare function setCachedFontMetrics(fontFamily: string, fontSize: number, bold: boolean, italic: boolean, metrics: FontMetricsEntry): void;
/**
 * Clear the font metrics cache
 */
declare function clearFontMetricsCache(): void;
/**
 * Set the maximum size of the font metrics cache
 */
declare function setFontCacheSize(size: number): void;
/**
 * Get current font metrics cache size
 */
declare function getFontCacheSize(): number;
/**
 * Generate a simple hash for a paragraph block
 * Used as cache key to identify identical content
 */
declare function hashParagraphBlock(block: ParagraphBlock): string;
/**
 * Get cached paragraph measurement or return undefined
 */
declare function getCachedParagraphMeasure(block: ParagraphBlock, maxWidth: number): ParagraphMeasure | undefined;
/**
 * Store paragraph measurement in cache
 */
declare function setCachedParagraphMeasure(block: ParagraphBlock, maxWidth: number, measure: ParagraphMeasure): void;
/**
 * Clear the paragraph measure cache
 */
declare function clearParagraphMeasureCache(): void;
/**
 * Set the maximum size of the paragraph measure cache
 */
declare function setParagraphCacheSize(size: number): void;
/**
 * Get current paragraph measure cache size
 */
declare function getParagraphCacheSize(): number;
/**
 * Clear all measurement caches
 * Call when fonts change, page width changes, or for testing
 */
declare function clearAllCaches(): void;
/**
 * Get total size of all caches
 */
declare function getTotalCacheSize(): number;

export { FloatingImageZone, type FontMetrics, type FontStyle, type MeasureParagraphOptions, type RunMeasurement, type TextMeasurement, buildFontString, clampFloatingWrapMargins, clearAllCaches, clearFontMetricsCache, clearParagraphMeasureCache, clearTextWidthCache, findCharacterAtX, getCachedFontMetrics, getCachedParagraphMeasure, getCachedTextWidth, getCanvasContext, getFontCacheSize, getFontMetrics, getParagraphCacheSize, getRunCharWidths, getTextCacheSize, getTotalCacheSize, getXForCharacter, halfPtToPx, hashParagraphBlock, measureParagraph, measureParagraphs, measureRun, measureText, measureTextWidth, ptToPx, pxToHalfPt, pxToPt, pxToTwips, resetCanvasContext, setCachedFontMetrics, setCachedParagraphMeasure, setCachedTextWidth, setFontCacheSize, setParagraphCacheSize, setTextCacheSize, twipsToPx };
