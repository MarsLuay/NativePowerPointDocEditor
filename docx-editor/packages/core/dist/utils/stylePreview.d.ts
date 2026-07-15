/**
 * Paragraph-style preview + option resolution — shared between the React and
 * Vue toolbars so the style-picker dropdown looks and behaves identically.
 *
 * Pure logic only: no i18n and no framework CSS types. The returned preview is
 * a plain `{ fontSize, lineHeight, fontWeight?, fontStyle?, color? }` object,
 * which is structurally assignable to both React's `CSSProperties` and Vue's
 * inline-style record, so neither adapter needs a cast. Name localization stays
 * in the adapters (they own the i18n `t()` boundary).
 * @packageDocumentation
 * @public
 */
import { a as Style } from '../styles-2J4U-Lgk.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';

/**
 * Paragraph-style preview + option resolution — shared between the React and
 * Vue toolbars so the style-picker dropdown looks and behaves identically.
 *
 * Pure logic only: no i18n and no framework CSS types. The returned preview is
 * a plain `{ fontSize, lineHeight, fontWeight?, fontStyle?, color? }` object,
 * which is structurally assignable to both React's `CSSProperties` and Vue's
 * inline-style record, so neither adapter needs a cast. Name localization stays
 * in the adapters (they own the i18n `t()` boundary).
 * @packageDocumentation
 * @public
 */

/**
 * Inline preview style for a paragraph-style dropdown item.
 * @public
 */
interface StylePreviewProps {
    fontSize: string;
    lineHeight: string;
    fontWeight?: 'bold';
    fontStyle?: 'italic';
    color?: string;
}
/**
 * A resolved paragraph-style option (document order, with extracted visual
 * fields). Adapters add their own localized label on top of `name`.
 * @public
 */
interface ResolvedStyleOption {
    styleId: string;
    name: string;
    priority: number;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
}
/** Dropdown preview sizes (px) for well-known styles, matching Google Docs. @public */
declare const STYLE_PREVIEW_SIZES: Record<string, number>;
/** Styles rendered bold in the dropdown preview by default. @public */
declare const STYLE_DEFAULT_BOLD: ReadonlySet<string>;
/** Default preview text color (hex, no `#`) for styles that don't set one. @public */
declare const STYLE_DEFAULT_COLOR: Record<string, string>;
/** Google-Docs heading preview color. @public */
declare const HEADING_COLOR = "#4a6c8c";
/**
 * Build the inline preview style for a style option. `fontSize` is in
 * half-points (OOXML); known styles use {@link STYLE_PREVIEW_SIZES} instead.
 * @public
 */
declare function getStylePreviewProps(input: {
    styleId: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
}): StylePreviewProps;
/**
 * Filter the document's styles to the visible paragraph styles, extract their
 * visual fields, and sort by UI priority — the list a toolbar style picker
 * shows. Returns `[]` when there are no styles (adapters fall back to presets).
 * @public
 */
declare function resolveParagraphStyleOptions(styles: Style[] | undefined): ResolvedStyleOption[];

export { HEADING_COLOR, type ResolvedStyleOption, STYLE_DEFAULT_BOLD, STYLE_DEFAULT_COLOR, STYLE_PREVIEW_SIZES, type StylePreviewProps, getStylePreviewProps, resolveParagraphStyleOptions };
