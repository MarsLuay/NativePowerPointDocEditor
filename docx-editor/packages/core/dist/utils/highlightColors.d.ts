/**
 * OOXML highlight color name <-> hex mapping. Used by both adapter
 * toolbars when serialising highlight selections to the
 * `w:highlight` attribute (which only accepts named colors).
 *
 * Lifted from packages/react/src/components/toolbarUtils.ts so both
 * sides share the same canonical table.
 * @packageDocumentation
 * @public
 */
declare const HIGHLIGHT_HEX_TO_NAME: Record<string, string>;
declare function mapHexToHighlightName(hex: string): string | null;

export { HIGHLIGHT_HEX_TO_NAME, mapHexToHighlightName };
