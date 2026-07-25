/**
 * ProseMirror to FlowBlock Converter
 *
 * Converts a ProseMirror document into FlowBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 *
 * The deep import `@npde/.../layout-bridge/toFlowBlocks` is part of the
 * public surface (Vue adapter + tests), so the per-domain helpers under
 * ./toFlowBlocks/ are re-exported from here to keep that path stable.
 * @packageDocumentation
 * @public
 */
import { Node } from 'prosemirror-model';
import { BorderStyle, FlowBlock } from '../layout-engine/types.mjs';
import { T as Theme } from '../styles-Diw0MASy.mjs';
import { N as NumberFormat } from '../lists-CyGxd5Y2.mjs';
import '../content-BZ9rYecc.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../colors-C3vA7HUU.mjs';
import '../docx/wrapTypes.mjs';
import '../watermark-D90356ZM.mjs';

/**
 * Shared types and primitive helpers used across toFlowBlocks sub-modules.
 */

/**
 * Options for the conversion.
 */
type ToFlowBlocksOptions = {
    /** Default font family. */
    defaultFont?: string;
    /** Default font size in points. */
    defaultSize?: number;
    /** Theme for resolving theme colors. */
    theme?: Theme | null;
    /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
    pageContentHeight?: number;
    /**
     * Document-wide `w:defaultTabStop` (§17.6.13) in twips. Stamped onto each
     * paragraph's attrs so the layout-time list-marker helper can snap body
     * text to the default tab grid when no custom `w:tabs` are defined.
     * Default 720 twips (Word's spec default).
     */
    defaultTabStopTwips?: number;
    /**
     * @internal Allocated by toFlowBlocks() and threaded through table /
     * text-box conversion so list numbering stays continuous across containers.
     * Keyed by abstractNumId when known (ECMA-376 §17.9.18: numIds sharing one
     * abstractNum share counter state); falls back to numId.
     */
    listCounters?: Map<number, number[]>;
    /**
     * @internal Tracks `${numId}:${ilvl}` pairs whose startOverride has already
     * been applied. Per ECMA-376 §17.9.27 the override fires the first time
     * each level of a numId is encountered, so a numId with overrides on
     * multiple ilvls fires each one independently.
     */
    listSeenNumIds?: Set<string>;
};
/**
 * Reset the block ID counter (useful for testing).
 */
declare function resetBlockIdCounter(): void;

/**
 * Border Conversion
 *
 * Shared OOXML BorderSpec → layout-engine BorderStyle conversion, used by
 * paragraph borders, table cell borders, and header/footer borders.
 */

/**
 * Convert an OOXML BorderSpec to a layout-engine BorderStyle.
 * Shared by paragraph borders, cell borders, and header/footer borders.
 */
declare function convertBorderSpecToLayout(border: {
    style?: string;
    size?: number;
    space?: number;
    color?: {
        rgb?: string;
        themeColor?: string;
        themeTint?: string;
        themeShade?: string;
    };
}, theme?: Theme | null): BorderStyle | undefined;

/**
 * List Marker Resolution
 *
 * Helpers for rendering OOXML list markers from the counter stack:
 *   - format numbers as decimal/roman/letter per ECMA-376 §17.9.16 numFmt
 *   - resolve lvlText templates ("%1.%2.") against the counter stack
 *   - drive the per-paragraph counter increment, including startOverride.
 */

/**
 * Resolve an OOXML lvlText template like "%1.%2." against the counter stack
 * and per-level numFmt list (ECMA-376 §17.9.11).
 *
 * When a referenced counter has no value yet (e.g. "%2" referenced from a
 * level-0 paragraph), the placeholder AND the punctuation immediately
 * following it are dropped — matches Word's behavior so "%1.%2." renders
 * "1." rather than "1..".
 *
 * Exported for unit testing.
 */
declare function resolveListTemplate(template: string, counters: number[], levelNumFmts: NumberFormat[] | undefined): string;

/**
 * ProseMirror to FlowBlock Converter
 *
 * Converts a ProseMirror document into FlowBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 *
 * The deep import `@npde/.../layout-bridge/toFlowBlocks` is part of the
 * public surface (Vue adapter + tests), so the per-domain helpers under
 * ./toFlowBlocks/ are re-exported from here to keep that path stable.
 * @packageDocumentation
 * @public
 */

/**
 * Convert a ProseMirror document to FlowBlock array.
 *
 * Walks the document tree, converting each node to the appropriate block type.
 * Tracks pmStart/pmEnd positions for each block for click-to-position mapping.
 */
declare function toFlowBlocks(doc: Node, options?: ToFlowBlocksOptions): FlowBlock[];

export { type ToFlowBlocksOptions, convertBorderSpecToLayout, resetBlockIdCounter, resolveListTemplate, toFlowBlocks };
