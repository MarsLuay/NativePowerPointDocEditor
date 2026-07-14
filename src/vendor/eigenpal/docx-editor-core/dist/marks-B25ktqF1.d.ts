import { T as ThemeColorSlot } from './colors-C3vA7HUU.js';
import { U as UnderlineStyle } from './formatting-JhqWT_XM.js';

/**
 * ProseMirror Mark Type Interfaces
 *
 * Type definitions for mark attributes used by conversion modules,
 * extensions, and other consumers. MarkSpec definitions have moved
 * to the extension system (extensions/marks/).
 */

/**
 * Text color mark attributes
 */
interface TextColorAttrs {
    rgb?: string;
    themeColor?: ThemeColorSlot;
    themeTint?: string;
    themeShade?: string;
}
/**
 * Underline mark attributes
 */
interface UnderlineAttrs {
    style?: UnderlineStyle;
    color?: TextColorAttrs;
}
/**
 * Font size mark attributes
 */
interface FontSizeAttrs {
    size?: number | null;
    sizeCs?: number | null;
}
/**
 * Font family mark attributes
 */
interface FontFamilyAttrs {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    cs?: string;
    asciiTheme?: string;
    hAnsiTheme?: string;
    eastAsiaTheme?: string;
    csTheme?: string;
}
/**
 * Hyperlink mark attributes
 */
interface HyperlinkAttrs {
    href: string;
    tooltip?: string;
    rId?: string;
}

export type { FontFamilyAttrs as F, HyperlinkAttrs as H, TextColorAttrs as T, UnderlineAttrs as U, FontSizeAttrs as a };
