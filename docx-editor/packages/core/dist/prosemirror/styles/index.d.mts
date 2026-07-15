/**
 * Style utilities for ProseMirror editor
 * @packageDocumentation
 * @public
 */
import { S as StyleResolver } from '../../styleResolver-B2Oxc8Lk.mjs';
export { R as ResolvedParagraphStyle, c as createStyleResolver } from '../../styleResolver-B2Oxc8Lk.mjs';
import { f as TableLook } from '../../formatting-DFtuRFQY.mjs';
import '../../styles-Diw0MASy.mjs';
import '../../colors-C3vA7HUU.mjs';

/**
 * Default-table-style resolution for newly created tables.
 *
 * When the editor (toolbar) or the agent inserts a brand-new table, it
 * should adopt the document's table styling rather than a generic black
 * grid. This module owns the precedence for picking that styleId so both
 * insert paths stay in lockstep.
 *
 * Precedence (most specific first):
 *   1. `w:defaultTableStyle` (settings.xml §17.15.1.44) — the styleId the
 *      document explicitly designates for newly created tables. A template
 *      author sets this so inserted tables match the template.
 *   2. The type-default table style (`w:default="1"`, §17.7.4.18), unless it
 *      is the no-op base style ("Table Normal" / styleId "TableNormal"),
 *      which carries no visual formatting and is the reset every table
 *      inherits from regardless.
 *   3. None — the caller falls back to its own default (a thin black border).
 *
 * The TableGrid builtin is intentionally NOT consulted by name: that would
 * be a guess rather than the document's declared default, and would change
 * insert behavior for documents that happen to ship the style but don't use
 * it. Templates express intent through the two spec elements above.
 */

/**
 * `tblLook` Word applies to a brand-new table (val="04A0"): style the first
 * row and first column, enable horizontal banding, disable vertical banding.
 * Setting this lets a resolved table style paint its header-row and banding
 * conditional formatting on the inserted table, matching how Word renders a
 * fresh table created with that style.
 */
declare const DEFAULT_NEW_TABLE_LOOK: TableLook;
/**
 * Resolve the styleId a newly inserted table should adopt, or `undefined`
 * when the document declares no usable default table style (caller falls
 * back to a plain border).
 *
 * @param defaultTableStyleId - `w:defaultTableStyle` from settings.xml
 * @param resolver - the document's style resolver (null when unavailable)
 */
declare function resolvePreferredNewTableStyleId(defaultTableStyleId: string | undefined | null, resolver: StyleResolver | null | undefined): string | undefined;

export { DEFAULT_NEW_TABLE_LOOK, StyleResolver, resolvePreferredNewTableStyleId };
