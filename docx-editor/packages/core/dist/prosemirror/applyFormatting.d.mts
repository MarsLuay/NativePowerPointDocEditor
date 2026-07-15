import { EditorView } from 'prosemirror-view';
import { S as StyleResolver } from '../styleResolver-B2Oxc8Lk.mjs';
import { N as NumberingMap } from '../numberingParser-nSPZlzKr.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../colors-C3vA7HUU.mjs';
import '../styles-Diw0MASy.mjs';
import '../lists-CyGxd5Y2.mjs';

/**
 * Agent-facing edit operations shared by the React and Vue adapters.
 *
 * `applyFormatting` maps a mark-toggle request (bold/italic/underline/strike/
 * color/highlight/fontSize/fontFamily) onto a PM transaction over a paragraph
 * range located by `paraId` (+ optional `search`). `setParagraphStyle` applies
 * a named paragraph style to that range. `insertBreak` inserts a page or
 * section break after the paragraph located by `paraId`.
 *
 * All take the `EditorView` as a parameter. `setParagraphStyle` takes the
 * style resolver as an injected dependency so each adapter keeps its own
 * resolver-sourcing strategy (React caches per styles object; Vue rebuilds).
 *
 * Previously duplicated byte-for-byte at
 * `packages/react/.../useDocxEditorRefApi.ts` and
 * `packages/vue/.../useFormattingActions.ts`.
 */

interface ApplyFormattingOptions {
    paraId: string;
    search?: string;
    marks: {
        bold?: boolean;
        italic?: boolean;
        underline?: boolean | {
            style?: string;
        };
        strike?: boolean;
        color?: {
            rgb?: string;
            themeColor?: string;
        };
        highlight?: string;
        fontSize?: number;
        fontFamily?: {
            ascii?: string;
            hAnsi?: string;
        };
    };
}
/**
 * Apply mark toggles to a paragraph range. Returns false when the paraId /
 * search can't be resolved; true (a no-op) when the resolved range is empty.
 */
declare function applyFormatting(view: EditorView, options: ApplyFormattingOptions): boolean;
/**
 * Apply a named paragraph style to the paragraph identified by `paraId`.
 *
 * The style resolver is injected: when present, unknown styleIds are rejected
 * (the agent gets a clear error instead of a silently-broken `<w:pStyle>`), and
 * the resolved paragraph/run formatting is threaded into `applyStyle`. Without
 * a resolver (no styles loaded) the styleId is applied as-is. Returns false
 * when the paraId can't be resolved or the styleId is unknown.
 */
declare function setParagraphStyle(view: EditorView, options: {
    paraId: string;
    styleId: string;
}, deps: {
    styleResolver: StyleResolver | null;
    numbering?: NumberingMap | null;
}): boolean;
/** Kind of break `insertBreak` can insert after a paragraph. */
type BreakKind = 'page' | 'sectionNextPage' | 'sectionContinuous';
interface InsertBreakOptions {
    paraId: string;
    type: BreakKind;
}
/**
 * Insert a page or section break after the paragraph identified by `paraId`.
 *
 * This is the agent edit path. It produces the same document shape as the
 * headless `reviewerBridge.insertBreak` (so a given `insert_break` call yields
 * identical output whether the agent drives a live editor or the headless
 * reviewer) — which is intentionally leaner than the interactive Insert > Break
 * menu command:
 *
 * - `page`: insert a single page-break node right after the target paragraph.
 *   `fromProseDoc` converts it to a break-run paragraph, matching the headless
 *   model. No trailing empty paragraph (the menu command adds one for caret
 *   placement; the agent keeps the caret where it was).
 * - `sectionNextPage` / `sectionContinuous`: a section break is the `sectPr`
 *   carried by the section's *last* paragraph, so mark the target paragraph
 *   directly. No new block — the existing following paragraph starts the new
 *   section.
 *
 * The user's selection is preserved (mapped through the edit) rather than
 * following the inserted break. Returns false when the paraId can't be resolved,
 * the target isn't a top-level paragraph, or `type` is unknown.
 */
declare function insertBreak(view: EditorView, options: InsertBreakOptions): boolean;

export { type ApplyFormattingOptions, type BreakKind, type InsertBreakOptions, applyFormatting, insertBreak, setParagraphStyle };
