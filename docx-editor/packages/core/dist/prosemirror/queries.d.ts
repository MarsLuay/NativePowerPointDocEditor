import { EditorView } from 'prosemirror-view';
import { Node } from 'prosemirror-model';
import { Layout } from '../layout-engine/types.js';
import '../content-B8ScSBzC.js';
import '../formatting-JhqWT_XM.js';
import '../colors-C3vA7HUU.js';
import '../docx/wrapTypes.js';
import '../lists-Bn29SzeS.js';
import '../watermark-D90356ZM.js';

/**
 * Pure ref-API query helpers — read-only inspectors over the PM document
 * and the paginated layout. Back the adapters' `findInDocument`,
 * `getSelectionInfo`, and `getPageContent` ref methods.
 *
 * Every function takes the PM view (or layout + view) as a parameter
 * instead of closing over a framework ref, so the React and Vue adapters
 * (and the future vanilla wrapper) share one implementation.
 */

/** A resolved PM position range — half-open `[from, to)` in PM coordinates. */
interface PmRange {
    from: number;
    to: number;
}
/**
 * Clamp a caller-supplied `[from, to]` range to a valid in-document span, or
 * return `null` when it cannot be made valid: non-integer, negative, reversed
 * (`to < from`), or a `from` past the document end. `to` is clamped to the
 * document size so an out-of-range end never makes `doc.resolve()` throw a
 * `RangeError`. Both adapters' `highlightRange` route raw caller positions
 * through this so the no-op contract holds identically.
 */
declare function clampRangeToDoc(doc: Node, from: number, to: number): PmRange | null;
/**
 * Resolve a `commentId` to the PM position range its `comment` mark
 * spans. Walks every inline node carrying a `comment` mark with the
 * matching id and returns the union range (earliest start → latest end),
 * so a comment whose range is interrupted by un-marked inline atoms still
 * resolves to a single span. Returns `null` when the id is no longer
 * present (the comment was deleted, or the marked text was removed) — the
 * caller distinguishes "scrolled" from "stale" on that signal.
 *
 * Pure read over `view.state`; no dispatch.
 */
declare function findCommentRange(view: EditorView | null, commentId: number): PmRange | null;
/**
 * Resolve a tracked-change `revisionId` to the PM position range of its
 * first site. Delegates to {@link extractTrackedChanges} so a coalesced
 * revision (sites scattered across paragraphs, replace pairs, Enter
 * chains) resolves to the same entry the sidebar shows. Matches on the
 * entry's primary `revisionId`, its `insertionRevisionId`, or any
 * `coalescedRevisionIds` member. Returns `null` when no entry carries the
 * id (the change was accepted/rejected/deleted) — the caller uses this to
 * show a "location no longer exists" affordance instead of a silent
 * no-op.
 *
 * Pure read over `view.state`; no dispatch.
 */
declare function findChangeRange(view: EditorView | null, revisionId: number): PmRange | null;
interface FindInDocumentMatch {
    paraId: string;
    match: string;
    before: string;
    after: string;
}
/**
 * Walk the PM doc looking for `query`. Returns up to `limit` matches —
 * one per paragraph (rejects paragraphs where the query appears more
 * than once, mirroring `findTextInPmParagraph`'s ambiguity guard so the
 * LLM gets a clearer error than a silent mistarget).
 */
declare function findInDocument(view: EditorView | null, query: string, opts?: {
    caseSensitive?: boolean;
    limit?: number;
}): FindInDocumentMatch[];
interface SelectionInfo {
    paraId: string | null;
    selectedText: string;
    paragraphText: string;
    before: string;
    after: string;
}
/**
 * Describe the current selection in agent-readable form — paraId of the
 * containing paragraph, the selected text, the full paragraph text, and
 * the leading/trailing slices. Vanilla view: insertion-marked text never
 * appears, matching what the agent reads and can anchor against.
 */
declare function getSelectionInfo(view: EditorView | null): SelectionInfo | null;
interface PageContent {
    pageNumber: number;
    text: string;
    paragraphs: Array<{
        paraId: string;
        text: string;
        styleId?: string;
    }>;
}
/**
 * Collect paragraphs visible on `pageNumber` (1-indexed) from the
 * paginated `layout`. Dedupes by paraId so paragraphs split across page
 * boundaries are reported once.
 */
declare function getPageContent(view: EditorView | null, layout: Layout | null, pageNumber: number): PageContent | null;

export { type FindInDocumentMatch, type PageContent, type PmRange, type SelectionInfo, clampRangeToDoc, findChangeRange, findCommentRange, findInDocument, getPageContent, getSelectionInfo };
