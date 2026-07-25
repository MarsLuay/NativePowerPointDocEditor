/**
 * Walk the PM doc once and derive (a) the tracked-change list and (b) a
 * comment→revision overlap map for threading. Adjacent entries from the
 * same revision are merged; deletion+insertion pairs from the same
 * author/date become a single `replacement` entry (matches Word's UX
 * for replace ops).
 *
 * Pure function — no React, no Vue, no side effects. Single O(N) walk
 * over text nodes. Consumers building custom sidebars should prefer the
 * adapter-specific wrappers (`useTrackedChanges` in
 * `@npde/docx-editor-react/hooks` and
 * `@npde/docx-editor-vue/composables`), which add the memoization
 * and reactivity layer. Reach for the core function directly for
 * server-side analysis or test fixtures.
 *
 * @packageDocumentation
 * @public
 */
import { EditorState } from 'prosemirror-state';
import { TrackedChangeEntry } from '../../utils/comments.js';
import '../../content-B8ScSBzC.js';
import '../../formatting-JhqWT_XM.js';
import '../../colors-C3vA7HUU.js';
import '../../docx/wrapTypes.js';
import '../../lists-Bn29SzeS.js';
import '../../watermark-D90356ZM.js';

/**
 * Walk the PM doc once and derive (a) the tracked-change list and (b) a
 * comment→revision overlap map for threading. Adjacent entries from the
 * same revision are merged; deletion+insertion pairs from the same
 * author/date become a single `replacement` entry (matches Word's UX
 * for replace ops).
 *
 * Pure function — no React, no Vue, no side effects. Single O(N) walk
 * over text nodes. Consumers building custom sidebars should prefer the
 * adapter-specific wrappers (`useTrackedChanges` in
 * `@npde/docx-editor-react/hooks` and
 * `@npde/docx-editor-vue/composables`), which add the memoization
 * and reactivity layer. Reach for the core function directly for
 * server-side analysis or test fixtures.
 *
 * @packageDocumentation
 * @public
 */

/**
 * Output of {@link extractTrackedChanges}.
 *
 * @public
 */
interface TrackedChangesResult {
    /** Tracked-change entries, sorted by document position, with adjacent same-revision entries merged. */
    entries: TrackedChangeEntry[];
    /**
     * Map of `commentId -> revisionId` for comments whose range overlaps a tracked-change mark.
     * Consumers (DocxEditor's threading effect) use this to thread comments under their tracked change.
     */
    commentToRevision: Map<number, number>;
}
/**
 * Walk the PM doc and extract every tracked change as a flat list of
 * `TrackedChangeEntry` plus a comment→revision overlap map. Adjacent
 * inline marks coalesce by `(type, revisionId, author, date)`; a
 * deletion immediately followed by an insertion (same author + same
 * date) collapses into a single `replacement` entry; paragraph-mark
 * cards (`paragraphMarkInsertion` / `paragraphMarkDeletion`) are
 * hidden when an inline entry already covers their revision triple
 * (one Accept clears every site of one conceptual change).
 *
 * Pure and deterministic. Returns `EMPTY_RESULT` on null state.
 *
 * @example
 * ```ts
 * import { extractTrackedChanges } from '@npde/docx-editor-core/prosemirror/utils/extractTrackedChanges';
 *
 * const { entries, commentToRevision } = extractTrackedChanges(view.state);
 * for (const e of entries) {
 *   console.log(e.type, e.author, e.text);
 * }
 * ```
 *
 * @public
 */
declare function extractTrackedChanges(state: EditorState | null): TrackedChangesResult;

export { type TrackedChangesResult, extractTrackedChanges };
