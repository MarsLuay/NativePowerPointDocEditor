import { Node } from 'prosemirror-model';

/**
 * Pure ProseMirror paragraph/text lookup helpers shared by the React and
 * Vue adapters.
 *
 * "Vanilla view" text extraction skips text inside `insertion` marks
 * (tracked-change additions that aren't accepted yet) so the agent's view
 * of the document matches what `addComment` / `proposeChange` can anchor
 * to. Tracked deletions stay included — they're still in the doc until
 * accepted.
 *
 * Previously duplicated at
 * `packages/react/src/components/DocxEditor/internals/{pmAnchors,vanillaText}.ts`
 * and `packages/vue/src/utils/paraTextHelpers.ts`. Both adapters now
 * re-export from here.
 */

/**
 * PM position range for a paragraph identified by Word `w14:paraId`.
 * Stable across edits — inverse of `formatContentForLLM`'s `[paraId]` line tag.
 *
 * Returns inclusive `from` (position before the textblock) and exclusive
 * `to` (`from + nodeSize`). Text content lives in `[from + 1, to - 1]`.
 */
declare function findParaIdRange(doc: Node, paraId: string): {
    from: number;
    to: number;
} | null;
/** Text of a single PM node (typically a paragraph), vanilla view. */
declare function getVanillaNodeText(node: Node): string;
/** Text between two doc positions, vanilla view. */
declare function getVanillaTextBetween(doc: Node, from: number, to: number): string;
/**
 * Find `searchText` within a PM paragraph range and return its position.
 *
 * Returns null if:
 *   - searchText is empty
 *   - searchText is not found
 *   - searchText appears more than once (ambiguous; caller disambiguates)
 *
 * The fullText is built from PM text nodes only and matches the vanilla
 * view the agent reads via `read_document`: tracked insertions are
 * excluded (not in the doc yet), tracked deletions are included (still
 * in the doc until accepted), and comment markers are stripped.
 */
declare function findTextInPmParagraph(doc: Node, paragraphFrom: number, paragraphTo: number, searchText: string): {
    from: number;
    to: number;
} | null;

export { findParaIdRange, findTextInPmParagraph, getVanillaNodeText, getVanillaTextBetween };
