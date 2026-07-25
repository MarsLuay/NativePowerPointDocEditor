import { Node } from 'prosemirror-model';

/**
 * ProseMirror position immediately before the first textblock whose `paraId`
 * attribute equals `paraId` (Word `w14:paraId` / OOXML paragraph id).
 *
 * Match is strict string equality on `node.attrs.paraId`.
 */
declare function findStartPosForParaId(doc: Node, paraId: string): number | null;

/**
 * ProseMirror position range for the paragraph (or any textblock) whose
 * `paraId` attribute equals `paraId`. Returns the inclusive `from` and
 * exclusive `to` positions, plus the node, so callers can both target
 * the paragraph (e.g. addMark over its text range) and inspect it.
 *
 * `from` is the position immediately before the textblock; `to` is
 * `from + node.nodeSize`. The text content lives at `[from + 1, to - 1]`.
 *
 * Returns null if no textblock with that paraId exists.
 */
declare function findParagraphByParaId(doc: Node, paraId: string): {
    node: Node;
    from: number;
    to: number;
} | null;

export { findStartPosForParaId as a, findParagraphByParaId as f };
