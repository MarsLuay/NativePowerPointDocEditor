import { EditorView } from 'prosemirror-view';
import { C as Comment } from '../content-BZ9rYecc.mjs';
import { CommentIdAllocator } from './commentIdAllocator.mjs';
import '../formatting-DFtuRFQY.mjs';
import '../colors-C3vA7HUU.mjs';
import '../docx/wrapTypes.mjs';
import '../lists-CyGxd5Y2.mjs';
import '../watermark-D90356ZM.mjs';

/**
 * Comment + tracked-change (proposeChange) PM transaction builders shared by
 * the React and Vue adapters. They locate the target range, apply the mark(s),
 * dispatch, and return a result. ID allocation lives in `commentIdAllocator.ts`
 * and is injected. Adapter-specific state mutation, event emission, and sidebar
 * UI stay in each adapter.
 */

/** Build a Comment object with a freshly-allocated ID. */
declare function createComment(allocator: CommentIdAllocator, text: string, authorName: string, parentId?: number): Comment;
interface AddCommentOptions {
    paraId: string;
    text: string;
    author: string;
    search?: string;
}
/**
 * Locate the comment range (paragraph, narrowed by `search`), add the comment
 * mark, dispatch, and return the created Comment. Returns null when the schema
 * lacks a comment mark or the range can't be resolved. The caller owns adding
 * the comment to its own state and showing the sidebar.
 */
declare function addCommentToRange(view: EditorView, options: AddCommentOptions, allocator: CommentIdAllocator): Comment | null;
interface ProposeChangeOptions {
    paraId: string;
    search: string;
    replaceWith: string;
    author: string;
}
/**
 * Apply a tracked change (insertion / deletion / replace) to the located range.
 * `search === ''` means a pure insertion at the paragraph end; `replaceWith ===
 * ''` means a pure deletion. Refuses to layer onto an existing tracked change.
 * Returns true when a change was dispatched, false otherwise.
 */
declare function applyProposedChange(view: EditorView, options: ProposeChangeOptions, allocator: CommentIdAllocator): boolean;

export { type AddCommentOptions, type ProposeChangeOptions, addCommentToRange, applyProposedChange, createComment };
