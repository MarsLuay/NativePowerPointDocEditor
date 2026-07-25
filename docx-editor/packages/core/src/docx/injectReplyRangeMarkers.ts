/**
 * Reply-range marker injection for serialization.
 *
 * Word / Pages / LibreOffice expect every comment in `comments.xml`
 * (including REPLY threads) to have matching `commentRangeStart` /
 * `commentRangeEnd` / `commentReference` markers in `document.xml`.
 * The PM document only stamps marks for the parent comment because
 * replies don't have their own visible range — they share the parent
 * thread's text. So before serialization we walk the body content and
 * synthesize parallel range markers for every reply.
 *
 * Two helpers, one per parent shape:
 * - `injectReplyRangeMarkers` — replies whose parent is another
 *   comment (regular threaded discussion). Finds the parent's
 *   `commentRangeStart`/`End` and adds parallel markers next to them.
 * - `injectTCReplyRangeMarkers` — replies whose parent is a tracked
 *   change (insertion/deletion). Wraps the TC content with
 *   commentRange markers.
 *
 * Pre-#... this code lived inside React's DocxEditor.tsx; Vue had no
 * equivalent and so silently lost reply markers when saving collab
 * documents. Living in core means both adapters get it for free.
 */

import type { BlockContent, Comment, ParagraphContent } from '../types/content';

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for reply comments
 * that share their parent comment's text range.
 */
/**
 * Collect every `commentRangeStart` id currently present in the body.
 */
function collectCommentRangeStartIds(blocks: BlockContent[], into: Set<number>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'commentRangeStart') {
          into.add(item.id);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          collectCommentRangeStartIds(cell.content, into);
        }
      }
    }
  }
}

function paragraphContentHasText(item: ParagraphContent): boolean {
  switch (item.type) {
    case 'run':
      return item.content.some(
        (part) =>
          (part.type === 'text' && part.text.length > 0) ||
          (part.type === 'instrText' && part.text.length > 0) ||
          part.type === 'tab' ||
          part.type === 'symbol' ||
          part.type === 'softHyphen' ||
          part.type === 'noBreakHyphen'
      );
    case 'hyperlink':
      return item.children.some(paragraphContentHasText);
    case 'simpleField':
      return item.content.some(paragraphContentHasText);
    case 'complexField':
      return item.fieldResult.some(paragraphContentHasText);
    case 'inlineSdt':
      return item.content.some(paragraphContentHasText);
    case 'insertion':
    case 'deletion':
    case 'moveFrom':
    case 'moveTo':
      return item.content.some(paragraphContentHasText);
    case 'mathEquation':
      return (item.plainText?.length ?? 0) > 0;
    default:
      return false;
  }
}

function paragraphHasText(para: { content: ParagraphContent[] }): boolean {
  return para.content.some(paragraphContentHasText);
}

function findLastParagraph(
  blocks: BlockContent[],
  requireText = false
): { content: ParagraphContent[] } | null {
  for (let i = blocks.length - 1; i >= 0; i--) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      if (!requireText || paragraphHasText(block)) return block;
    } else if (block.type === 'table') {
      for (let r = block.rows.length - 1; r >= 0; r--) {
        for (let c = block.rows[r].cells.length - 1; c >= 0; c--) {
          const nested = findLastParagraph(block.rows[r].cells[c].content, requireText);
          if (nested) return nested;
        }
      }
    }
  }
  return null;
}

function findCommentRangeIdsWithoutText(blocks: BlockContent[], into: Set<number>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      const starts = new Map<number, number>();
      for (let i = 0; i < block.content.length; i++) {
        const item = block.content[i];
        if (item.type === 'commentRangeStart') {
          starts.set(item.id, i);
        } else if (item.type === 'commentRangeEnd') {
          const startIdx = starts.get(item.id);
          if (startIdx == null) continue;
          const hasText = block.content
            .slice(startIdx + 1, i)
            .some(paragraphContentHasText);
          if (!hasText) into.add(item.id);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          findCommentRangeIdsWithoutText(cell.content, into);
        }
      }
    }
  }
}

function removeCommentRangeMarkers(blocks: BlockContent[], ids: Set<number>): void {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      block.content = block.content.filter(
        (item) =>
          !(
            (item.type === 'commentRangeStart' || item.type === 'commentRangeEnd') &&
            ids.has(item.id)
          )
      );
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          removeCommentRangeMarkers(cell.content, ids);
        }
      }
    }
  }
}

/**
 * Re-wrap comment ranges that enclose no text (start/end around empty runs).
 * `toProseDoc` only stamps PM marks on text runs, so empty ranges vanish on
 * reload even when `comments.xml` is intact.
 */
export function repairEmptyCommentRanges(content: BlockContent[]): void {
  const emptyIds = new Set<number>();
  findCommentRangeIdsWithoutText(content, emptyIds);
  if (emptyIds.size === 0) return;

  removeCommentRangeMarkers(content, emptyIds);

  const target = findLastParagraph(content, true) ?? findLastParagraph(content, false);
  if (!target) return;

  let inner = target.content;
  const ids = [...emptyIds].sort((a, b) => a - b);
  for (let index = ids.length - 1; index >= 0; index--) {
    const id = ids[index];
    inner = [
      { type: 'commentRangeStart', id },
      ...inner,
      { type: 'commentRangeEnd', id },
    ];
  }
  target.content = inner;
}

/**
 * Re-anchor top-level comments that exist in `comments.xml` / React state
 * but have no `commentRangeStart` in the body. Without this, reload →
 * PM marks → orphan cleanup deletes the thread, and the next save can
 * rewrite `comments.xml` without anchors.
 *
 * Wraps the last body paragraph's existing inline content so `toProseDoc`
 * can stamp PM `comment` marks on enclosed runs. Bare start/end pairs
 * (zero-width) do not produce marks and were a silent reload failure mode.
 */
export function injectMissingTopLevelCommentRangeMarkers(
  content: BlockContent[],
  comments: Comment[]
): void {
  const topLevel = comments.filter((c) => c.parentId == null);
  if (topLevel.length === 0) return;

  // Fix empty (textless) ranges first so they look "present" but unload.
  repairEmptyCommentRanges(content);

  const existingIds = new Set<number>();
  collectCommentRangeStartIds(content, existingIds);

  const missing = topLevel.filter((c) => !existingIds.has(c.id));
  if (missing.length === 0) return;

  const lastParagraph = findLastParagraph(content, true) ?? findLastParagraph(content, false);
  if (!lastParagraph) return;

  // Nest ranges around existing content (outermost = first missing id).
  let inner = lastParagraph.content;
  for (let index = missing.length - 1; index >= 0; index--) {
    const id = missing[index].id;
    inner = [
      { type: 'commentRangeStart', id },
      ...inner,
      { type: 'commentRangeEnd', id },
    ];
  }
  lastParagraph.content = inner;
}

export function injectReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void {
  const replies = comments.filter((c) => c.parentId != null);
  if (replies.length === 0) return;

  // Build parentId → reply IDs map
  const replyIdsByParent = new Map<number, number[]>();
  for (const r of replies) {
    const arr = replyIdsByParent.get(r.parentId!);
    if (arr) arr.push(r.id);
    else replyIdsByParent.set(r.parentId!, [r.id]);
  }

  function walkBlocks(blocks: BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        // Skip paragraphs without any comment range markers
        if (
          !block.content.some((i) => i.type === 'commentRangeStart' || i.type === 'commentRangeEnd')
        )
          continue;
        const newItems: ParagraphContent[] = [];
        for (const item of block.content) {
          if (item.type === 'commentRangeStart') {
            newItems.push(item);
            const replyIds = replyIdsByParent.get(item.id);
            if (replyIds) {
              for (const rid of replyIds) {
                newItems.push({ type: 'commentRangeStart', id: rid });
              }
            }
          } else if (item.type === 'commentRangeEnd') {
            newItems.push(item);
            const replyIds = replyIdsByParent.get(item.id);
            if (replyIds) {
              for (const rid of replyIds) {
                newItems.push({ type: 'commentRangeEnd', id: rid });
              }
            }
          } else {
            newItems.push(item);
          }
        }
        block.content = newItems;
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkBlocks(cell.content);
          }
        }
      }
    }
  }

  walkBlocks(content);
}

/**
 * Inject `commentRangeStart`/`commentRangeEnd` for comments whose
 * parent is a tracked-change revision (insertion/deletion). The TC
 * content nodes don't carry the comment's range, so we wrap them.
 */
export function injectTCReplyRangeMarkers(content: BlockContent[], comments: Comment[]): void {
  const commentIds = new Set(comments.map((c) => c.id));
  const tcReplies = comments.filter((c) => c.parentId != null && !commentIds.has(c.parentId));
  if (tcReplies.length === 0) return;

  const replyIdsByRevision = new Map<number, number[]>();
  for (const r of tcReplies) {
    const arr = replyIdsByRevision.get(r.parentId!);
    if (arr) arr.push(r.id);
    else replyIdsByRevision.set(r.parentId!, [r.id]);
  }

  function walkBlocks(blocks: BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        const hasTC = block.content.some(
          (item) =>
            (item.type === 'insertion' || item.type === 'deletion') &&
            replyIdsByRevision.has(item.info.id)
        );
        if (!hasTC) continue;

        const newItems: ParagraphContent[] = [];
        const items = block.content;
        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          if (
            (item.type === 'insertion' || item.type === 'deletion') &&
            replyIdsByRevision.has(item.info.id)
          ) {
            const replyIds = replyIdsByRevision.get(item.info.id)!;
            for (const rid of replyIds) {
              newItems.push({ type: 'commentRangeStart', id: rid });
            }
            newItems.push(item);
            // Adjacent del+ins replacement pair share author+date —
            // include the second half inside the comment range so we
            // don't break del-ins adjacency in the saved doc.
            const next = items[i + 1];
            if (
              next &&
              (next.type === 'insertion' || next.type === 'deletion') &&
              next.type !== item.type &&
              next.info.author === item.info.author &&
              next.info.date === item.info.date
            ) {
              newItems.push(next);
              i++;
            }
            for (const rid of replyIds) {
              newItems.push({ type: 'commentRangeEnd', id: rid });
            }
          } else {
            newItems.push(item);
          }
        }
        block.content = newItems;
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walkBlocks(cell.content);
          }
        }
      }
    }
  }

  walkBlocks(content);
}
