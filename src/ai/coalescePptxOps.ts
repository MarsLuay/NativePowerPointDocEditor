import type { DocumentOp } from './types';

/**
 * Merge consecutive same-slide `pptx.deleteShape` ops into one
 * `pptx.deleteShapes` payload so indices are resolved before any renumber.
 */
export function coalescePptxOps(ops: DocumentOp[]): DocumentOp[] {
  const coalesced: DocumentOp[] = [];
  let index = 0;
  while (index < ops.length) {
    const op = ops[index];
    if (!op || op.op !== 'pptx.deleteShape') {
      coalesced.push(op!);
      index += 1;
      continue;
    }

    const slideIndex = typeof op.slideIndex === 'number' ? op.slideIndex : null;
    const shapeIndexes: number[] = [];
    let cursor = index;
    while (cursor < ops.length) {
      const candidate = ops[cursor];
      if (
        !candidate
        || candidate.op !== 'pptx.deleteShape'
        || typeof candidate.slideIndex !== 'number'
        || candidate.slideIndex !== slideIndex
        || typeof candidate.shapeIndex !== 'number'
      ) {
        break;
      }
      shapeIndexes.push(candidate.shapeIndex);
      cursor += 1;
    }

    if (slideIndex === null || shapeIndexes.length <= 1) {
      coalesced.push(op);
      index += 1;
      continue;
    }

    coalesced.push({
      op: 'pptx.deleteShapes',
      slideIndex,
      shapeIndexes: [...new Set(shapeIndexes)].sort((left, right) => right - left),
    });
    index = cursor;
  }
  return coalesced;
}
