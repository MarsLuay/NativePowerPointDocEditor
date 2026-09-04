import { describe, expect, test } from 'bun:test';

import {
  adjacentParagraphBorders,
  paragraphBorderFlowInsets,
  resolveRenderedParagraphBorders,
} from './paragraphBorders';
import type { FlowBlock, ParagraphBlock } from './types';

const headingBorders = {
  bottom: { style: 'solid', width: 1, color: '#000000', space: 7 },
};

function para(id: string, borders?: ParagraphBlock['attrs']): ParagraphBlock {
  return {
    kind: 'paragraph',
    id,
    runs: [{ kind: 'text', text: id, fontSize: 14 }],
    attrs: borders,
  };
}

describe('paragraph border flow insets', () => {
  test('bottom rule reserves space plus stroke so the next paragraph cannot sit on the line', () => {
    const insets = paragraphBorderFlowInsets({
      bottom: { style: 'solid', width: 1, color: '#000', space: 7 },
    });
    expect(insets.top).toBe(0);
    expect(insets.bottom).toBe(8);
  });

  test('grouped interiors keep the shared bottom rule on the last paragraph only', () => {
    const shared = {
      top: { style: 'solid', width: 2, color: '#000' },
      bottom: { style: 'solid', width: 2, color: '#000', space: 9 },
    };
    const first = paragraphBorderFlowInsets(shared, undefined, shared);
    const last = paragraphBorderFlowInsets(shared, shared, undefined);
    expect(first.bottom).toBe(0);
    expect(last.bottom).toBe(11);
    expect(resolveRenderedParagraphBorders(shared, shared, undefined).bottom?.width).toBe(2);
  });

  test('adjacentParagraphBorders stops grouping at a non-paragraph', () => {
    const blocks: FlowBlock[] = [
      para('h', { borders: headingBorders }),
      { kind: 'pageBreak' },
      para('body'),
    ];
    expect(adjacentParagraphBorders(blocks, 0).next).toBeUndefined();
    expect(adjacentParagraphBorders(blocks, 2).prev).toBeUndefined();
  });
});
