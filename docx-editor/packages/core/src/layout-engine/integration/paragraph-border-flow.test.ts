import { describe, expect, test } from 'bun:test';

import { layoutDocument } from '../index';
import type { FlowBlock, Measure, ParagraphBlock, ParagraphFragment } from '../types';
import { makeLayoutOptions, makeLine, makeParagraphMeasure } from './helpers';

describe('paragraph border flow', () => {
  test('heading bottom rule is reserved in fragment height so body starts after the stroke', () => {
    const heading: ParagraphBlock = {
      kind: 'paragraph',
      id: 'heading',
      runs: [{ kind: 'text', text: 'SUMMARY', fontSize: 14 }],
      attrs: {
        borders: {
          bottom: { style: 'solid', width: 1, color: '#000000', space: 7 },
        },
      },
    };
    const body: ParagraphBlock = {
      kind: 'paragraph',
      id: 'body',
      runs: [{ kind: 'text', text: 'Computer Engineering transfer student', fontSize: 12 }],
    };
    const headingLine = 19;
    const bodyLine = 17;
    const blocks: FlowBlock[] = [heading, body];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 7, 200, headingLine)]),
      makeParagraphMeasure([makeLine(0, 0, 0, 20, 400, bodyLine)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());
    const fragments = layout.pages[0].fragments as ParagraphFragment[];
    expect(fragments).toHaveLength(2);

    const headingFragment = fragments[0];
    const bodyFragment = fragments[1];
    expect(headingFragment.height).toBe(headingLine + 7 + 1);
    expect(bodyFragment.y).toBe(headingFragment.y + headingFragment.height);
    expect(bodyFragment.y).toBeGreaterThanOrEqual(headingFragment.y + headingLine + 7 + 1);
  });
});
