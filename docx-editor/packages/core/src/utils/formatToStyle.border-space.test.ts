import { describe, expect, test } from 'bun:test';

import { borderToStyle } from './formatToStyle';
import { formatPx, pointsToPixels } from './units';

describe('borderToStyle', () => {
  test('maps OOXML w:space onto CSS padding so the stroke is not inside the glyphs', () => {
    const style = borderToStyle(
      { style: 'single', size: 8, space: 5, color: { rgb: '000000' } },
      'Bottom'
    );
    expect(style.borderBottomStyle).toBe('solid');
    expect(style.paddingBottom).toBe(formatPx(pointsToPixels(5)));
  });
});
