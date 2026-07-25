import { describe, expect, test } from 'bun:test';
import {
  mergeParagraphAttrsWithOriginalFormatting,
  originalFormattingAfterApplyStyle,
} from '../syncOriginalFormatting';

describe('syncOriginalFormatting', () => {
  test('writes spacing/indent into _originalFormatting on edit', () => {
    const next = mergeParagraphAttrsWithOriginalFormatting(
      {
        spaceBefore: 0,
        _originalFormatting: { styleId: 'Normal', spaceBefore: 120 },
      },
      { spaceBefore: 240, indentLeft: 720 }
    );

    expect(next.spaceBefore).toBe(240);
    expect(next.indentLeft).toBe(720);
    expect(next._originalFormatting).toEqual({
      styleId: 'Normal',
      spaceBefore: 240,
      indentLeft: 720,
    });
  });

  test('null clears direct indent from _originalFormatting', () => {
    const next = mergeParagraphAttrsWithOriginalFormatting(
      {
        indentLeft: 720,
        _originalFormatting: { indentLeft: 720, indentFirstLine: 360, styleId: 'Normal' },
      },
      { indentLeft: null, indentFirstLine: null, hangingIndent: null }
    );

    expect(next.indentLeft).toBeNull();
    expect(next._originalFormatting).toEqual({ styleId: 'Normal' });
  });

  test('applyStyle clears stale direct pPr/rPr but keeps styleId', () => {
    const next = originalFormattingAfterApplyStyle(
      {
        styleId: 'Heading1',
        spaceBefore: 240,
        lineSpacing: 480,
        runProperties: { bold: true, fontSize: 28 },
        widowControl: true,
      },
      'Normal'
    );

    expect(next).toEqual({
      styleId: 'Normal',
      widowControl: true,
    });
  });
});
