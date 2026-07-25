/**
 * Empty-paragraph font/size must serialize to pPr/rPr so tab switch / reload
 * keeps the caret formatting. setMark syncs _originalFormatting.runProperties;
 * paragraphs without _originalFormatting emit runProperties from
 * defaultTextFormatting.
 */

import { describe, test, expect } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { convertPMParagraph } from '../fromProseDoc/paragraph';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        defaultTextFormatting: { default: null },
        _originalFormatting: { default: null },
        alignment: { default: null },
        styleId: { default: null },
        numPr: { default: null },
        numPrFromStyle: { default: null },
        pageBreakBefore: { default: null },
        bidi: { default: null },
        paraId: { default: null },
        textId: { default: null },
        bookmarks: { default: null },
        renderedPageBreakBefore: { default: null },
        pPrIns: { default: null },
        pPrDel: { default: null },
        pPrChange: { default: null },
        _sectionProperties: { default: null },
        sectionBreakType: { default: null },
        _originalRunBoundaries: { default: null },
        spaceBefore: { default: null },
        spaceAfter: { default: null },
        lineSpacing: { default: null },
        lineSpacingRule: { default: null },
        indentLeft: { default: null },
        indentRight: { default: null },
        indentFirstLine: { default: null },
        hangingIndent: { default: null },
        borders: { default: null },
        shading: { default: null },
        tabs: { default: null },
        outlineLevel: { default: null },
        contextualSpacing: { default: null },
      },
    },
    text: { group: 'inline' },
  },
});

describe('empty paragraph font serialize', () => {
  test('loaded empty para with synced _originalFormatting.runProperties emits pPr/rPr', () => {
    const para = schema.node('paragraph', {
      _originalFormatting: {
        alignment: 'left',
        runProperties: {
          fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
          fontSize: 48,
        },
      },
      defaultTextFormatting: {
        fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
        fontSize: 48,
      },
    }, []);

    const converted = convertPMParagraph(para);
    expect(converted.formatting?.runProperties).toEqual({
      fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
      fontSize: 48,
    });
  });

  test('new empty para with only defaultTextFormatting emits runProperties', () => {
    const para = schema.node('paragraph', {
      defaultTextFormatting: {
        fontFamily: { ascii: 'Verdana', hAnsi: 'Verdana' },
        fontSize: 36,
      },
    }, []);

    const converted = convertPMParagraph(para);
    expect(converted.formatting?.runProperties).toEqual({
      fontFamily: { ascii: 'Verdana', hAnsi: 'Verdana' },
      fontSize: 36,
    });
  });

  test('style-resolved defaultTextFormatting alone does not invent runProperties when original has none', () => {
    // Load paints caret from style-merged defaultTextFormatting, but
    // _originalFormatting.runProperties stays undefined until the user edits.
    const para = schema.node('paragraph', {
      _originalFormatting: { alignment: 'left' },
      defaultTextFormatting: {
        fontFamily: { ascii: 'Calibri', hAnsi: 'Calibri' },
        fontSize: 22,
      },
    }, []);

    const converted = convertPMParagraph(para);
    expect(converted.formatting?.runProperties).toBeUndefined();
  });
});
