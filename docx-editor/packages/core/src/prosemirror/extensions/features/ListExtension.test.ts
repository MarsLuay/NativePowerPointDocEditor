import { describe, expect, test } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';

import { exitListAtCaretStart } from './ListExtension';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        numPr: { default: null },
        listIsBullet: { default: null },
        listNumFmt: { default: null },
        listMarker: { default: null },
        indentLeft: { default: null },
        indentFirstLine: { default: null },
        hangingIndent: { default: null },
        _originalFormatting: { default: null },
      },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
  },
  marks: {},
});

describe('exitListAtCaretStart', () => {
  test('removes a bullet and its list indent without deleting text', () => {
    const doc = schema.node('doc', null, [
      schema.node(
        'paragraph',
        {
          numPr: { numId: 1, ilvl: 0 },
          listIsBullet: true,
          listNumFmt: 'bullet',
          listMarker: '•',
          indentLeft: 360,
          indentFirstLine: -360,
          hangingIndent: true,
        },
        schema.text('Windows'),
      ),
    ]);
    const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
    let transaction: Transaction | undefined;

    expect(exitListAtCaretStart()(state, (tr) => { transaction = tr; })).toBe(true);

    const paragraph = transaction?.doc.firstChild;
    expect(paragraph?.textContent).toBe('Windows');
    expect(paragraph?.attrs.numPr).toBeNull();
    expect(paragraph?.attrs.listIsBullet).toBeNull();
    expect(paragraph?.attrs.indentLeft).toBeNull();
    expect(paragraph?.attrs.indentFirstLine).toBeNull();
    expect(paragraph?.attrs.hangingIndent).toBeNull();
  });

  test('does not remove a bullet when the caret is inside its text', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { numPr: { numId: 1, ilvl: 0 } }, schema.text('Windows')),
    ]);
    const state = EditorState.create({ doc, selection: TextSelection.create(doc, 2) });

    expect(exitListAtCaretStart()(state)).toBe(false);
  });
});
