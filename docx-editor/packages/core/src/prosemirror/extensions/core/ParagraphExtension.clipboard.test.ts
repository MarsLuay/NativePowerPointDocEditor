import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { DOMParser as ProseMirrorDOMParser, DOMSerializer } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';

import { clearParagraphBorders, setParagraphBottomBorder } from '../../commands';
import { singletonManager } from '../../schema';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const schema = singletonManager.getSchema();

function parseParagraph(dom: HTMLElement) {
  const host = document.createElement('div');
  host.appendChild(dom);
  return ProseMirrorDOMParser.fromSchema(schema).parse(host).firstChild!;
}

describe('ParagraphExtension clipboard metadata', () => {
  test('round-trips a bullet paragraph through the HTML clipboard', () => {
    const paragraph = schema.nodes.paragraph.create(
      {
        numPr: { numId: 1, ilvl: 1 },
        listIsBullet: true,
        listNumFmt: 'bullet',
        listMarker: '•',
      },
      schema.text('Linux'),
    );

    const dom = DOMSerializer.fromSchema(schema).serializeNode(paragraph) as HTMLElement;
    expect(dom.dataset.listNumId).toBe('1');
    expect(dom.dataset.listLevel).toBe('1');
    expect(dom.dataset.listIsBullet).toBe('true');

    const parsed = parseParagraph(dom);
    expect(parsed.attrs.numPr).toEqual({ numId: 1, ilvl: 1 });
    expect(parsed.attrs.listIsBullet).toBe(true);
    expect(parsed.attrs.listNumFmt).toBe('bullet');
    expect(parsed.attrs.listMarker).toBe('•');
  });

  test('parses an external HTML bullet list as list paragraphs', () => {
    const host = document.createElement('div');
    host.innerHTML = '<ul><li>Windows</li><li>Linux</li></ul>';

    const parsed = ProseMirrorDOMParser.fromSchema(schema).parse(host);
    expect(parsed.childCount).toBe(2);
    for (let index = 0; index < parsed.childCount; index += 1) {
      const paragraph = parsed.child(index);
      expect(paragraph.attrs.numPr).toEqual({ numId: 1, ilvl: 0 });
      expect(paragraph.attrs.listIsBullet).toBe(true);
      expect(paragraph.attrs.listMarker).toBe('•');
    }
  });

  test('round-trips an empty paragraph font size through the HTML clipboard', () => {
    const paragraph = schema.nodes.paragraph.create({
      defaultTextFormatting: {
        fontSize: 36,
        fontSizeCs: 36,
        fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
      },
    });

    const dom = DOMSerializer.fromSchema(schema).serializeNode(paragraph) as HTMLElement;
    expect(dom.dataset.defaultTextFormatting).toBeTruthy();

    const parsed = parseParagraph(dom);
    expect(parsed.attrs.defaultTextFormatting).toEqual({
      fontSize: 36,
      fontSizeCs: 36,
      fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
    });
  });

  test('round-trips a heading bottom border through the HTML clipboard', () => {
    const paragraph = schema.nodes.paragraph.create(
      {
        borders: {
          bottom: { style: 'single', size: 6, space: 1, color: { rgb: '000000' } },
        },
      },
      schema.text('TECHNICAL SKILLS'),
    );

    const dom = DOMSerializer.fromSchema(schema).serializeNode(paragraph) as HTMLElement;
    expect(dom.dataset.paragraphBorders).toBeTruthy();
    expect(dom.style.borderBottomStyle).toBe('solid');

    const parsed = parseParagraph(dom);
    expect(parsed.attrs.borders).toEqual({
      bottom: { style: 'single', size: 6, space: 1, color: { rgb: '000000' } },
    });
  });

  test('removes paragraph borders without changing the paragraph content', () => {
    const paragraph = schema.nodes.paragraph.create(
      { borders: { bottom: { style: 'single', size: 6, color: { rgb: '000000' } } } },
      schema.text('TECHNICAL SKILLS'),
    );
    const doc = schema.nodes.doc.create(null, [paragraph]);
    const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
    let transaction: Transaction | undefined;

    expect(clearParagraphBorders(state, (tr) => { transaction = tr; })).toBe(true);
    expect(transaction?.doc.firstChild?.textContent).toBe('TECHNICAL SKILLS');
    expect(transaction?.doc.firstChild?.attrs.borders).toBeNull();
  });

  test('adds a bottom border without replacing other paragraph borders', () => {
    const paragraph = schema.nodes.paragraph.create(
      { borders: { top: { style: 'single', size: 4, color: { rgb: '000000' } } } },
      schema.text('TECHNICAL SKILLS'),
    );
    const doc = schema.nodes.doc.create(null, [paragraph]);
    const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
    let transaction: Transaction | undefined;

    expect(setParagraphBottomBorder()(state, (tr) => { transaction = tr; })).toBe(true);
    expect(transaction?.doc.firstChild?.attrs.borders).toEqual({
      top: { style: 'single', size: 4, color: { rgb: '000000' } },
      bottom: { style: 'single', size: 6, space: 1, color: { rgb: '000000' } },
    });
  });
});
