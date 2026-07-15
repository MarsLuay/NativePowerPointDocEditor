import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { DOMSerializer } from 'prosemirror-model';

import { singletonManager } from '../../schema';
import {
  DOC_X_P_BLOCK_IMAGE_CLASS,
  DOC_X_P_HAS_FLOAT_CLASS,
  getParagraphImageLayoutClasses,
} from './paragraphImageLayout';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('paragraphImageLayout', () => {
  test('stamps docx-p-block-image when the only child is a block image', () => {
    const schema = singletonManager.getSchema();
    const image = schema.nodes.image.create({
      src: 'data:image/png;base64,aa',
      displayMode: 'block',
    });
    const paragraph = schema.nodes.paragraph.create(null, image);

    expect(getParagraphImageLayoutClasses(paragraph)).toEqual([DOC_X_P_BLOCK_IMAGE_CLASS]);

    const dom = DOMSerializer.fromSchema(schema).serializeNode(paragraph) as HTMLElement;
    expect(dom.classList.contains(DOC_X_P_BLOCK_IMAGE_CLASS)).toBe(true);
    expect(dom.classList.contains(DOC_X_P_HAS_FLOAT_CLASS)).toBe(false);
    expect(dom.getAttribute('style') || '').toMatch(/text-indent:\s*0/);
  });

  test('stamps docx-p-has-float when a floated image is present', () => {
    const schema = singletonManager.getSchema();
    const image = schema.nodes.image.create({
      src: 'data:image/png;base64,aa',
      displayMode: 'float',
      cssFloat: 'left',
    });
    const paragraph = schema.nodes.paragraph.create(null, [
      schema.text('wrap '),
      image,
    ]);

    expect(getParagraphImageLayoutClasses(paragraph)).toEqual([DOC_X_P_HAS_FLOAT_CLASS]);

    const dom = DOMSerializer.fromSchema(schema).serializeNode(paragraph) as HTMLElement;
    expect(dom.classList.contains(DOC_X_P_HAS_FLOAT_CLASS)).toBe(true);
    expect(dom.classList.contains(DOC_X_P_BLOCK_IMAGE_CLASS)).toBe(false);
  });

  test('does not stamp classes for inline images', () => {
    const schema = singletonManager.getSchema();
    const image = schema.nodes.image.create({
      src: 'data:image/png;base64,aa',
      displayMode: 'inline',
    });
    const paragraph = schema.nodes.paragraph.create(null, image);

    expect(getParagraphImageLayoutClasses(paragraph)).toEqual([]);
  });
});
