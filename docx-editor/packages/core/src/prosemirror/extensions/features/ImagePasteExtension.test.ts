/**
 * External image drop helpers for ImagePasteExtension.
 *
 * Pins: dragover acceptance via Files/types, drop inserts at pointer pos,
 * non-image drops ignored, internal image reposition not claimed.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

import { singletonManager } from '../../schema';
import {
  insertImageFiles,
  isInternalImageDragging,
  resolveExternalImageDropPos,
} from './ImagePasteExtension';
import { dataTransferLooksLikeExternalImageDrop } from '../../../utils/clipboard';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const schema = singletonManager.getSchema();

/** Tiny valid 1×1 PNG. */
const PNG_BYTES = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  ),
  (c) => c.charCodeAt(0)
);

function makeView(text = 'hello world') {
  const para = schema.nodes.paragraph.create(null, [schema.text(text)]);
  const doc = schema.nodes.doc.create(null, [para]);
  const dom = document.createElement('div');
  const view = {
    dom,
    state: EditorState.create({
      schema,
      doc,
      selection: TextSelection.create(doc, 1),
    }),
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
    focus() {},
    posAtCoords: null as
      | null
      | ((coords: { left: number; top: number }) => { pos: number; inside: number } | null),
  };
  view.posAtCoords = () => ({ pos: 4, inside: 1 });
  return view as unknown as EditorView & {
    state: EditorState;
    dom: HTMLDivElement;
    posAtCoords: (coords: { left: number; top: number }) => { pos: number; inside: number } | null;
  };
}

function countImages(state: EditorState): number {
  let count = 0;
  state.doc.descendants((node) => {
    if (node.type.name === 'image') count += 1;
  });
  return count;
}

function imagePositions(state: EditorState): number[] {
  const positions: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'image') positions.push(pos);
  });
  return positions;
}

describe('isInternalImageDragging', () => {
  test('true when pm-image-dragging is on the editor DOM', () => {
    const view = makeView();
    expect(isInternalImageDragging(view)).toBe(false);
    view.dom.classList.add('pm-image-dragging');
    expect(isInternalImageDragging(view)).toBe(true);
  });
});

describe('dataTransferLooksLikeExternalImageDrop', () => {
  test('true for image Files on drop DataTransfer', () => {
    const imageFile = new File([PNG_BYTES], 'photo.png', { type: 'image/png' });
    const dataTransfer = {
      files: [imageFile],
      items: [],
      types: ['Files'],
    } as unknown as DataTransfer;
    expect(dataTransferLooksLikeExternalImageDrop(dataTransfer)).toBe(true);
  });

  test('true for Files type during dragover when files list is empty', () => {
    const dataTransfer = {
      files: [],
      items: [],
      types: ['Files'],
    } as unknown as DataTransfer;
    expect(dataTransferLooksLikeExternalImageDrop(dataTransfer)).toBe(true);
  });

  test('true when items advertise image/* without getAsFile payload yet', () => {
    const dataTransfer = {
      files: [],
      items: [{ kind: 'file', type: 'image/png', getAsFile: () => null }],
      types: ['Files'],
    } as unknown as DataTransfer;
    expect(dataTransferLooksLikeExternalImageDrop(dataTransfer)).toBe(true);
  });

  test('false for null / non-file transfers', () => {
    expect(dataTransferLooksLikeExternalImageDrop(null)).toBe(false);
    const textOnly = {
      files: [],
      items: [],
      types: ['text/plain'],
    } as unknown as DataTransfer;
    expect(dataTransferLooksLikeExternalImageDrop(textOnly)).toBe(false);
  });
});

describe('resolveExternalImageDropPos', () => {
  test('maps pointer coords to a near text selection position', () => {
    const view = makeView();
    expect(resolveExternalImageDropPos(view, 10, 20)).toBe(4);
  });

  test('returns null when posAtCoords cannot map', () => {
    const view = makeView();
    view.posAtCoords = () => null;
    expect(resolveExternalImageDropPos(view, 10, 20)).toBeNull();
  });
});

describe('insertImageFiles at drop pos', () => {
  test('inserts an image near the captured start position', async () => {
    const OriginalImage = globalThis.Image;
    globalThis.Image = class {
      naturalWidth = 1;
      naturalHeight = 1;
      onload: (() => void) | null = null;
      onerror: ((error?: unknown) => void) | null = null;
      set src(_value: string) {
        queueMicrotask(() => this.onload?.());
      }
    } as unknown as typeof Image;

    try {
      const view = makeView();
      // Move selection away from drop target so a race would land elsewhere.
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
      const dropPos = 7; // inside "hello world" after "hello "
      const file = new File([PNG_BYTES], 'drop.png', { type: 'image/png' });

      await insertImageFiles(view, [file], dropPos);

      expect(countImages(view.state)).toBe(1);
      const [imgPos] = imagePositions(view.state);
      expect(imgPos).toBe(dropPos);
      expect(view.state.selection.from).toBe(dropPos + 1);
      expect(view.state.doc.textContent).toBe('hello world');
    } finally {
      globalThis.Image = OriginalImage;
    }
  });

  test('does nothing for an empty file list', async () => {
    const view = makeView();
    await insertImageFiles(view, [], 4);
    expect(countImages(view.state)).toBe(0);
  });
});
