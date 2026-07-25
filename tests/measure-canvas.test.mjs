import assert from 'node:assert/strict';
import test from 'node:test';

import { createDetachedMeasureCanvas } from '../src/powerpoint/measureCanvas.ts';

test('createDetachedMeasureCanvas prefers Window.createEl over Document.createEl', () => {
  const created = { tagName: 'CANVAS' };
  let windowCreateElCalls = 0;
  let documentCreateElCalls = 0;

  const fakeWindow = {
    createEl(tag) {
      windowCreateElCalls += 1;
      assert.equal(tag, 'canvas');
      return created;
    },
  };

  const fakeDocument = {
    defaultView: fakeWindow,
    createEl() {
      documentCreateElCalls += 1;
      throw new Error('Only one element on document allowed.');
    },
    createElement() {
      throw new Error('createElement should not be used when createEl exists');
    },
  };

  const canvas = createDetachedMeasureCanvas(fakeDocument);
  assert.equal(canvas, created);
  assert.equal(windowCreateElCalls, 1);
  assert.equal(documentCreateElCalls, 0);
});

test('createDetachedMeasureCanvas falls back to doc.win.createEl when Window.createEl missing', () => {
  const created = { tagName: 'CANVAS' };
  const fakeWindow = {
    createEl(tag) {
      assert.equal(tag, 'canvas');
      return created;
    },
  };
  const fakeDocument = {
    defaultView: {},
    win: fakeWindow,
  };

  const canvas = createDetachedMeasureCanvas(fakeDocument);
  assert.equal(canvas, created);
});

test('createDetachedMeasureCanvas returns null when win.createEl throws HierarchyRequestError', () => {
  const fakeDocument = {
    defaultView: {},
    win: {
      createEl() {
        const error = new Error('Failed to execute \'appendChild\' on \'Node\': Only one element on document allowed.');
        error.name = 'HierarchyRequestError';
        throw error;
      },
    },
  };

  assert.equal(createDetachedMeasureCanvas(fakeDocument), null);
});

test('createDetachedMeasureCanvas returns null when no createEl host exists', () => {
  assert.equal(createDetachedMeasureCanvas({ defaultView: {} }), null);
});
