import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadTextToolbarControllerModule } from './helpers/load-plugin-modules.mjs';

test('font-size increments accumulate and immediately update the toolbar value', async () => {
  const { TextToolbarController } = await loadTextToolbarControllerModule();
  const changes = [];
  const host = {
    currentRunStyle: { fontSizePt: 44 },
    getTextStyleContext() {
      return null;
    },
    applyRunStyle(change) {
      changes.push(change);
    },
  };
  const controller = new TextToolbarController(host);
  controller.textToolbarControls = {
    fontSizeInput: { value: '44' },
  };

  // Fire two steps without awaiting between them so coalesce can kick in, then
  // drain the apply queue. Toolbar value must reflect the final size either way.
  const first = controller.stepFontSize(1);
  const second = controller.stepFontSize(1);
  await Promise.all([first, second]);

  assert.equal(host.currentRunStyle.fontSizePt, 46);
  assert.equal(controller.textToolbarControls.fontSizeInput.value, '46');
  assert.ok(changes.length >= 1 && changes.length <= 2, `expected 1–2 applies, got ${changes.length}`);
  assert.deepEqual(changes[changes.length - 1], { fontSizePt: 46 });
});

test('font-size selection retains its latest value when the toolbar rerenders', async () => {
  const { TextToolbarController } = await loadTextToolbarControllerModule();
  globalThis.activeDocument ??= { activeElement: null };
  const changes = [];
  const ranges = [{ paragraphIndex: 0, start: 2, end: 6 }];
  const context = {
    shapeIndex: 4,
    run: { paragraphIndex: 0, runIndex: 0 },
    anchor: { left: 0, top: 0, width: 1, height: 1 },
  };
  const inactiveButton = { toggleClass() {} };
  const host = {
    currentSlide: 0,
    currentRunStyle: null,
    engine: {
      getRunStyle() {
        return { fontSizePt: 18, fontFamily: 'Arial', bold: false, italic: false, underline: false, alignment: 'l' };
      },
      areRangesStyled() {
        return false;
      },
    },
    getTextStyleContext() {
      return context;
    },
    getStoredInlineSelectionRanges() {
      return ranges;
    },
    getSelectedRangeFontSizePt() {
      return 44;
    },
    applyRunStyle(change) {
      changes.push(change);
    },
  };
  const controller = new TextToolbarController(host);
  controller.textToolbarControls = {
    bold: inactiveButton,
    italic: inactiveButton,
    underline: inactiveButton,
    fontSizeInput: { value: '' },
    textColorBar: { style: { setProperty() {} } },
    highlightBar: { style: { setProperty() {} } },
    alignButtons: { l: inactiveButton, ctr: inactiveButton, r: inactiveButton, just: inactiveButton },
  };

  controller.reflectTextToolbarState(context);
  await controller.stepFontSize(1);
  controller.reflectTextToolbarState(context);
  await controller.stepFontSize(1);

  assert.equal(controller.textToolbarControls.fontSizeInput.value, '46');
  assert.deepEqual(changes, [{ fontSizePt: 45 }, { fontSizePt: 46 }]);
});

test('rapid font-size steps coalesce to the latest size', async () => {
  const { TextToolbarController } = await loadTextToolbarControllerModule();
  const changes = [];
  let releaseApply;
  const applyGate = new Promise((resolve) => {
    releaseApply = resolve;
  });
  const host = {
    currentRunStyle: { fontSizePt: 20 },
    getTextStyleContext() {
      return null;
    },
    async applyRunStyle(change) {
      changes.push(change);
      await applyGate;
    },
  };
  const controller = new TextToolbarController(host);
  controller.textToolbarControls = {
    fontSizeInput: { value: '20' },
  };

  const p1 = controller.stepFontSize(-1);
  const p2 = controller.stepFontSize(-1);
  const p3 = controller.stepFontSize(-1);
  // First apply is in flight at 19; later steps only update pending.
  assert.deepEqual(changes, [{ fontSizePt: 19 }]);
  releaseApply();
  await Promise.all([p1, p2, p3]);

  assert.equal(controller.textToolbarControls.fontSizeInput.value, '17');
  assert.deepEqual(changes, [{ fontSizePt: 19 }, { fontSizePt: 17 }]);
});
