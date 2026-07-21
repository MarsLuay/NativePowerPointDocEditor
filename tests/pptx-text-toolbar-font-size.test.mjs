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

  controller.stepFontSize(1);
  controller.stepFontSize(1);

  assert.deepEqual(changes, [{ fontSizePt: 45 }, { fontSizePt: 46 }]);
  assert.equal(host.currentRunStyle.fontSizePt, 46);
  assert.equal(controller.textToolbarControls.fontSizeInput.value, '46');
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
  controller.stepFontSize(1);
  controller.reflectTextToolbarState(context);
  controller.stepFontSize(1);

  assert.deepEqual(changes, [{ fontSizePt: 45 }, { fontSizePt: 46 }]);
  assert.equal(controller.textToolbarControls.fontSizeInput.value, '46');
});
