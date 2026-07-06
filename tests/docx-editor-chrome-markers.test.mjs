import assert from "node:assert/strict";
import test from "node:test";
import { loadDocxEditorChromeMarkersModule } from "./helpers/load-plugin-modules.mjs";

test("editor chrome markers use the selector spelling expected by deduplication", async () => {
  const {
    EDITOR_CHROME_MENU_ITEMS,
    EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE,
    markEditorChromeMenuItem,
    markEditorChromeNoToolbarTooltip,
  } = await loadDocxEditorChromeMarkersModule();
  const attributes = new Map();
  const element = {
    setAttribute(name, value) {
      attributes.set(name, value);
    },
  };

  markEditorChromeMenuItem(element, "search");
  markEditorChromeMenuItem(element, "search");
  markEditorChromeNoToolbarTooltip(element);

  assert.equal(attributes.get("data-native-powerpoint-doc-editor-search-menu-item"), "true");
  assert.equal(attributes.get(EDITOR_CHROME_NO_TOOLTIP_ATTRIBUTE), "true");
  assert.equal(attributes.has("data-native-power-point-doc-editor-search-menu-item"), false);
  assert.match(
    EDITOR_CHROME_MENU_ITEMS.search.selector,
    /\[data-native-powerpoint-doc-editor-search-menu-item\]/,
  );
  assert.match(
    EDITOR_CHROME_MENU_ITEMS.search.selector,
    /\.native-powerpoint-doc-editor-search-menu-item/,
    "legacy controls without the corrected marker remain discoverable for cleanup",
  );
  assert.equal(attributes.size, 2, "repeated synchronization remains idempotent");
});
