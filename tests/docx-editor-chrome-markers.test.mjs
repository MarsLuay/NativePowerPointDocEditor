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

/**
 * Minimal DOM stub: enough for stampDocxEditorChromeRegions selectors without jsdom.
 */
function createStampFixture() {
  function createElement(tagName, attrs = {}, children = []) {
    const className = String(attrs.className ?? attrs.class ?? "");
    const attributes = { ...attrs };
    delete attributes.className;
    delete attributes.class;
    if (className) {
      attributes.class = className;
    }
    const node = {
      nodeType: 1,
      tagName: tagName.toUpperCase(),
      children,
      parentElement: null,
      classList: {
        contains(token) {
          return className.split(/\s+/).includes(token);
        },
      },
      hasAttribute(name) {
        return Object.prototype.hasOwnProperty.call(attributes, name);
      },
      getAttribute(name) {
        return attributes[name] ?? null;
      },
      setAttribute(name, value) {
        attributes[name] = String(value);
      },
      get attributes() {
        return attributes;
      },
      matches(selector) {
        return matchesSelector(node, selector);
      },
      querySelector(selector) {
        return querySelectorAll(node, selector)[0] ?? null;
      },
      querySelectorAll(selector) {
        return querySelectorAll(node, selector);
      },
    };
    for (const child of children) {
      child.parentElement = node;
    }
    return node;
  }

  function classesOf(node) {
    return String(node.attributes.class ?? "").split(/\s+/).filter(Boolean);
  }

  function matchesSimple(node, simple) {
    let rest = simple.trim();
    if (!rest || rest === "*") {
      return true;
    }
    if (rest === ":scope") {
      return true;
    }
    const tagMatch = rest.match(/^([a-zA-Z][\w-]*)/);
    if (tagMatch) {
      if (node.tagName !== tagMatch[1].toUpperCase()) {
        return false;
      }
      rest = rest.slice(tagMatch[1].length);
    }
    for (const classToken of rest.matchAll(/\.([a-zA-Z_][\w-]*)/g)) {
      if (!classesOf(node).includes(classToken[1])) {
        return false;
      }
    }
    for (const attrToken of rest.matchAll(/\[([a-zA-Z_][\w-]*)(?:="([^"]*)")?\]/g)) {
      const [, name, value] = attrToken;
      if (!node.hasAttribute(name)) {
        return false;
      }
      if (value !== undefined && node.getAttribute(name) !== value) {
        return false;
      }
    }
    return true;
  }

  function matchesSelector(node, selector) {
    const parts = selector.split(",").map((part) => part.trim());
    return parts.some((part) => {
      if (part.includes(">")) {
        const [ancestorSel, childSel] = part.split(">").map((s) => s.trim());
        if (ancestorSel === ":scope") {
          return matchesSimple(node, childSel);
        }
        return false;
      }
      const compound = part.split(/\s+/).filter(Boolean);
      return matchesSimple(node, compound[compound.length - 1]);
    });
  }

  function descendants(node, includeSelf = false) {
    const out = includeSelf ? [node] : [];
    for (const child of node.children) {
      out.push(child, ...descendants(child, false));
    }
    return out;
  }

  function querySelectorAll(root, selector) {
    const parts = selector.split(",").map((part) => part.trim());
    const found = [];
    for (const part of parts) {
      if (part.includes(">")) {
        const [left, right] = part.split(">").map((s) => s.trim());
        for (const candidate of descendants(root, true)) {
          if (left === ":scope") {
            if (candidate.parentElement === root && matchesSimple(candidate, right)) {
              found.push(candidate);
            }
            continue;
          }
          if (!matchesSimple(candidate, left)) {
            continue;
          }
          for (const child of candidate.children) {
            if (matchesSimple(child, right)) {
              found.push(child);
            }
          }
        }
        continue;
      }
      const tokens = part.split(/\s+/).filter(Boolean);
      if (tokens.length === 1) {
        for (const candidate of descendants(root, false)) {
          if (matchesSimple(candidate, tokens[0])) {
            found.push(candidate);
          }
        }
        continue;
      }
      // Descendant: A B — find B under A under root
      for (const ancestor of descendants(root, true)) {
        if (!matchesSimple(ancestor, tokens[0])) {
          continue;
        }
        for (const candidate of descendants(ancestor, false)) {
          if (matchesSimple(candidate, tokens[1])) {
            found.push(candidate);
          }
        }
      }
    }
    return [...new Set(found)];
  }

  const saveItemBtn = createElement("button", { type: "button" }, []);
  saveItemBtn.textContent = "Save";
  const saveItem = createElement("div", {}, [saveItemBtn]);
  const fileDropdown = createElement("div", {}, [saveItem]);
  const fileBtn = createElement("button", { type: "button" });
  const editBtn = createElement("button", { type: "button" });
  const fileMenuRoot = createElement("div", {}, [fileBtn, fileDropdown]);
  const editMenuRoot = createElement("div", {}, [editBtn]);
  const menubar = createElement("div", { role: "menubar" }, [
    fileMenuRoot,
    editMenuRoot,
  ]);
  const titleBar = createElement("div", { "data-testid": "title-bar" }, [menubar]);
  const formattingBar = createElement("div", { "data-testid": "formatting-bar" });
  const toolbar = createElement("div", { "data-testid": "editor-toolbar" }, [formattingBar]);
  const pages = createElement("div", { className: "paged-editor__pages" });
  const scroll = createElement("div", { className: "docx-editor__scroll-container" }, [pages]);
  const root = createElement(
    "div",
    { className: "docx-editor-root docx-editor", "data-testid": "docx-editor" },
    [titleBar, toolbar, scroll],
  );
  const host = createElement("div", { className: "host" }, [root]);
  return {
    host,
    root,
    toolbar,
    titleBar,
    formattingBar,
    menubar,
    scroll,
    fileBtn,
    fileDropdown,
    saveItem,
    saveItemBtn,
  };
}

test("stampDocxEditorChromeRegions maps vendor hooks onto plugin attrs", async () => {
  const {
    DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE,
    DOCX_EDITOR_MENUBAR_ATTRIBUTE,
    DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE,
    DOCX_EDITOR_MENU_DROPDOWN_ATTRIBUTE,
    DOCX_EDITOR_MENU_ITEM_ATTRIBUTE,
    DOCX_EDITOR_MENU_ITEM_BUTTON_ATTRIBUTE,
    DOCX_EDITOR_MENU_ROOT_ATTRIBUTE,
    DOCX_EDITOR_ROOT_ATTRIBUTE,
    DOCX_EDITOR_SCROLL_CONTAINER_ATTRIBUTE,
    DOCX_EDITOR_TITLE_BAR_ATTRIBUTE,
    DOCX_EDITOR_TOOLBAR_ATTRIBUTE,
    stampDocxEditorChromeRegions,
  } = await loadDocxEditorChromeMarkersModule();

  const {
    host,
    root,
    toolbar,
    titleBar,
    formattingBar,
    menubar,
    scroll,
    fileBtn,
    fileDropdown,
    saveItem,
    saveItemBtn,
  } = createStampFixture();

  stampDocxEditorChromeRegions(host);
  stampDocxEditorChromeRegions(host);

  assert.equal(root.getAttribute(DOCX_EDITOR_ROOT_ATTRIBUTE), "true");
  assert.equal(toolbar.getAttribute(DOCX_EDITOR_TOOLBAR_ATTRIBUTE), "true");
  assert.equal(titleBar.getAttribute(DOCX_EDITOR_TITLE_BAR_ATTRIBUTE), "true");
  assert.equal(formattingBar.getAttribute(DOCX_EDITOR_FORMATTING_BAR_ATTRIBUTE), "true");
  assert.equal(menubar.getAttribute(DOCX_EDITOR_MENUBAR_ATTRIBUTE), "true");
  assert.equal(scroll.getAttribute(DOCX_EDITOR_SCROLL_CONTAINER_ATTRIBUTE), "true");
  assert.equal(fileBtn.getAttribute(DOCX_EDITOR_MENU_BUTTON_ATTRIBUTE), "true");
  assert.equal(
    menubar.children[0].getAttribute(DOCX_EDITOR_MENU_ROOT_ATTRIBUTE),
    "true",
  );
  assert.equal(fileDropdown.getAttribute(DOCX_EDITOR_MENU_DROPDOWN_ATTRIBUTE), "true");
  assert.equal(saveItem.getAttribute(DOCX_EDITOR_MENU_ITEM_ATTRIBUTE), "true");
  assert.equal(saveItemBtn.getAttribute(DOCX_EDITOR_MENU_ITEM_BUTTON_ATTRIBUTE), "true");
});
