import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import {
  loadDocxToolbarTooltipModule,
  loadPowerPointToolbarTooltipTargetModule,
  loadTooltipControllerModule,
} from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tooltipClass = "native-powerpoint-doc-editor-toolbar-tooltip";

class FakeStyle {
  constructor() {
    this.properties = new Map();
  }

  setProperty(name, value, priority = "") {
    this.properties.set(name, { value, priority });
  }

  getPropertyValue(name) {
    return this.properties.get(name)?.value ?? "";
  }
}

class FakeClassList {
  constructor(classes = []) {
    this.classes = new Set(classes.filter(Boolean));
  }

  add(...classes) {
    for (const className of classes) this.classes.add(className);
  }

  remove(...classes) {
    for (const className of classes) this.classes.delete(className);
  }

  contains(className) {
    return this.classes.has(className);
  }

  toString() {
    return Array.from(this.classes).join(" ");
  }
}

function parseSelectorList(selector) {
  return selector.split(",").map((part) => part.trim()).filter(Boolean);
}

function splitSelectorParts(selector) {
  return selector.split(/\s+/).filter(Boolean);
}

function matchesAttribute(element, selector) {
  const match = selector.match(/^\[([^\]\^\$\*\|~=\s]+)\s*(?:(\*=|=)\s*(["']?)(.*?)\3)?\]$/);
  if (!match) return false;
  const [, name, operator, , value] = match;
  const attributeValue = element.getAttribute(name);
  if (operator === undefined) return attributeValue !== null;
  if (operator === "*=") return Boolean(attributeValue?.includes(value));
  return attributeValue === value;
}

function matchesSimpleSelector(element, selector) {
  if (!selector) return false;
  if (selector.includes(":not([role])")) {
    return matchesSimpleSelector(element, selector.replace(":not([role])", "")) && element.getAttribute("role") === null;
  }
  const attributeStart = selector.indexOf("[");
  if (attributeStart > 0) {
    return matchesSimpleSelector(element, selector.slice(0, attributeStart))
      && matchesAttribute(element, selector.slice(attributeStart));
  }
  if (selector.startsWith("[")) return matchesAttribute(element, selector);
  if (selector.startsWith(".")) {
    return selector
      .split(".")
      .filter(Boolean)
      .every((className) => element.classList.contains(className));
  }
  return element.tagName.toLowerCase() === selector.toLowerCase();
}

function matchesSelector(element, selector) {
  return parseSelectorList(selector).some((selectorPart) => {
    const parts = splitSelectorParts(selectorPart);
    let current = element;
    const last = parts.pop();
    if (!last || !matchesSimpleSelector(current, last)) return false;
    for (let index = parts.length - 1; index >= 0; index -= 1) {
      const part = parts[index];
      current = current.parentElement;
      while (current && !matchesSimpleSelector(current, part)) {
        current = current.parentElement;
      }
      if (!current) return false;
    }
    return true;
  });
}

class FakeElement {
  constructor(tagName = "div", options = {}) {
    this.tagName = tagName.toUpperCase();
    this.attributes = new Map();
    this.children = [];
    this.dataset = { ...(options.dataset ?? {}) };
    this.ownerDocument = options.ownerDocument ?? null;
    this.parentElement = null;
    this.style = new FakeStyle();
    this.textContent = options.textContent ?? "";
    this.listeners = new Map();
    this.connected = options.connected ?? true;
    this.rect = options.rect ?? { left: 40, right: 80, top: 20, bottom: 40, width: 40, height: 20 };
    this.classList = new FakeClassList(options.classes ?? []);
    for (const [name, value] of Object.entries(options.attrs ?? {})) {
      this.setAttribute(name, value);
    }
  }

  get className() {
    return this.classList.toString();
  }

  set className(value) {
    this.classList = new FakeClassList(String(value).split(/\s+/));
  }

  get isConnected() {
    return this.connected;
  }

  appendChild(child) {
    child.parentElement = this;
    child.ownerDocument = this.ownerDocument;
    child.connected = true;
    this.children.push(child);
    return child;
  }

  prepend(child) {
    child.parentElement = this;
    child.ownerDocument = this.ownerDocument;
    child.connected = true;
    this.children.unshift(child);
    return child;
  }

  remove() {
    if (this.parentElement) {
      this.parentElement.children = this.parentElement.children.filter((child) => child !== this);
    }
    this.connected = false;
    this.parentElement = null;
  }

  contains(target) {
    if (target === this) return true;
    return this.children.some((child) => child.contains(target));
  }

  closest(selector) {
    let current = this;
    while (current) {
      if (current.matches(selector)) return current;
      current = current.parentElement;
    }
    return null;
  }

  matches(selector) {
    return matchesSelector(this, selector);
  }

  querySelectorAll(selector) {
    const matches = [];
    const visit = (element) => {
      for (const child of element.children) {
        if (child.matches(selector)) matches.push(child);
        visit(child);
      }
    };
    visit(this);
    return matches;
  }

  getAttribute(name) {
    if (name === "class") return this.className || null;
    return this.attributes.has(name) ? this.attributes.get(name) : null;
  }

  setAttribute(name, value) {
    const stringValue = String(value);
    this.attributes.set(name, stringValue);
    if (name === "class") this.className = stringValue;
    if (name.startsWith("data-")) {
      const datasetKey = name
        .slice(5)
        .replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
      this.dataset[datasetKey] = stringValue;
    }
  }

  removeAttribute(name) {
    this.attributes.delete(name);
    if (name.startsWith("data-")) {
      const datasetKey = name
        .slice(5)
        .replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
      delete this.dataset[datasetKey];
    }
  }

  addEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    this.listeners.set(type, listeners.filter((entry) => entry !== listener));
  }

  dispatch(type, event = {}) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener({ target: this, relatedTarget: null, ...event });
    }
  }

  getBoundingClientRect() {
    return this.rect;
  }
}

class FakeWindow {
  constructor() {
    this.innerWidth = 320;
    this.listeners = new Map();
    this.timers = new Map();
    this.nextTimerId = 1;
    this.setTimeoutCount = 0;
  }

  setTimeout(callback) {
    const id = this.nextTimerId;
    this.nextTimerId += 1;
    this.setTimeoutCount += 1;
    this.timers.set(id, callback);
    return id;
  }

  clearTimeout(id) {
    this.timers.delete(id);
  }

  flushTimers() {
    const callbacks = Array.from(this.timers.values());
    this.timers.clear();
    for (const callback of callbacks) callback();
  }

  addEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    this.listeners.set(type, listeners.filter((entry) => entry !== listener));
  }

  dispatch(type) {
    for (const listener of this.listeners.get(type) ?? []) listener({ target: this });
  }
}

class FakeDocument {
  constructor() {
    this.nodeType = 9;
    this.defaultView = new FakeWindow();
    this.listeners = new Map();
    this.body = new FakeElement("body", { ownerDocument: this });
  }

  createElement(tagName) {
    return new FakeElement(tagName, {
      ownerDocument: this,
      rect: this.nextElementRect ?? { left: 80, right: 180, top: 48, bottom: 68, width: 100, height: 20 },
    });
  }

  querySelectorAll(selector) {
    return this.body.querySelectorAll(selector);
  }

  contains(target) {
    return this.body.contains(target);
  }

  addEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    listeners.push(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type, listener) {
    const listeners = this.listeners.get(type) ?? [];
    this.listeners.set(type, listeners.filter((entry) => entry !== listener));
  }

  dispatch(type, event = {}) {
    for (const listener of this.listeners.get(type) ?? []) {
      listener({ target: this, relatedTarget: null, ...event });
    }
  }
}

function append(parent, tagName = "div", options = {}) {
  const child = new FakeElement(tagName, { ...options, ownerDocument: parent.ownerDocument });
  parent.appendChild(child);
  return child;
}

function createControllerFixture() {
  const ownerDocument = new FakeDocument();
  const root = append(ownerDocument.body, "div", { classes: ["root"] });
  const toolbar = append(root, "div", { classes: ["native-powerpoint-toolbar"] });
  const button = append(toolbar, "button", {
    attrs: { "aria-label": "Bold", title: "Native bold" },
    dataset: { tooltip: "Bold" },
  });
  const icon = append(button, "span", { classes: ["icon"] });
  const otherButton = append(toolbar, "button", {
    attrs: { "aria-label": "Italic", title: "Native italic" },
    dataset: { tooltip: "Italic" },
    rect: { left: 120, right: 160, top: 20, bottom: 40, width: 40, height: 20 },
  });
  return { ownerDocument, root, toolbar, button, icon, otherButton, view: ownerDocument.defaultView };
}

test("tooltip text contract uses data-tooltip, aria-label, then title", async () => {
  const { getToolbarTooltipText } = await loadDocxToolbarTooltipModule();
  const button = new FakeElement("button", {
    attrs: { "aria-label": "Accessible label", title: "Native title" },
    dataset: { tooltip: "Visible label" },
  });

  assert.equal(getToolbarTooltipText(button), "Visible label");
  delete button.dataset.tooltip;
  button.removeAttribute("data-tooltip");
  assert.equal(getToolbarTooltipText(button), "Accessible label");
  button.removeAttribute("aria-label");
  assert.equal(getToolbarTooltipText(button), "Native title");
});

test("native title is suspended and restored while still resolving tooltip text", async () => {
  const { getToolbarTooltipText, restoreNativeTitle, suspendNativeTitle } = await loadDocxToolbarTooltipModule();
  const button = new FakeElement("button", { attrs: { title: "Insert link (Ctrl+K)" } });

  suspendNativeTitle(button);

  assert.equal(button.getAttribute("title"), null);
  assert.equal(button.dataset.nativePowerPointDocEditorTooltipTitle, "Insert link (Ctrl+K)");
  assert.equal(getToolbarTooltipText(button), "Insert link (Ctrl+K)");

  restoreNativeTitle(button);

  assert.equal(button.getAttribute("title"), "Insert link (Ctrl+K)");
  assert.equal(button.dataset.nativePowerPointDocEditorTooltipTitle, undefined);
});

test("controller keeps one tooltip and switches targets with title restoration", async () => {
  const { TooltipController } = await loadTooltipControllerModule();
  const { ownerDocument, root, button, otherButton, view } = createControllerFixture();
  const controller = new TooltipController({
    root,
    ownerDocument,
    view,
    delayMs: 1,
    getTarget: (target) => target.closest("button"),
  });

  root.dispatch("pointerover", { target: button });
  assert.equal(button.getAttribute("title"), null);
  view.flushTimers();
  assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 1);
  assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`)[0].textContent, "Bold");

  root.dispatch("pointerover", { target: otherButton });
  assert.equal(button.getAttribute("title"), "Native bold");
  assert.equal(otherButton.getAttribute("title"), null);
  view.flushTimers();

  const tooltips = ownerDocument.body.querySelectorAll(`.${tooltipClass}`);
  assert.equal(tooltips.length, 1);
  assert.equal(tooltips[0].textContent, "Italic");

  root.dispatch("pointerout", { target: otherButton, relatedTarget: root });
  assert.equal(otherButton.getAttribute("title"), "Native italic");
  assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 0);
  controller.detach();
});

test("moving between an icon and parent button does not restart the timer", async () => {
  const { TooltipController } = await loadTooltipControllerModule();
  const { ownerDocument, root, button, icon, view } = createControllerFixture();
  const controller = new TooltipController({
    root,
    ownerDocument,
    view,
    delayMs: 1,
    getTarget: (target) => target.closest("button"),
  });

  root.dispatch("pointerover", { target: button });
  root.dispatch("pointerover", { target: icon });

  assert.equal(view.setTimeoutCount, 1);
  controller.detach();
});

test("scroll, resize, blur, and detach remove active tooltip", async () => {
  const { TooltipController } = await loadTooltipControllerModule();

  for (const eventName of ["scroll", "resize", "blur", "detach"]) {
    const { ownerDocument, root, button, view } = createControllerFixture();
    const controller = new TooltipController({
      root,
      ownerDocument,
      view,
      delayMs: 1,
      getTarget: (target) => target.closest("button"),
    });
    root.dispatch("pointerover", { target: button });
    view.flushTimers();
    assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 1);

    if (eventName === "scroll") root.dispatch("scroll", { target: root });
    else if (eventName === "resize") view.dispatch("resize");
    else if (eventName === "blur") view.dispatch("blur");
    else controller.detach();

    assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 0);
    assert.equal(button.getAttribute("title"), "Native bold");
    controller.detach();
  }
});

test("positioning clamps left and right viewport edges", async () => {
  const { positionToolbarTooltip } = await loadDocxToolbarTooltipModule();
  const target = new FakeElement("button", {
    rect: { left: 0, right: 10, top: 0, bottom: 20, width: 10, height: 20 },
  });
  const tooltip = new FakeElement("div", {
    rect: { left: -20, right: 40, top: 28, bottom: 48, width: 60, height: 20 },
  });

  positionToolbarTooltip(target, tooltip, { viewportWidth: 300 });
  assert.equal(tooltip.style.getPropertyValue("--native-powerpoint-doc-editor-toolbar-tooltip-left"), "8px");
  assert.equal(tooltip.classList.contains("is-left-aligned"), true);

  tooltip.rect = { left: 260, right: 330, top: 28, bottom: 48, width: 70, height: 20 };
  positionToolbarTooltip(target, tooltip, { viewportWidth: 300 });
  assert.equal(tooltip.style.getPropertyValue("--native-powerpoint-doc-editor-toolbar-tooltip-left"), "292px");
  assert.equal(tooltip.classList.contains("is-right-aligned"), true);
});

test("DOCX resolver matches toolbar controls but not links, menus, or popovers", async () => {
	const { getDocumentLinkTitleTarget, getToolbarTooltipTarget } = await loadDocxToolbarTooltipModule();
	const ownerDocument = new FakeDocument();
	const editorRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-1"] });
	const toolbar = append(editorRoot, "div", { attrs: { "data-testid": "editor-toolbar" } });
	const button = append(toolbar, "button", { attrs: { "aria-label": "Bold" } });
	const icon = append(button, "span");
	const epRoot = append(editorRoot, "div", { classes: ["ep-root"] });
	const formattingBar = append(epRoot, "div", { attrs: { "data-testid": "formatting-bar" } });
	const formattingButton = append(formattingBar, "button", { attrs: { "aria-label": "Align Left" } });
	const otherRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-2"] });
	const otherToolbar = append(otherRoot, "div", { attrs: { "data-testid": "editor-toolbar" } });
	const otherButton = append(otherToolbar, "button", { attrs: { "aria-label": "Other save" } });

	assert.equal(getToolbarTooltipTarget(icon, editorRoot), button);
	assert.equal(getToolbarTooltipTarget(formattingButton, editorRoot), null);
	assert.equal(getToolbarTooltipTarget(otherButton, editorRoot), null);

  const page = append(editorRoot, "div", { classes: ["layout-page"] });
  const link = append(page, "a", { attrs: { title: "https://example.com" } });
  assert.equal(getToolbarTooltipTarget(link, editorRoot), null);
  assert.equal(getDocumentLinkTitleTarget(link, editorRoot), link);

  const menu = append(toolbar, "div", { attrs: { role: "menu" } });
  const menuButton = append(menu, "button", { attrs: { "aria-label": "Menu item" } });
  assert.equal(getToolbarTooltipTarget(menuButton, editorRoot), null);

  const popover = append(toolbar, "div", { classes: ["ep-hyperlink-popup"] });
	const popoverButton = append(popover, "button", { attrs: { "aria-label": "Edit link" } });
	assert.equal(getToolbarTooltipTarget(popoverButton, editorRoot), null);
});

test("DOCX tooltip metadata sync leaves formatting labels alone", async () => {
	const { neutralizeToolbarButtonTooltipSources } = await loadDocxToolbarTooltipModule();
	const ownerDocument = new FakeDocument();
	const editorRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-1"] });
	const epRoot = append(editorRoot, "div", { classes: ["ep-root"] });
	const formattingBar = append(epRoot, "div", {
		attrs: {
			"data-testid": "formatting-bar",
			role: "toolbar",
			"aria-label": "Formatting toolbar",
			title: "Formatting toolbar",
			"data-tooltip": "Formatting toolbar",
		},
	});
	const group = append(formattingBar, "div", { attrs: { role: "group", "aria-label": "Alignment" } });
	const button = append(group, "button", { attrs: { "aria-label": "Align Left" } });
	const editorToolbar = append(epRoot, "div", { attrs: { "data-testid": "editor-toolbar" } });
	const editorButton = append(editorToolbar, "button", { attrs: { "aria-label": "Save" } });

	neutralizeToolbarButtonTooltipSources(editorRoot);

	assert.equal(formattingBar.getAttribute("aria-label"), "Formatting toolbar");
	assert.equal(formattingBar.getAttribute("title"), "Formatting toolbar");
	assert.equal(formattingBar.getAttribute("data-tooltip"), "Formatting toolbar");
	assert.equal(group.getAttribute("aria-label"), "Alignment");
	assert.equal(group.dataset.nativePowerPointDocEditorToolbarGroupLabel, undefined);
	assert.equal(button.getAttribute("aria-label"), "Align Left");
	assert.equal(button.dataset.tooltip, undefined);
	assert.equal(editorButton.dataset.tooltip, "Save");
});

test("DOCX suppression hides inline Eigenpal toolbar tooltips without touching body portals", async () => {
	const { suppressEigenpalToolbarTooltips } = await loadDocxToolbarTooltipModule();
	const ownerDocument = new FakeDocument();
	const editorRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-1"] });
	const epRoot = append(editorRoot, "div", { classes: ["ep-root"] });
	const inlineTooltip = append(epRoot, "div", {
		classes: ["fixed", "z-50", "px-2", "py-1", "rounded-md", "shadow-lg"],
		textContent: "Align Left",
	});
	const portalRoot = append(ownerDocument.body, "div", { classes: ["ep-root"] });
	const portaledTooltip = append(portalRoot, "div", {
		classes: ["fixed", "z-50", "px-2", "py-1", "rounded-md", "shadow-lg"],
		textContent: "Alignment",
	});
	const linkPopup = append(portalRoot, "div", { classes: ["ep-hyperlink-popup"] });
	const linkPreview = append(linkPopup, "div", {
		classes: ["fixed", "z-50", "px-2", "py-1", "rounded-md", "shadow-lg"],
		textContent: "https://example.com",
	});

	suppressEigenpalToolbarTooltips(editorRoot);

	assert.equal(inlineTooltip.hidden, true);
	assert.equal(inlineTooltip.getAttribute("data-native-powerpoint-doc-editor-eigenpal-tooltip"), "true");
	assert.equal(portaledTooltip.hidden, undefined);
	assert.equal(portaledTooltip.getAttribute("data-native-powerpoint-doc-editor-eigenpal-tooltip"), null);
	assert.equal(linkPreview.hidden, undefined);
	assert.equal(linkPreview.getAttribute("data-native-powerpoint-doc-editor-eigenpal-tooltip"), null);
});

test("DOCX toolbar tooltip manager is scoped and singleton per editor root", async () => {
	const { attachDocxToolbarTooltipManager } = await loadDocxToolbarTooltipModule();
	const ownerDocument = new FakeDocument();
	const editorRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-1"] });
	const toolbar = append(editorRoot, "div", { attrs: { "data-testid": "editor-toolbar" } });
	const button = append(toolbar, "button", {
		attrs: { "aria-label": "Save", title: "Save" },
		dataset: { tooltip: "Save" },
	});
	const icon = append(button, "span");
	const otherRoot = append(ownerDocument.body, "div", { classes: ["native-powerpoint-doc-editor-editor-2"] });
	const otherToolbar = append(otherRoot, "div", { attrs: { "data-testid": "editor-toolbar" } });
	const otherButton = append(otherToolbar, "button", {
		attrs: { "aria-label": "Other save", title: "Other save" },
		dataset: { tooltip: "Other save" },
	});
	const originalMutationObserver = globalThis.MutationObserver;

	globalThis.MutationObserver = class {
		observe() {}
		disconnect() {}
	};

	try {
		const detachFirst = attachDocxToolbarTooltipManager(editorRoot);
		const rootPointerOverCount = editorRoot.listeners.get("pointerover")?.length ?? 0;
		const detachSecond = attachDocxToolbarTooltipManager(editorRoot);

		assert.equal(editorRoot.listeners.get("pointerover")?.length ?? 0, rootPointerOverCount);
		assert.equal(ownerDocument.listeners.get("pointerover")?.length ?? 0, 0);
		assert.equal(otherRoot.listeners.get("pointerover")?.length ?? 0, 0);

		editorRoot.dispatch("pointerover", { target: icon });
		ownerDocument.defaultView.flushTimers();
		assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 1);
		assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`)[0].textContent, "Save");

		otherRoot.dispatch("pointerover", { target: otherButton });
		ownerDocument.defaultView.flushTimers();
		assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 1);
		assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`)[0].textContent, "Save");

		detachFirst();
		assert.equal(editorRoot.listeners.get("pointerover")?.length ?? 0, rootPointerOverCount);
		detachSecond();
		assert.equal(editorRoot.listeners.get("pointerover")?.length ?? 0, 0);
		assert.equal(ownerDocument.body.querySelectorAll(`.${tooltipClass}`).length, 0);
	} finally {
		if (originalMutationObserver === undefined) {
			delete globalThis.MutationObserver;
		} else {
			globalThis.MutationObserver = originalMutationObserver;
		}
	}
});

test("PowerPoint resolver matches toolbar controls but not popovers", async () => {
	const { resolvePowerPointTooltipTarget } = await loadPowerPointToolbarTooltipTargetModule();
	const ownerDocument = new FakeDocument();
	const root = append(ownerDocument.body, "div", { classes: ["native-powerpoint-root"] });
  const toolbar = append(root, "div", { classes: ["native-powerpoint-toolbar"] });
  const button = append(toolbar, "button", { attrs: { "aria-label": "Undo" } });
  const icon = append(button, "span");
  const textToolbar = append(root, "div", { classes: ["native-powerpoint-text-toolbar"] });
  const textButton = append(textToolbar, "button", { attrs: { "aria-label": "Bold" } });
  const findPanel = append(root, "div", { classes: ["native-powerpoint-find-panel"] });
  const findButton = append(findPanel, "button", { attrs: { "aria-label": "Next match" } });
  const rotateHandle = append(root, "div", { classes: ["native-powerpoint-rotate-handle"] });
  const popover = append(root, "div", { classes: ["native-powerpoint-toolbar-popover"] });
  const popoverButton = append(popover, "button", { attrs: { "aria-label": "Red" } });

  assert.equal(resolvePowerPointTooltipTarget(icon, root), button);
  assert.equal(resolvePowerPointTooltipTarget(textButton, root), textButton);
  assert.equal(resolvePowerPointTooltipTarget(findButton, root), findButton);
  assert.equal(resolvePowerPointTooltipTarget(rotateHandle, root), rotateHandle);
  assert.equal(resolvePowerPointTooltipTarget(popoverButton, root), null);
});

test("tooltip CSS keeps the same dark palette and semantic Eigenpal marker", async () => {
  const css = await readFile(path.join(projectRoot, "styles.css"), "utf8");
  const backgroundValues = [...css.matchAll(/--npde-docx-toolbar-tooltip-bg:\s*([^;]+);/g)].map((match) => match[1].trim());
  const textValues = [...css.matchAll(/--npde-docx-toolbar-tooltip-text:\s*([^;]+);/g)].map((match) => match[1].trim());

  assert.deepEqual(new Set(backgroundValues), new Set(["#0f172a"]));
  assert.deepEqual(new Set(textValues), new Set(["#f8fafc"]));
  assert.match(css, /\[data-native-powerpoint-doc-editor-eigenpal-tooltip='true'\]/);
  assert.doesNotMatch(css, /\.fixed\.z-50\.px-2\.py-1\.rounded-md\.shadow-lg:not\(\[role\]\)/);
  assert.doesNotMatch(css, /\.native-powerpoint-doc-editor-sr-only-label/);
  assert.match(css, new RegExp(`\\.${tooltipClass}\\s*\\{`));
});
