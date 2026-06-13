// Minimal DOM query polyfills for @xmldom/xmldom-parsed SVG trees. The stamping
// module expects querySelector/closest; xmldom only provides getElementsByTagName.

import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { DOMParser } = require("@xmldom/xmldom");

function tagName(el) {
  return (el.tagName || el.nodeName || "").toLowerCase().replace(/^.*:/, "");
}

function walkElements(root, visit) {
  if (!root || root.nodeType !== 1) return;
  visit(root);
  for (const child of Array.from(root.childNodes || [])) {
    walkElements(child, visit);
  }
}

function matchesSelector(el, selector) {
  const m = selector.match(/^([a-z*]+|\*)\[([^\]=]+)(?:="([^"]*)")?\]$/i);
  if (!m) return false;
  const wantedTag = m[1].toLowerCase();
  const attr = m[2];
  const value = m[3];
  if (wantedTag !== "*" && tagName(el) !== wantedTag) return false;
  const actual = el.getAttribute?.(attr);
  if (actual == null) return false;
  return value === undefined || actual === value;
}

function querySelectorAll(root, selector) {
  const results = [];
  walkElements(root, (el) => {
    if (matchesSelector(el, selector)) results.push(el);
  });
  return results;
}

function querySelector(root, selector) {
  return querySelectorAll(root, selector)[0] ?? null;
}

function closest(el, selector) {
  let node = el;
  while (node && node.nodeType === 1) {
    if (matchesSelector(node, selector)) return node;
    node = node.parentNode;
  }
  return null;
}

function augmentDomQueries(root) {
  walkElements(root, (el) => {
    if (typeof el.querySelector !== "function") {
      el.querySelector = (selector) => querySelector(el, selector);
      el.querySelectorAll = (selector) => querySelectorAll(el, selector);
      el.closest = (selector) => closest(el, selector);
    }
  });
}

/** Parse rendered slide SVG into a queryable document element tree. */
export function parseSvgForStamping(svgString) {
  const doc = new DOMParser().parseFromString(svgString, "image/svg+xml");
  const svg = doc.documentElement;
  augmentDomQueries(svg);
  if (!svg.querySelector) {
    svg.querySelector = (selector) => querySelector(svg, selector);
    svg.querySelectorAll = (selector) => querySelectorAll(svg, selector);
    svg.closest = (selector) => closest(svg, selector);
  }
  return svg;
}
