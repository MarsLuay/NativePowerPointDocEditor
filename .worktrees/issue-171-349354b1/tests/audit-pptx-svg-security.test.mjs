import assert from "node:assert/strict";
import { test } from "node:test";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");

// sanitizeSvg is pure but depends on the global DOMParser/XMLSerializer that the
// Obsidian/browser runtime provides. Mirror tests/helpers/load-plugin-modules.mjs.
globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

// The Obsidian/Electron browser runtime provides ChildNode.remove(), which
// sanitizeElement() relies on to drop blocked nodes. @xmldom/xmldom does not
// implement it, so polyfill the parser-backed Element prototype to match the
// production DOM. (children/isConnected/removeAttribute are already supported.)
{
  const probe = new DOMParser().parseFromString("<svg><rect/></svg>", "image/svg+xml");
  const elementProto = Object.getPrototypeOf(probe.documentElement.firstChild);
  if (typeof elementProto.remove !== "function") {
    elementProto.remove = function remove() {
      if (this.parentNode) {
        this.parentNode.removeChild(this);
      }
    };
  }
}

// Mirror bundleSource() from tests/helpers/load-plugin-modules.mjs so this file
// stays self-contained (no edits to the shared helper).
let svgSecurityModulePromise;
function loadSvgSecurityModule() {
  svgSecurityModulePromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "native-powerpoint-svg-audit-"));
    const outfile = path.join(outputDirectory, "svg-security.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/SvgSecurity.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });
    return require(outfile);
  })();
  return svgSecurityModulePromise;
}

function wrapSvg(inner) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">${inner}</svg>`;
}

test("strict mode strips <script> while keeping benign shapes", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const dirty = wrapSvg(
    '<rect x="1" y="2" width="10" height="10" fill="red"/>' +
      '<script>alert(1)</script>',
  );

  const result = sanitizeSvg(dirty, { mode: "strict" });

  assert.ok(!/<script/i.test(result.svg), "script element must be removed");
  assert.match(result.svg, /<rect/i, "benign rect must survive");
  assert.ok(
    result.issues.some((issue) => issue.detail.toLowerCase().includes("script")),
    "an issue must be recorded for the blocked script",
  );
});

test("permanently blocked elements (iframe, foreignObject) are stripped in strict mode", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const dirty = wrapSvg(
    '<iframe src="https://evil.example"></iframe>' +
      '<foreignObject><div xmlns="http://www.w3.org/1999/xhtml">x</div></foreignObject>' +
      '<circle cx="5" cy="5" r="3"/>',
  );

  const result = sanitizeSvg(dirty, { mode: "strict" });

  assert.ok(!/<iframe/i.test(result.svg), "iframe must be removed");
  assert.ok(!/<foreignObject/i.test(result.svg), "foreignObject must be removed");
  assert.match(result.svg, /<circle/i, "benign circle must survive");
});

test("on* event handlers are stripped in strict mode", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const dirty = wrapSvg('<rect x="0" y="0" width="9" height="9" onload="alert(1)" onclick="steal()"/>');

  const result = sanitizeSvg(dirty, { mode: "strict" });

  assert.ok(!/onload/i.test(result.svg), "onload handler must be removed");
  assert.ok(!/onclick/i.test(result.svg), "onclick handler must be removed");
  assert.match(result.svg, /<rect/i, "the element itself survives once handlers are stripped");
  assert.ok(
    result.issues.some((issue) => issue.reason === "Detected risky action"),
    "risky-action issue must be reported for on* handlers",
  );
});

test("javascript: URLs in href are stripped in strict mode", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const dirty = wrapSvg('<use href="javascript:alert(1)"/>');

  const result = sanitizeSvg(dirty, { mode: "strict" });

  assert.ok(!/javascript:/i.test(result.svg), "scriptable protocol must be removed");
});

test("dangerous elements/handlers are ALSO blocked in compatibility mode", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const dirty = wrapSvg(
    '<script>alert(1)</script>' +
      '<iframe src="https://evil.example"></iframe>' +
      '<rect x="0" y="0" width="5" height="5" onload="alert(1)"/>',
  );

  const result = sanitizeSvg(dirty, { mode: "compatibility" });

  assert.ok(!/<script/i.test(result.svg), "script must be blocked even in compatibility mode");
  assert.ok(!/<iframe/i.test(result.svg), "iframe must be blocked even in compatibility mode");
  assert.ok(!/onload/i.test(result.svg), "on* handlers must be blocked even in compatibility mode");
});

test("summarizeSvgSecurityIssues aggregates issues by reason", async () => {
  const { sanitizeSvg, summarizeSvgSecurityIssues } = await loadSvgSecurityModule();
  const dirty = wrapSvg(
    '<script>a()</script>' +
      '<rect onload="x()" onclick="y()" x="0" y="0" width="4" height="4"/>',
  );

  const result = sanitizeSvg(dirty, { mode: "strict" });
  const summary = summarizeSvgSecurityIssues(result.issues);

  assert.ok(result.issues.length > 0, "issues must be recorded");
  assert.ok(summary.length > 0, "summary must contain at least one reason");
  for (const line of summary) {
    assert.match(line, /: \d+$/, "each summary line ends with a count");
  }
  const riskyAction = summary.find((line) => line.startsWith("Detected risky action:"));
  assert.ok(riskyAction, "summary must aggregate the two on* handlers");
  assert.equal(riskyAction, "Detected risky action: 2", "both on* handlers counted under one reason");
});

test("benign full SVG passes strict sanitize untouched of issues", async () => {
  const { sanitizeSvg } = await loadSvgSecurityModule();
  const clean = wrapSvg(
    '<g><rect x="1" y="1" width="8" height="8" fill="#0a0"/>' +
      '<circle cx="5" cy="5" r="2" stroke="black"/>' +
      '<text x="2" y="2" font-size="3">hi</text></g>',
  );

  const result = sanitizeSvg(clean, { mode: "strict" });

  assert.equal(result.issues.length, 0, "no issues for benign content");
  assert.match(result.svg, /<rect/i);
  assert.match(result.svg, /<circle/i);
  assert.match(result.svg, /<text/i);
});
