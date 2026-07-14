// Audit test for the "PPTX · Runtime & JS fallback" feature group.
//
// Covers two layers:
//   1. The pure runtime-compat helpers (src/powerpoint/runtimeCompat.ts):
//      isWasmGcUnsupportedError / cleanError. These must recognise the exact
//      failure string pptx-svg throws ("Wasm init failed — requires
//      WebAssembly GC support …") so the engine and view both fall back/branch
//      correctly.
//   2. The forced-JS path end-to-end: setForceJsBackendOverride(true) makes
//      PresentationEngine skip the Wasm backend and render through the
//      pure-JS engine — the same path mobile WebViews use. Mirrors
//      scripts/smoke-mobile-pptx.mjs.

import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { build } from "esbuild";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

// Mirror of the (un-exported) bundleSource pattern in
// tests/helpers/load-plugin-modules.mjs: compile a TS source file to a CJS
// bundle in a temp dir, then require it.
async function bundleRuntimeCompat() {
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "native-powerpoint-runtime-audit-"));
  const outfile = path.join(outputDirectory, "runtime-compat.cjs");
  await build({
    entryPoints: [path.join(projectRoot, "src/powerpoint/runtimeCompat.ts")],
    bundle: true,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  return require(outfile);
}

test("isWasmGcUnsupportedError matches the real pptx-svg failure string", async () => {
  const { isWasmGcUnsupportedError } = await bundleRuntimeCompat();

  // The exact message pptx-svg's wasm-compat.js throws when every
  // instantiation tier fails (node_modules/pptx-svg/dist/wasm-compat.js:155).
  const realMessage =
    "Wasm init failed — requires WebAssembly GC support (Chrome 111+, Node.js 22+). " +
    "Tier-3 error: CompileError: WebAssembly.instantiate(): ...";

  assert.equal(isWasmGcUnsupportedError(new Error(realMessage)), true, "real Error message");
  assert.equal(isWasmGcUnsupportedError(realMessage), true, "real message as string");
  assert.equal(isWasmGcUnsupportedError(new Error("Wasm init failed")), true, "prefix only");
  assert.equal(
    isWasmGcUnsupportedError(new Error("This runtime requires WebAssembly GC support")),
    true,
    "suffix only",
  );
  // Case-insensitive.
  assert.equal(isWasmGcUnsupportedError(new Error("wasm INIT FAILED")), true, "case-insensitive");
});

test("isWasmGcUnsupportedError is false for unrelated errors", async () => {
  const { isWasmGcUnsupportedError } = await bundleRuntimeCompat();

  assert.equal(isWasmGcUnsupportedError(new Error("Could not render slide.")), false);
  assert.equal(isWasmGcUnsupportedError(new Error("ENOENT: no such file")), false);
  assert.equal(isWasmGcUnsupportedError("some unrelated string"), false);
  assert.equal(isWasmGcUnsupportedError(undefined), false);
  assert.equal(isWasmGcUnsupportedError(null), false);
  assert.equal(isWasmGcUnsupportedError(0), false);
});

test("cleanError normalizes Error, string, and empty/undefined inputs", async () => {
  const { cleanError } = await bundleRuntimeCompat();

  assert.equal(cleanError(new Error("boom")), "boom");
  assert.equal(cleanError("plain string"), "plain string");
  assert.equal(cleanError(undefined), "Unknown error");
  assert.equal(cleanError(null), "Unknown error");
  assert.equal(cleanError(""), "Unknown error");
  assert.equal(cleanError(42), "42");
});

test("forced-JS path renders a deck end-to-end via PresentationEngine", async () => {
  const fixturePath = path.join(projectRoot, "tests/fixtures/decks/features.pptx");
  assert.ok(existsSync(fixturePath), `missing fixture: ${fixturePath}`);

  const { PresentationEngine, setForceJsBackendOverride, resetForceJsBackendOverride } =
    await loadPresentationEngineModule();
  setForceJsBackendOverride(true);
  try {
    const fileBuffer = await readFile(fixturePath);
    const buffer = fileBuffer.buffer.slice(
      fileBuffer.byteOffset,
      fileBuffer.byteOffset + fileBuffer.byteLength,
    );
    const engine = await PresentationEngine.load(buffer);

    assert.equal(engine.getRendererBackend(), "js");
    assert.ok(engine.slideCount > 0, `expected slideCount > 0, got ${engine.slideCount}`);

    const { svg } = engine.renderSlide(0);
    assert.equal(typeof svg, "string");
    assert.ok(svg.includes("<svg"), "renderSlide(0) did not return SVG markup");
    assert.ok(!svg.startsWith("ERROR:"), `renderSlide returned ${svg.slice(0, 80)}`);
  } finally {
    resetForceJsBackendOverride();
  }
});
