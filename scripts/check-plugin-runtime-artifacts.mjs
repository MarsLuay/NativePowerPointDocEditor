import assert from 'node:assert/strict';
import { stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const syncStandardLimitBytes = 5_000_000;
const runtimeArtifacts = [
  'main.js',
  'pptx-js-engine.mjs',
  'pptx-wasm-renderer.mjs',
  'heic-decode.mjs',
];

const artifactPaths = runtimeArtifacts.map((artifact) => path.join(projectRoot, artifact));
const sizes = await Promise.all(artifactPaths.map(async (artifactPath) => ({
  artifactPath,
  size: (await stat(artifactPath)).size,
})));

for (const { artifactPath, size } of sizes) {
  assert.ok(
    size <= syncStandardLimitBytes,
    `${artifactPath} is ${size} bytes, above Obsidian Sync Standard's 5 MB per-file limit.`,
  );
}

const [jsFallback, wasmRenderer, heicDecoder] = await Promise.all([
  import(pathToFileURL(path.join(projectRoot, 'pptx-js-engine.mjs')).href),
  import(pathToFileURL(path.join(projectRoot, 'pptx-wasm-renderer.mjs')).href),
  import(pathToFileURL(path.join(projectRoot, 'heic-decode.mjs')).href),
]);

assert.equal(typeof jsFallback.createPptxJsEngine, 'function', 'JS fallback artifact must export createPptxJsEngine().');
assert.equal(typeof wasmRenderer.PptxRenderer, 'function', 'WASM renderer artifact must export PptxRenderer.');
assert.ok(wasmRenderer.wasmBytes instanceof Uint8Array, 'WASM renderer artifact must export Uint8Array wasmBytes.');
assert.equal(typeof heicDecoder.default, 'function', 'HEIC decoder artifact must have a default decoder export.');

console.log(
  `[check:plugin-runtime-artifacts] Sync-safe runtime artifacts: ${sizes
    .map(({ artifactPath, size }) => `${path.basename(artifactPath)}=${size}`)
    .join(', ')}`,
);
