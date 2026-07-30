import assert from 'node:assert/strict';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { OBSIDIAN_SUPPORTED_RELEASE_ASSETS } from './lib/pptx-runtime-artifact-spec.mjs';

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

const mainBundle = await readFile(path.join(projectRoot, 'main.js'), 'utf8');
for (const artifact of runtimeArtifacts.slice(1)) {
  const escapedArtifact = artifact.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  assert.doesNotMatch(
    mainBundle,
    new RegExp(`import\\(["']\\./${escapedArtifact}["']\\)`),
    `main.js must load ${artifact} through Obsidian's plugin resource URL, not a relative dynamic import.`,
  );
  assert.match(
    mainBundle,
    new RegExp(escapedArtifact),
    `main.js must embed ${artifact} so community installs can materialize it without unsupported release assets.`,
  );
}

assert.match(
  mainBundle,
  /ensurePptxRuntimeArtifacts|Materialized optional PowerPoint runtime artifact/,
  'main.js must include the runtime artifact materializer for community installs.',
);

assert.equal(typeof jsFallback.createPptxJsEngine, 'function', 'JS fallback artifact must export createPptxJsEngine().');
assert.equal(typeof wasmRenderer.PptxRenderer, 'function', 'WASM renderer artifact must export PptxRenderer.');
assert.ok(wasmRenderer.wasmBytes instanceof Uint8Array, 'WASM renderer artifact must export Uint8Array wasmBytes.');
assert.equal(typeof heicDecoder.default, 'function', 'HEIC decoder artifact must have a default decoder export.');

const releaseWorkflow = await readFile(path.join(projectRoot, '.github/workflows/release.yml'), 'utf8');
for (const artifact of runtimeArtifacts.slice(1)) {
  assert.doesNotMatch(
    releaseWorkflow,
    new RegExp(`^\\s*${artifact.replace(/\./g, '\\.')}\\s*$`, 'm'),
    `release.yml must not upload unsupported Obsidian asset ${artifact}; community installs only download ${OBSIDIAN_SUPPORTED_RELEASE_ASSETS.join(', ')}.`,
  );
}
for (const asset of OBSIDIAN_SUPPORTED_RELEASE_ASSETS) {
  assert.match(
    releaseWorkflow,
    new RegExp(`^\\s*${asset.replace(/\./g, '\\.')}\\s*$`, 'm'),
    `release.yml must upload supported Obsidian asset ${asset}.`,
  );
}

console.log(
  `[check:plugin-runtime-artifacts] Sync-safe runtime artifacts: ${sizes
    .map(({ artifactPath, size }) => `${path.basename(artifactPath)}=${size}`)
    .join(', ')}; release assets=${OBSIDIAN_SUPPORTED_RELEASE_ASSETS.join(', ')}`,
);
