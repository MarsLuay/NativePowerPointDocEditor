import assert from 'node:assert/strict';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { OBSIDIAN_SUPPORTED_RELEASE_ASSETS } from './lib/pptx-runtime-artifact-spec.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const syncStandardLimitBytes = 5_000_000;
// main.js embeds the optional runtime payloads for community installs. Keep a
// small, explicit headroom budget for that bundle while retaining the strict
// Sync limit for every materialized sidecar.
const embeddedMainBundleLimitBytes = 5_010_000;
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

const allowOversizeMain = process.env.NPDE_ALLOW_OVERSIZE_MAIN === '1';

for (const { artifactPath, size } of sizes) {
  const artifact = path.basename(artifactPath);
  if (artifact === 'main.js' && allowOversizeMain && size > syncStandardLimitBytes) {
    console.warn(
      `[check:plugin-runtime-artifacts] BRAT-only override: ${artifact}=${size} bytes exceeds the ${syncStandardLimitBytes}-byte Sync Standard budget.`,
    );
    continue;
  }
  const limit = artifact === 'main.js' ? embeddedMainBundleLimitBytes : syncStandardLimitBytes;
  assert.ok(
    size <= limit,
    `${artifactPath} is ${size} bytes, above its ${limit}-byte runtime artifact limit.`,
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
