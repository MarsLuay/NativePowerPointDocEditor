import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

async function loadRuntimeArtifactLoader() {
  const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-runtime-artifact-loader-'));
  const outfile = path.join(outputDirectory, 'runtime-artifact-loader.cjs');
  await build({
    entryPoints: [path.join(projectRoot, 'src/powerpoint/runtimeArtifactLoader.ts')],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
  });
  return require(outfile);
}

test('optional PowerPoint runtimes use the configured plugin resource URLs', async () => {
  const {
    PPTX_RUNTIME_ARTIFACTS,
    configurePptxRuntimeArtifactLoader,
    getPptxRuntimeArtifactResource,
    loadPptxRuntimeArtifact,
  } = await loadRuntimeArtifactLoader();

  assert.throws(
    () => getPptxRuntimeArtifactResource('pptx-wasm-renderer.mjs'),
    /was not configured/,
  );

  const resolvedPaths = [];
  configurePptxRuntimeArtifactLoader((artifact) => {
    const path = `.obsidian/plugins/native-powerpoint-doc-editor/${artifact}`;
    const resourceUrl = `data:text/javascript,${encodeURIComponent(`export const loadedArtifact = ${JSON.stringify(artifact)};`)}`;
    resolvedPaths.push(path);
    return { path, resourceUrl };
  });

  assert.deepEqual(resolvedPaths, PPTX_RUNTIME_ARTIFACTS.map(
    (artifact) => `.obsidian/plugins/native-powerpoint-doc-editor/${artifact}`,
  ));
  for (const artifact of PPTX_RUNTIME_ARTIFACTS) {
    const resource = getPptxRuntimeArtifactResource(artifact);
    assert.equal(resource.path, `.obsidian/plugins/native-powerpoint-doc-editor/${artifact}`);
    const module = await loadPptxRuntimeArtifact(artifact);
    assert.equal(module.loadedArtifact, artifact);
  }
});
