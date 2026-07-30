import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
import { createHash } from 'node:crypto';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

async function loadMaterializerWithPayload(bytes) {
  const sha256 = createHash('sha256').update(bytes).digest('hex');
  const gzipBase64 = gzipSync(bytes, { level: 9 }).toString('base64');
  const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-runtime-materializer-'));
  const payloadPath = path.join(outputDirectory, 'payloads.json');
  await writeFile(
    payloadPath,
    `${JSON.stringify([{
      artifact: 'pptx-js-engine.mjs',
      sha256,
      gzipBase64,
    }])}\n`,
  );
  const outfile = path.join(outputDirectory, 'materializer.cjs');
  await build({
    entryPoints: [path.join(projectRoot, 'src/powerpoint/runtimeArtifactMaterializer.ts')],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
    plugins: [
      {
        name: 'stub-runtime-payloads',
        setup(buildApi) {
          buildApi.onResolve({ filter: /runtimeArtifactPayloads(\.json)?$/ }, () => ({
            path: payloadPath,
          }));
        },
      },
    ],
  });
  return require(outfile);
}

test('ensurePptxRuntimeArtifacts writes missing gzip-embedded sidecars', async () => {
  const expected = Buffer.from('export const createPptxJsEngine = () => ({});\n', 'utf8');
  const { ensurePptxRuntimeArtifacts } = await loadMaterializerWithPayload(expected);
  const files = new Map();
  await ensurePptxRuntimeArtifacts(
    'plugin',
    {
      exists: async (target) => files.has(target),
      read: async (target) => {
        const value = files.get(target);
        if (!value) throw new Error(`missing ${target}`);
        return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength);
      },
      write: async (target, data) => {
        files.set(target, Buffer.from(data));
      },
    },
    (dir, artifact) => `${dir}/${artifact}`,
  );
  assert.equal(files.get('plugin/pptx-js-engine.mjs').toString('utf8'), expected.toString('utf8'));
});

test('ensurePptxRuntimeArtifacts skips matching on-disk sidecars', async () => {
  const expected = Buffer.from('export const createPptxJsEngine = () => ({ ok: true });\n', 'utf8');
  const { ensurePptxRuntimeArtifacts } = await loadMaterializerWithPayload(expected);
  let writes = 0;
  const files = new Map([['plugin/pptx-js-engine.mjs', expected]]);
  await ensurePptxRuntimeArtifacts(
    'plugin',
    {
      exists: async (target) => files.has(target),
      read: async (target) => {
        const value = files.get(target);
        return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength);
      },
      write: async () => {
        writes += 1;
      },
    },
    (dir, artifact) => `${dir}/${artifact}`,
  );
  assert.equal(writes, 0);
});
