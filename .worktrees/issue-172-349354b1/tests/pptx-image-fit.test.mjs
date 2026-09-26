import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function loadImageFitModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'npde-image-fit-test-'));
  const outfile = path.join(outdir, 'image-fit.cjs');
  await build({
    absWorkingDir: projectRoot,
    entryPoints: ['src/powerpoint/imageFit.ts'],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
  });
  return import(pathToFileURL(outfile).href);
}

test('PowerPoint image fit center-crops to cover without distortion', async () => {
  const { computeCenteredCoverCrop } = await loadImageFitModule();

  assert.deepEqual(
    computeCenteredCoverCrop({ width: 1600, height: 900 }, 400, 300),
    { left: 12.5, top: 0, right: 12.5, bottom: 0 },
  );
  assert.deepEqual(
    computeCenteredCoverCrop({ width: 900, height: 1600 }, 400, 300),
    { left: 0, top: 28.90625, right: 0, bottom: 28.90625 },
  );
  assert.deepEqual(
    computeCenteredCoverCrop({ width: 1600, height: 900 }, 1600, 900),
    { left: 0, top: 0, right: 0, bottom: 0 },
  );
  assert.equal(computeCenteredCoverCrop(null, 400, 300), null);
});

test('PowerPoint image fit keeps direct insertions inside their default bounds', async () => {
  const { fitImageWithinBounds } = await loadImageFitModule();

  assert.deepEqual(fitImageWithinBounds({ width: 4000, height: 2000 }, 320, 240), { width: 320, height: 160 });
  assert.deepEqual(fitImageWithinBounds({ width: 1000, height: 2000 }, 320, 240), { width: 120, height: 240 });
  assert.deepEqual(fitImageWithinBounds(null, 320, 240), { width: 320, height: 240 });
});
