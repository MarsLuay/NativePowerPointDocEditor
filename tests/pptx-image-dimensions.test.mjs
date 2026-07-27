import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function loadImageDimensionsModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'npde-image-dimensions-test-'));
  const outfile = path.join(outdir, 'image-dimensions.cjs');
  await build({
    absWorkingDir: projectRoot,
    entryPoints: ['src/powerpoint/imageDimensions.ts'],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
  });
  return import(pathToFileURL(outfile).href);
}

test('PowerPoint image dimensions read raster headers without a browser decoder', async () => {
  const { readRasterImageDimensions } = await loadImageDimensionsModule();

  const png = new Uint8Array(24);
  png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  new DataView(png.buffer).setUint32(16, 1920);
  new DataView(png.buffer).setUint32(20, 1080);
  assert.deepEqual(readRasterImageDimensions(png), { width: 1920, height: 1080 });

  const jpeg = new Uint8Array([
    0xff, 0xd8, 0xff, 0xc2, 0x00, 0x11, 0x08, 0x02, 0xd0, 0x05, 0x00,
    0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00,
  ]);
  assert.deepEqual(readRasterImageDimensions(jpeg), { width: 1280, height: 720 });

  const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x80, 0x02, 0xe0, 0x01]);
  assert.deepEqual(readRasterImageDimensions(gif), { width: 640, height: 480 });
  assert.equal(readRasterImageDimensions(new Uint8Array([0x89, 0x50])), null);
});
