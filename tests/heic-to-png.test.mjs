import assert from 'node:assert/strict';
import test from 'node:test';

import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

async function loadHeicHelpers() {
  const outfile = join(root, 'results', 'heic-to-png-test.cjs');
  await build({
    entryPoints: [join(root, 'src/powerpoint/heicToPng.ts')],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    logLevel: 'silent',
    // Decode path needs canvas; unit tests only exercise detection helpers.
    external: ['heic-decode'],
  });
  return require(outfile);
}

function ftypBox(brand) {
  const bytes = new Uint8Array(12);
  bytes[0] = 0;
  bytes[1] = 0;
  bytes[2] = 0;
  bytes[3] = 12;
  bytes.set([0x66, 0x74, 0x79, 0x70], 4); // ftyp
  bytes.set(Array.from(brand, (ch) => ch.charCodeAt(0)), 8);
  return bytes;
}

test('heic-decode brand check needs Uint8Array not ArrayBuffer', async () => {
  // Mirrors heic-decode/lib.js uint8ArrayUtf8ByteString — ArrayBuffer.slice is
  // not iterable, so spreading it throws the paste error we hit in Obsidian.
  const bytes = ftypBox('heic');
  const asArrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  assert.throws(
    () => String.fromCharCode(...asArrayBuffer.slice(8, 12)),
    /iterable|iterator/i,
  );
  assert.equal(
    String.fromCharCode(...bytes.slice(8, 12)),
    'heic',
  );
});

test('fitWithinMaxEdge downscales phone-sized HEIC rasters', async () => {
  const { fitWithinMaxEdge, HEIC_PNG_MAX_EDGE } = await loadHeicHelpers();
  assert.equal(HEIC_PNG_MAX_EDGE, 2560);
  assert.deepEqual(fitWithinMaxEdge(2000, 1500), { width: 2000, height: 1500, scale: 1 });
  const fitted = fitWithinMaxEdge(4032, 3024);
  assert.equal(fitted.width, 2560);
  assert.equal(fitted.height, 1920);
  assert.ok(fitted.scale < 1);
  assert.equal(Math.max(fitted.width, fitted.height), HEIC_PNG_MAX_EDGE);
});

test('looksLikeHeicBytes recognizes common HEIC/HEIF brands', async () => {
  const {
    looksLikeHeicBytes,
    isHeicMimeType,
    isHeicExtension,
    shouldConvertHeicToPng,
  } = await loadHeicHelpers();

  assert.equal(looksLikeHeicBytes(ftypBox('heic')), true);
  assert.equal(looksLikeHeicBytes(ftypBox('mif1')), true);
  assert.equal(looksLikeHeicBytes(ftypBox('msf1')), true);
  assert.equal(looksLikeHeicBytes(ftypBox('heif')), false);
  assert.equal(looksLikeHeicBytes(new Uint8Array([0x89, 0x50, 0x4e, 0x47])), false);

  assert.equal(isHeicMimeType('image/heic'), true);
  assert.equal(isHeicMimeType('image/heif; codecs=hevc'), true);
  assert.equal(isHeicMimeType('image/png'), false);
  assert.equal(isHeicExtension('HEIC'), true);
  assert.equal(isHeicExtension('png'), false);

  assert.equal(shouldConvertHeicToPng(ftypBox('heic')), true);
  assert.equal(shouldConvertHeicToPng(new Uint8Array([1, 2, 3]), 'image/heic'), true);
  assert.equal(shouldConvertHeicToPng(new Uint8Array([1, 2, 3]), null, 'photo.heif'), true);
  assert.equal(shouldConvertHeicToPng(new Uint8Array([1, 2, 3]), 'image/png', 'photo.png'), false);
});

test('clipboard image helper prefers HEIC MIME types', async () => {
  const outfile = join(root, 'results', 'clipboard-image-test.cjs');
  await build({
    entryPoints: [join(root, 'src/powerpoint/clipboardImage.ts')],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    logLevel: 'silent',
  });
  // Exercise DataTransfer-shaped file list without a real DOM clipboard.
  const { readDataTransferRasterImage } = require(outfile);
  const heicBytes = ftypBox('heic');
  const file = {
    name: 'IMG_0001.HEIC',
    type: '',
    arrayBuffer: async () => heicBytes.buffer.slice(0),
  };
  const dataTransfer = {
    files: [file],
    items: [],
  };
  const image = await readDataTransferRasterImage(dataTransfer);
  assert.ok(image);
  assert.equal(image.fileName, 'IMG_0001.HEIC');
  assert.equal(image.bytes.byteLength, heicBytes.byteLength);
});
