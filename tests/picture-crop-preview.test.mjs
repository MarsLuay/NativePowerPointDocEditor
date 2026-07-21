import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function loadHelper() {
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "picture-crop-preview-"));
  const outfile = path.join(outputDirectory, "picture-crop-preview.cjs");
  await build({
    entryPoints: [path.join(projectRoot, "src/powerpoint/pictureCropPreview.ts")],
    bundle: true,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  return require(outfile);
}

const { scaleCroppedPicturePreview } = await loadHelper();

test("scaleCroppedPicturePreview doubles the frame and keeps crop ratios", () => {
  // Frame 144x108; source expanded so left=25%, top=20%, right=15%, bottom=30%.
  const image = { x: 552, y: -7.2, width: 240, height: 216 };
  const clip = { x: 612, y: 36, width: 144, height: 108 };
  const nextClip = { x: 612, y: 36, width: 288, height: 216 };

  const scaled = scaleCroppedPicturePreview(image, clip, nextClip);

  assert.deepEqual(scaled.clip, nextClip);
  assert.ok(Math.abs(scaled.image.width - 480) < 1e-6);
  assert.ok(Math.abs(scaled.image.height - 432) < 1e-6);
  assert.ok(Math.abs(scaled.image.x - 492) < 1e-6);
  assert.ok(Math.abs(scaled.image.y - -50.4) < 1e-6);

  // Visible inset fractions stay the same.
  assert.ok(Math.abs((scaled.clip.x - scaled.image.x) / scaled.image.width - 0.25) < 1e-6);
  assert.ok(Math.abs((scaled.clip.y - scaled.image.y) / scaled.image.height - 0.2) < 1e-6);
  assert.ok(Math.abs(scaled.clip.width / scaled.image.width - 0.6) < 1e-6);
  assert.ok(Math.abs(scaled.clip.height / scaled.image.height - 0.5) < 1e-6);
});

test("scaleCroppedPicturePreview moves the frame without changing size ratios", () => {
  const image = { x: 100, y: 50, width: 200, height: 100 };
  const clip = { x: 120, y: 60, width: 160, height: 80 };
  const nextClip = { x: 220, y: 160, width: 160, height: 80 };

  const scaled = scaleCroppedPicturePreview(image, clip, nextClip);
  assert.deepEqual(scaled.clip, nextClip);
  assert.equal(scaled.image.width, 200);
  assert.equal(scaled.image.height, 100);
  assert.equal(scaled.image.x, 200);
  assert.equal(scaled.image.y, 150);
});
