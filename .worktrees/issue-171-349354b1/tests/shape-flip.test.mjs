import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { build } from "esbuild";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

let shapeFlipModulePromise;

async function loadShapeFlipModule() {
  shapeFlipModulePromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "native-powerpoint-shape-flip-"));
    const outfile = path.join(outputDirectory, "shape-flip.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/shapeFlipTransforms.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });
    return import(pathToFileURL(outfile).href);
  })();
  return shapeFlipModulePromise;
}

function pathToFileURL(filePath) {
  return new URL(`file://${filePath}`);
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function mockShape(attrs) {
  return {
    getAttribute(name) {
      return attrs[name] ?? null;
    }
  };
}

test("flipShape toggles flipH in the exported OOXML model", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const deck = toArrayBuffer(await readFile(path.join(projectRoot, "tests/fixtures/decks/features.pptx")));

  const engine = await PresentationEngine.load(deck);
  const slide = 0;
  const pictureIndex = 2;

  const before = engine.renderShape(slide, pictureIndex);
  assert.match(before, /data-ooxml-shape-type="picture"/);
  assert.doesNotMatch(before, /data-ooxml-flip-h="1"/);

  await engine.flipShape(slide, pictureIndex, "horizontal");
  const afterFlip = engine.renderShape(slide, pictureIndex);
  assert.match(afterFlip, /data-ooxml-flip-h="1"/);

  await engine.flipShape(slide, pictureIndex, "horizontal");
  const afterToggleOff = engine.renderShape(slide, pictureIndex);
  assert.doesNotMatch(afterToggleOff, /data-ooxml-flip-h="1"/);
});

test("flipTransformForShape builds a center-anchored mirror transform", async () => {
  const { flipTransformForShape, shapeFlipRenderedInGeometry } = await loadShapeFlipModule();
  const scale = 9525;
  const shape = mockShape({
    "data-ooxml-flip-h": "1",
    "data-ooxml-x": "1000000",
    "data-ooxml-y": "2000000",
    "data-ooxml-cx": "3000000",
    "data-ooxml-cy": "4000000"
  });

  const transform = flipTransformForShape(shape, scale);
  assert.match(transform ?? "", /scale\(-1\s*,\s*1\)/);
  assert.equal(shapeFlipRenderedInGeometry(mockShape({ "data-ooxml-geom": "line" })), true);
  assert.equal(shapeFlipRenderedInGeometry(mockShape({ "data-ooxml-geom": "straightConnector1" })), true);
  assert.equal(shapeFlipRenderedInGeometry(mockShape({ "data-ooxml-geom": "rect" })), false);
});

test("live SVG flip wrappers are created detached from their owning document", async () => {
  const source = await readFile(
    path.join(projectRoot, "src/powerpoint/shapeFlipTransforms.ts"),
    "utf8",
  );

  assert.match(source, /shape\.ownerDocument\.createElementNS\(SVG_NAMESPACE, 'g'\)/);
  assert.doesNotMatch(source, /createSvg\(/);
});

test("flipShape works for autoshapes, not only pictures", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const { flipTransformForShape } = await loadShapeFlipModule();
  const deck = toArrayBuffer(await readFile(path.join(projectRoot, "tests/fixtures/decks/features.pptx")));

  const engine = await PresentationEngine.load(deck);
  await engine.flipShape(0, 0, "vertical");

  const fragment = engine.renderShape(0, 0);
  assert.match(fragment, /data-ooxml-flip-v="1"/);

  const match = fragment.match(/data-ooxml-x="(\d+)"[^>]*data-ooxml-y="(\d+)"[^>]*data-ooxml-cx="(\d+)"[^>]*data-ooxml-cy="(\d+)"/);
  assert.ok(match, "expected shape transform attrs on rendered fragment");
  const shape = mockShape({
    "data-ooxml-flip-v": "1",
    "data-ooxml-x": match[1],
    "data-ooxml-y": match[2],
    "data-ooxml-cx": match[3],
    "data-ooxml-cy": match[4]
  });
  assert.match(flipTransformForShape(shape, 9525) ?? "", /scale\(\s*1\s*,\s*-1\s*\)/);
});
