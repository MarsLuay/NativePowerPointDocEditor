import assert from "node:assert/strict";
import { extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./renderer.mjs";

export async function loadEngine(fixtureName = "features.pptx") {
  const { PresentationEngine } = await loadPresentationEngineModule();
  return PresentationEngine.load(toArrayBuffer(await readDeck(fixtureName)));
}

export async function withDeck(fixtureName, fn) {
  return fn(await loadEngine(fixtureName));
}

export async function exportBytes(engine) {
  const exported = await engine.export();
  assert.ok(
    exported instanceof ArrayBuffer && exported.byteLength > 0,
    "export produced no bytes",
  );
  return exported;
}

export function assertSlideCount(engine, expectedSlides) {
  assert.equal(engine.slideCount, expectedSlides, `expected ${expectedSlides} slides, got ${engine.slideCount}`);
}

export function assertShapeIndex(label, shapeIndex) {
  assert.equal(
    Number.isInteger(shapeIndex) && shapeIndex >= 0,
    true,
    `${label}: expected a valid inserted shape index, got ${shapeIndex}`,
  );
}

export async function assertExportRoundTrips(label, engine, { expectedSlides = engine.slideCount } = {}) {
  for (let slideIndex = 0; slideIndex < expectedSlides; slideIndex += 1) {
    assert.match(engine.renderSlide(slideIndex).svg, /^<svg\b/, `${label}: slide ${slideIndex} must render to SVG`);
  }

  const exported = await exportBytes(engine);
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();
  const structure = validatePowerPointPackageStructure(inspectPowerPointPackage(exported), expectedSlides);
  assert.equal(
    structure.ok,
    true,
    `${label}: exported deck failed structural validation: ${JSON.stringify(structure)}`,
  );

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reloaded = await PresentationEngine.load(exported);
  assertSlideCount(reloaded, expectedSlides);
  for (let slideIndex = 0; slideIndex < expectedSlides; slideIndex += 1) {
    assert.match(
      reloaded.renderSlide(slideIndex).svg,
      /^<svg\b/,
      `${label}: reloaded slide ${slideIndex} must render to SVG`,
    );
  }

  return exported;
}

export async function readFixtureImage(fixture = "features.pptx") {
  const zip = await extractZip(toArrayBuffer(await readDeck(fixture)));
  const image = zip.binaryFiles.get("ppt/media/image1.png");
  assert.ok(image, `${fixture}: expected ppt/media/image1.png`);
  return image;
}
