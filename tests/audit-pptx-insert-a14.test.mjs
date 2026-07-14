import assert from "node:assert/strict";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const FIXTURE = "features.pptx";
const SLIDE_INDEX = 0;
const A14_EXT_MARKUP =
  '<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' +
  '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' +
  "</a:ext>" +
  '<a:ext uri="{5C8F9657-305E-4C14-9661-3D7EBBD3B07E}">' +
  '<a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
  "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>" +
  "</a14:hiddenFill></a:ext></a:extLst>";

async function injectA14BlipExtensions(buffer, blipIndex = 0) {
  const zip = await extractZip(buffer);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  assert.ok(slideXml, "fixture must include slide1.xml");
  assert.match(slideXml, /<a:blip\b/, "fixture must include at least one blip");

  let matchIndex = 0;
  const updated = slideXml.replace(/<a:blip\b([^>]*)(?:\/>|>([\s\S]*?)<\/a:blip>)/g, (match, attrs, inner) => {
    if (matchIndex++ !== blipIndex) return match;
    if (inner !== undefined) {
      return `<a:blip${attrs}>${A14_EXT_MARKUP}${inner}</a:blip>`;
    }
    return `<a:blip${attrs}>${A14_EXT_MARKUP}</a:blip>`;
  });
  assert.notEqual(updated, slideXml, "expected to inject a14 extensions into the fixture blip");

  return buildZip(buffer, new Map([[slidePath, updated]]));
}

test("insert shape preserves a14 blip extensions through export validation", async () => {
  const sourceBuffer = await injectA14BlipExtensions(toArrayBuffer(await readDeck(FIXTURE)));
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(sourceBuffer);

  const shapeIndex = await engine.addShapeGeometry(SLIDE_INDEX, "rect");
  assert.equal(Number.isInteger(shapeIndex) && shapeIndex >= 0, true);

  const exported = await engine.export();
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const validation = await validatePowerPointExportContents(sourceBuffer, exported);
  assert.equal(
    validation.ok,
    true,
    `export validation failed after shape insert: ${validation.errors.join("; ")}`,
  );

  const zip = await extractZip(exported);
  const slideXml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.match(slideXml, /<a14:useLocalDpi\b/, "exported slide should keep a14:useLocalDpi");
  assert.match(slideXml, /<a14:hiddenFill\b/, "exported slide should keep a14:hiddenFill");
});

test("insert shape preserves a14 on every blip when slide already has multiple images", async () => {
  let sourceBuffer = toArrayBuffer(await readDeck(FIXTURE));
  const blipCount = ((await extractZip(sourceBuffer)).textFiles.get("ppt/slides/slide1.xml").match(/<a:blip\b/g) ?? []).length;
  assert.ok(blipCount >= 1, "fixture should include at least one blip");
  for (let index = 0; index < blipCount; index++) {
    sourceBuffer = await injectA14BlipExtensions(sourceBuffer, index);
  }
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(sourceBuffer);

  const before = (await extractZip(sourceBuffer)).textFiles.get("ppt/slides/slide1.xml");
  const beforeUseLocalDpi = (before.match(/<a14:useLocalDpi\b/g) ?? []).length;
  assert.ok(beforeUseLocalDpi >= 1, "fixture should contain injected a14 blip extensions");

  const shapeIndex = await engine.addShapeGeometry(SLIDE_INDEX, "ellipse");
  assert.equal(Number.isInteger(shapeIndex) && shapeIndex >= 0, true);

  const exported = await engine.export();
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const validation = await validatePowerPointExportContents(sourceBuffer, exported);
  assert.equal(
    validation.ok,
    true,
    `export validation failed after shape insert on multi-blip slide: ${validation.errors.join("; ")}`,
  );

  const slideXml = (await extractZip(exported)).textFiles.get("ppt/slides/slide1.xml");
  assert.equal(
    (slideXml.match(/<a14:useLocalDpi\b/g) ?? []).length,
    beforeUseLocalDpi,
    "every a14:useLocalDpi instance should survive shape insert",
  );
});
