import assert from "node:assert/strict";
import { test } from "node:test";
import { extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

// Audit coverage for the "PPTX · Insert objects" feature group. Every insert
// operation that is cleanly callable headlessly is exercised through the public
// PresentationEngine API, then the result is rendered and exported, and the
// exported deck is asserted to remain a structurally valid, still-loadable
// single-slide presentation.

const FIXTURE = "features.pptx";
const SLIDE_INDEX = 0;

async function loadFreshEngine() {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const bytes = await readDeck(FIXTURE);
  return PresentationEngine.load(toArrayBuffer(bytes));
}

async function readFixtureImage() {
  const bytes = await readDeck(FIXTURE);
  const zip = await extractZip(toArrayBuffer(bytes));
  const image = zip.binaryFiles.get("ppt/media/image1.png");
  assert.ok(image, "fixture must ship ppt/media/image1.png to drive the image insert path");
  return image;
}

// Assert a freshly inserted shape index is a real, non-negative integer.
function assertShapeIndex(label, shapeIndex) {
  assert.equal(
    Number.isInteger(shapeIndex) && shapeIndex >= 0,
    true,
    `${label}: expected a valid inserted shape index, got ${shapeIndex}`,
  );
}

// Render the slide, export, and assert the export round-trips to a structurally
// valid deck that still loads and renders as a single-slide presentation.
async function assertExportRoundTrips(label, engine) {
  const svg = engine.renderSlide(SLIDE_INDEX).svg;
  assert.match(svg, /^<svg\b/, `${label}: slide must render to SVG after insert`);

  const exported = await engine.export();
  assert.ok(exported instanceof ArrayBuffer && exported.byteLength > 0, `${label}: export produced no bytes`);

  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();
  const inspected = inspectPowerPointPackage(exported);
  const structure = validatePowerPointPackageStructure(inspected, 1);
  assert.equal(structure.ok, true, `${label}: exported deck failed structural validation: ${JSON.stringify(structure)}`);

  const { PresentationEngine } = await loadPresentationEngineModule();
  const reloaded = await PresentationEngine.load(exported);
  assert.equal(reloaded.slideCount, 1, `${label}: reloaded deck must keep its single slide`);
  assert.match(reloaded.renderSlide(SLIDE_INDEX).svg, /^<svg\b/, `${label}: reloaded deck must still render`);

  return exported;
}

test("insert text box round-trips to a valid, loadable deck", async () => {
  const engine = await loadFreshEngine();
  const shapeIndex = await engine.addTextBox(SLIDE_INDEX);
  assertShapeIndex("text box", shapeIndex);
  await assertExportRoundTrips("text box", engine);
});

test("insert text box honors a requested slide-space origin and clamps it to the slide", async () => {
  const engine = await loadFreshEngine();
  const slideSize = await engine.getSlideSizeEmu();
  const shapeIndex = await engine.addTextBox(SLIDE_INDEX, { x: -50, y: slideSize.cy + 50 });
  assertShapeIndex("positioned text box", shapeIndex);

  const slideXml = engine.getSlideXml(SLIDE_INDEX);
  const textBoxNameIndex = slideXml.indexOf('name="TextBox"');
  const textBoxStart = slideXml.lastIndexOf('<p:sp>', textBoxNameIndex);
  const textBoxEnd = slideXml.indexOf('</p:sp>', textBoxNameIndex);
  const textBox = textBoxStart >= 0 && textBoxEnd >= 0
    ? slideXml.slice(textBoxStart, textBoxEnd + '</p:sp>'.length)
    : null;
  assert.ok(textBox, "expected the inserted text box shape in slide XML");
  const maxY = slideSize.cy - 685800;
  assert.match(textBox, new RegExp(`<a:off\\b[^>]*\\bx="0"[^>]*\\by="${maxY}"`));
	assert.match(
		textBox,
		/<a:bodyPr\b[^>]*\blIns="91440"[^>]*\btIns="45720"[^>]*\brIns="91440"[^>]*\bbIns="45720"/,
		"inserted text boxes need PowerPoint's normal interior buffer around their text",
	);
	assert.match(textBox, /<p:cNvSpPr\b[^>]*\btxBox="1"/, "inserted text boxes must be marked txBox");
	assert.match(textBox, /<a:spAutoFit\s*\/>/, "inserted text boxes must use spAutoFit so overflow does not shrink fonts");
  await assertExportRoundTrips("positioned text box", engine);
});

test("insert shape geometries round-trip to a valid, loadable deck", async (t) => {
  const geometries = ["rect", "ellipse", "roundRect", "line", "rightArrow", "leftArrow", "upArrow", "downArrow"];
  for (const geometry of geometries) {
    await t.test(geometry, async () => {
      const engine = await loadFreshEngine();
      const shapeIndex = await engine.addShapeGeometry(SLIDE_INDEX, geometry);
      assertShapeIndex(geometry, shapeIndex);
      await assertExportRoundTrips(geometry, engine);
    });
  }
});

test("insert image round-trips to a valid, loadable deck", async () => {
  const engine = await loadFreshEngine();
  const image = await readFixtureImage();
  const shapeIndex = await engine.addImage(SLIDE_INDEX, image, "image/png");
  assertShapeIndex("image", shapeIndex);
  await assertExportRoundTrips("image", engine);
});

test("insert table round-trips and embeds the new table grid", async () => {
  const engine = await loadFreshEngine();
  const shapeIndex = await engine.addTable(SLIDE_INDEX, 3, 4);
  assertShapeIndex("table", shapeIndex);
  const exported = await assertExportRoundTrips("table", engine);

  const zip = await extractZip(exported);
  const slideXml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.ok(slideXml, "table export must retain slide1.xml");
  // The 3x4 table the engine inserts must survive export as 3 rows / 4 grid cols.
  assert.equal((slideXml.match(/<a:tr\b/g) || []).length >= 3, true, "expected at least 3 table rows in export");
  assert.equal((slideXml.match(/<a:gridCol\b/g) || []).length >= 4, true, "expected at least 4 grid columns in export");
});

test("insert chart round-trips and writes the chart package parts", async () => {
  const engine = await loadFreshEngine();
  const shapeIndex = await engine.addChart(SLIDE_INDEX);
  assertShapeIndex("chart", shapeIndex);
  const exported = await assertExportRoundTrips("chart", engine);

  const zip = await extractZip(exported);
  const slideXml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.match(slideXml, /<c:chart\b/, "exported slide must reference an inserted chart");

  // A newly inserted chart must ship its own chart part + embedded workbook so
  // the package merge logic kept the freshly created parts.
  const chartParts = [...zip.textFiles.keys()].filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p));
  assert.ok(chartParts.length >= 1, "exported deck must contain at least one chart part");
  const workbooks = [...zip.binaryFiles.keys()].filter((p) => /^ppt\/embeddings\/.*\.xlsx$/.test(p));
  assert.ok(workbooks.length >= 1, "exported chart must keep its embedded workbook");
});

test("apply bullet list writes hanging indent and renders marker outside run text", async () => {
  const engine = await loadFreshEngine();
  await engine.applyListStyle(SLIDE_INDEX, 0, 0, "bullet");

  const exported = await assertExportRoundTrips("bulleted list", engine);
  const zip = await extractZip(exported);
  const slideXml = zip.textFiles.get("ppt/slides/slide1.xml");
  assert.ok(slideXml, "bulleted list export must retain slide1.xml");
  assert.match(
    slideXml,
    /<a:pPr\b[^>]*\bmarL="285750"[^>]*\bindent="-285750"[^>]*>/,
    "list paragraph gets a hanging indent instead of an inline-only marker",
  );
  assert.match(slideXml, /<a:buFont\b[^>]*\btypeface="Arial"/, "bullet marker gets a stable bullet font");
  assert.match(slideXml, /<a:buChar\b[^>]*\bchar="•"/, "bullet marker is written");

  const rendered = engine.renderSlide(SLIDE_INDEX).svg;
  const document = new globalThis.DOMParser().parseFromString(rendered, "text/xml");
  const shape = Array.from(document.getElementsByTagName("g")).find(
    (element) => element.getAttribute("data-ooxml-shape-idx") === "0",
  );
  assert.ok(shape, "fixture title shape renders");

  const paragraphContainers = Array.from(shape.getElementsByTagName("tspan")).filter(
    (element) => element.getAttribute("data-ooxml-para-idx") === "0",
  );
  const hasRun = (element) =>
    Array.from(element.getElementsByTagName("tspan")).some((child) =>
      child.getAttribute("data-ooxml-run-idx") !== null
    );
  const markerContainer = paragraphContainers.find((element) =>
    (element.textContent || "").trim().startsWith("•") && !hasRun(element)
  );
  const runContainer = paragraphContainers.find((element) => hasRun(element));

  assert.ok(markerContainer, "bullet marker renders in its own run-less container");
  assert.ok(runContainer, "paragraph run text still renders");
  assert.notEqual(markerContainer, runContainer, "marker and text are separate visual containers");
  assert.equal((runContainer.textContent || "").includes("•"), false, "run text does not include the bullet glyph");
});
