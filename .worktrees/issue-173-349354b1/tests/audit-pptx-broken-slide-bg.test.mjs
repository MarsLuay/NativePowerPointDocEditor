import assert from "node:assert/strict";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import { loadPresentationEngineModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const FIXTURE = "features.pptx";
const SLIDE_INDEX = 0;
const ONE_PIXEL_LAYOUT_BACKGROUND = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  "base64",
);
const BROKEN_SLIDE_BACKGROUND = [
  "<p:bg>",
  "<p:bgPr>",
  '<a:blipFill rotWithShape="1">',
  '<a:blip r:embed="rIdNotes"/>',
  "<a:stretch><a:fillRect/></a:stretch>",
  "</a:blipFill>",
  "<a:effectLst/>",
  "</p:bgPr>",
  "</p:bg>",
].join("");

async function injectBrokenSlideBackgroundWithLayoutFallback(buffer) {
  const zip = await extractZip(buffer);
  const slidePath = "ppt/slides/slide1.xml";
  const slideRelsPath = "ppt/slides/_rels/slide1.xml.rels";
  const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
  const layoutRelsPath = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
  const slideXml = zip.textFiles.get(slidePath);
  const slideRelsXml = zip.textFiles.get(slideRelsPath);
  const layoutXml = zip.textFiles.get(layoutPath);
  const layoutRelsXml = zip.textFiles.get(layoutRelsPath);
  assert.ok(slideXml && slideRelsXml && layoutXml && layoutRelsXml);

  const layoutBackgroundXml = [
    "<p:bg><p:bgPr><a:blipFill>",
    '<a:blip r:embed="rIdImage"/>',
    "<a:stretch><a:fillRect/></a:stretch>",
    "</a:blipFill><a:effectLst/></p:bgPr></p:bg>",
  ].join("");
  const fixedLayoutXml = layoutXml.includes("<p:bg")
    ? layoutXml
    : layoutXml.replace("<p:cSld name=", `${layoutBackgroundXml}<p:cSld name=`);
  const patchedLayoutRelsXml = layoutRelsXml.replace(
    "</Relationships>",
    '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-bg.png"/></Relationships>',
  );
  const patchedSlideXml = slideXml.replace("<p:cSld", `${BROKEN_SLIDE_BACKGROUND}<p:cSld`);
  const patchedSlideRelsXml = slideRelsXml.replace(
    "</Relationships>",
    '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
  );

  return buildZip(
    buffer,
    new Map([
      [layoutPath, fixedLayoutXml],
      [layoutRelsPath, patchedLayoutRelsXml],
      [slidePath, patchedSlideXml],
      [slideRelsPath, patchedSlideRelsXml],
    ]),
    undefined,
    new Map([["ppt/media/layout-bg.png", ONE_PIXEL_LAYOUT_BACKGROUND]]),
  );
}

test("broken slide background falls back to layout background image", async () => {
  const sourceBuffer = await injectBrokenSlideBackgroundWithLayoutFallback(
    toArrayBuffer(await readDeck(FIXTURE)),
  );
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(sourceBuffer);

  const background = engine.getSlideBackgroundDescribe(SLIDE_INDEX);
  assert.ok(background.imageHref, "expected layout background after broken slide p:bg");
  assert.match(background.imageHref, /^data:image\/png;base64,/);

  const svg = engine.renderSlide(SLIDE_INDEX).svg;
  assert.match(svg, /<image\b[^>]*href="data:image\/png;base64,/);
});
