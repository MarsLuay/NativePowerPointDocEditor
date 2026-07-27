import assert from "node:assert/strict";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
  loadShapeClipboardModule,
  loadSlideExtensionPreserveModule,
} from "./helpers/load-plugin-modules.mjs";
import { createRenderer, readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const editableFixtures = ["features.pptx", "features.ppsx", "features.potx"];
const macroFixtures = ["macro-view-only.pptm", "macro-view-only.ppsm", "macro-view-only.potm"];
const fixtureTitleParagraph = '<a:p><a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Native PowerPoint fixture</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>';
const onePixelLayoutBackground = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  "base64"
);

function paragraphXml(text, properties = 'lang="en-US" sz="2800" b="0"') {
  return `<a:p><a:r><a:rPr ${properties}/><a:t>${text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`;
}

async function createTitleParagraphFixture(paragraphs) {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const sourceZip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = sourceZip.textFiles.get(slidePath);
  assert.ok(slideXml);
  return buildZip(source, new Map([[slidePath, slideXml.replace(fixtureTitleParagraph, paragraphs.join(""))]]));
}

async function createLayoutBackgroundRelationshipCollisionFixture() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const sourceZip = await extractZip(source);
  const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
  const layoutRelsPath = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
  const layoutXml = sourceZip.textFiles.get(layoutPath);
  const layoutRelsXml = sourceZip.textFiles.get(layoutRelsPath);
  assert.ok(layoutXml);
  assert.ok(layoutRelsXml);

  const backgroundXml = [
    '<p:bg><p:bgPr><a:blipFill>',
    '<a:blip r:embed="rIdImage"/>',
    '<a:stretch><a:fillRect/></a:stretch>',
    '</a:blipFill><a:effectLst/></p:bgPr></p:bg>'
  ].join("");
  const patchedLayoutXml = layoutXml.replace('<p:cSld name="Blank"><p:spTree>', `<p:cSld name="Blank">${backgroundXml}<p:spTree>`);
  const patchedLayoutRelsXml = layoutRelsXml.replace(
    '</Relationships>',
    '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-bg.png"/></Relationships>'
  );

  return buildZip(
    source,
    new Map([
      [layoutPath, patchedLayoutXml],
      [layoutRelsPath, patchedLayoutRelsXml],
    ]),
    undefined,
    new Map([["ppt/media/layout-bg.png", onePixelLayoutBackground]])
  );
}

async function createCroppedLayoutBackgroundFixture() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const sourceZip = await extractZip(source);
  const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
  const layoutRelsPath = "ppt/slideLayouts/_rels/slideLayout1.xml.rels";
  const layoutXml = sourceZip.textFiles.get(layoutPath);
  const layoutRelsXml = sourceZip.textFiles.get(layoutRelsPath);
  assert.ok(layoutXml);
  assert.ok(layoutRelsXml);

  const backgroundXml = [
    '<p:bg><p:bgPr><a:blipFill>',
    '<a:blip r:embed="rIdImage"/>',
    '<a:srcRect l="25000" t="0" r="0" b="0"/>',
    '<a:stretch><a:fillRect/></a:stretch>',
    '</a:blipFill><a:effectLst/></p:bgPr></p:bg>'
  ].join("");
  const patchedLayoutXml = layoutXml.replace('<p:cSld name="Blank"><p:spTree>', `<p:cSld name="Blank">${backgroundXml}<p:spTree>`);
  const patchedLayoutRelsXml = layoutRelsXml.replace(
    '</Relationships>',
    '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-bg.png"/></Relationships>'
  );

  return buildZip(
    source,
    new Map([
      [layoutPath, patchedLayoutXml],
      [layoutRelsPath, patchedLayoutRelsXml],
    ]),
    undefined,
    new Map([["ppt/media/layout-bg.png", onePixelLayoutBackground]])
  );
}

async function createCroppedPictureFixture() {
  const input = await readDeck("features.pptx");
  const source = toArrayBuffer(input);
  const sourceZip = await extractZip(source);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = sourceZip.textFiles.get(slidePath);
  assert.ok(slideXml);

  const cropped = slideXml.replace(
    '<p:blipFill><a:blip r:embed="rIdImage"/><a:stretch>',
    '<p:blipFill><a:blip r:embed="rIdImage"/><a:srcRect l="0" t="23370" r="0" b="23370"/><a:stretch>',
  );
  assert.notEqual(cropped, slideXml, "fixture image should accept the authored crop");
  return buildZip(source, new Map([[slidePath, cropped]]));
}

test("pptx, ppsx, and potx fixtures load, render, export, and validate", async (t) => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
    validatePowerPointPackageStructure,
  } = await loadPowerPointPackageModule();

  for (const name of editableFixtures) {
    await t.test(name, async () => {
      const input = await readDeck(name);
      const original = inspectPowerPointPackage(toArrayBuffer(input));
      assert.equal(validatePowerPointPackageStructure(original, 1).ok, true);

      const renderer = await createRenderer(input);
      assert.equal(renderer.getSlideCount(), 1);
      assert.match(renderer.renderSlideSvg(0), /^<svg\b/);

      const output = await renderer.exportPptx();
      const exported = inspectPowerPointPackage(output);
      assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
      assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);
    });
  }
});

test("feature fixture preserves rich OOXML parts during an untouched round trip", async () => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();
  const input = await readDeck("features.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  const renderer = await createRenderer(input);

  assert.deepEqual(renderer.getSlideNotes(0), ["Fixture speaker notes survive round trip."]);
  const initialSlide = renderer.getSlideOoxml(0);
  assert.match(initialSlide, /<a:hlinkClick\b/);
  assert.match(initialSlide, /<c:chart\b/);
  assert.match(initialSlide, /<a:tbl>/);
  assert.match(initialSlide, /<p:grpSp>/);
  assert.match(initialSlide, /<p:timing>/);
  assert.match(initialSlide, /preserve="unknown-ooxml"/);

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
  assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);

  const originalZip = await extractZip(toArrayBuffer(input));
  const exportedZip = await extractZip(output);
  for (const path of [
    "ppt/theme/theme1.xml",
    "ppt/notesSlides/notesSlide1.xml",
    "ppt/charts/chart1.xml",
    "customXml/native-powerpoint-extension.xml",
  ]) {
    assert.equal(exportedZip.textFiles.get(path), originalZip.textFiles.get(path), `${path} changed`);
  }
  assert.deepEqual(exportedZip.binaryFiles.get("ppt/media/image1.png"), originalZip.binaryFiles.get("ppt/media/image1.png"));

  const exportedSlide = exportedZip.textFiles.get("ppt/slides/slide1.xml");
  const exportedRelationships = exportedZip.textFiles.get("ppt/slides/_rels/slide1.xml.rels");
  assert.match(exportedSlide, /<c:chart\b/);
  assert.match(exportedSlide, /<a:tbl>/);
  assert.match(exportedSlide, /<p:grpSp>/);
  assert.match(exportedSlide, /<p:timing>/);
  assert.match(exportedSlide, /preserve="unknown-ooxml"/);
  assert.ok(exportedRelationships.includes('Target="https://example.com/native-powerpoint"'));
  assert.match(exportedRelationships, /TargetMode="External"/);
});

test("rendered layout backgrounds resolve image rels from the layout part", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(await createLayoutBackgroundRelationshipCollisionFixture());
  const svg = engine.renderSlide(0).svg;
  const backgroundMatch = svg.match(
    /<image\b(?=[^>]*\bx="0")(?=[^>]*\by="0")(?=[^>]*\bpreserveAspectRatio="none")[^>]*\bhref="([^"]+)"/
  );

  assert.ok(backgroundMatch, "expected a full-slide background image in the rendered SVG");
  assert.equal(
    backgroundMatch[1],
    `data:image/png;base64,${onePixelLayoutBackground.toString("base64")}`
  );
});

test("rendered cropped layout backgrounds reconcile href without duplicating full-bleed image", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(await createCroppedLayoutBackgroundFixture());
  const svg = engine.renderSlide(0).svg;

  const fullBleedMatches = svg.match(
    /<image\b(?=[^>]*\bx="0")(?=[^>]*\by="0")(?=[^>]*\bwidth="960")(?=[^>]*\bheight="540")(?=[^>]*\bpreserveAspectRatio="none")[^>]*>/g
  ) ?? [];
  assert.equal(fullBleedMatches.length, 0, "cropped backgrounds should not also emit a full-bleed image");

  assert.match(svg, /clip-path="url\(#bgclip-s1\)"/);
  const expectedHref = `data:image/png;base64,${onePixelLayoutBackground.toString("base64")}`;
  assert.ok(svg.includes(`href="${expectedHref}"`), "expected reconciled cropped background href");
});

test("text formatting keeps authored picture crops through save and reload", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(await createCroppedPictureFixture());

  assert.deepEqual(engine.getImageCrop(0, 2), {
    left: 0,
    top: 23.37,
    right: 0,
    bottom: 23.37,
  });
  await engine.setRunStyle(0, 0, { paragraphIndex: 0, runIndex: 0 }, { fontSizePt: 27 });

  const exported = await engine.export();
  const exportedSlideXml = (await extractZip(exported)).textFiles.get("ppt/slides/slide1.xml");
  assert.ok(exportedSlideXml);
  assert.match(exportedSlideXml, /<a:srcRect l="0" t="23370" r="0" b="23370"\/>/);

  const reloaded = await PresentationEngine.load(exported);
  assert.deepEqual(reloaded.getImageCrop(0, 2), {
    left: 0,
    top: 23.37,
    right: 0,
    bottom: 23.37,
  });
});

test("content validation blocks lossy renderer rewrites of opaque slide markup", async () => {
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const input = await readDeck("features.pptx");
  const renderer = await createRenderer(input);

  assert.doesNotMatch(renderer.updateShapeText(0, 0, 0, 0, "Edited fixture title"), /^ERROR:/);
  const validation = await validatePowerPointExportContents(toArrayBuffer(input), await renderer.exportPptx());
  assert.equal(validation.ok, false);
  assert.ok(validation.errors.some((error) => error.includes("chart")));
  assert.ok(validation.errors.some((error) => error.includes("unknown OOXML element <np:feature>")));
});

test("macro-enabled fixtures preserve VBA bytes during a renderer round trip", async (t) => {
  const { inspectPowerPointPackage, validatePowerPointExport } = await loadPowerPointPackageModule();

  for (const name of macroFixtures) {
    await t.test(name, async () => {
      const input = await readDeck(name);
      const original = inspectPowerPointPackage(toArrayBuffer(input));
      assert.equal(original.hasVbaProject, true);

      const renderer = await createRenderer(input);
      const output = await renderer.exportPptx();
      const exported = inspectPowerPointPackage(output);
      assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);

      const originalZip = await extractZip(toArrayBuffer(input));
      const exportedZip = await extractZip(output);
      assert.deepEqual(exportedZip.binaryFiles.get("ppt/vbaProject.bin"), originalZip.binaryFiles.get("ppt/vbaProject.bin"));
    });
  }
});

test("malformed ZIP fixtures are rejected without crashing", async () => {
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } = await loadPowerPointPackageModule();

  for (const name of ["malformed-random.pptx", "malformed-truncated.pptx"]) {
    const input = await readDeck(name);
    assert.throws(() => inspectPowerPointPackage(toArrayBuffer(input)), /ZIP|Open XML/);
  }

  const unsafe = inspectPowerPointPackage(toArrayBuffer(await readDeck("malformed-unsafe-path.pptx")));
  assert.equal(validatePowerPointPackageStructure(unsafe).ok, false);
  assert.deepEqual(unsafe.unsafePaths, ["../escape.xml"]);

  const duplicate = inspectPowerPointPackage(toArrayBuffer(await readDeck("malformed-duplicate-entry.pptx")));
  assert.equal(validatePowerPointPackageStructure(duplicate).ok, false);
  assert.deepEqual(duplicate.duplicateEntries, ["ppt/presentation.xml"]);
});

test("rapid text edits export one valid final presentation", async () => {
  const {
    inspectPowerPointPackage,
    validatePowerPointExport,
    validatePowerPointExportContents,
  } = await loadPowerPointPackageModule();
  const input = await readDeck("simple-edit.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  const renderer = await createRenderer(input);

  for (let index = 0; index < 40; index += 1) {
    assert.doesNotMatch(renderer.updateShapeText(0, 0, 0, 0, `Rapid edit ${index}`), /^ERROR:/);
  }

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 1).errors, []);
  assert.deepEqual((await validatePowerPointExportContents(toArrayBuffer(input), output)).errors, []);

  const reloaded = await createRenderer(new Uint8Array(output));
  assert.match(reloaded.getSlideOoxml(0), /Rapid edit 39/);
});

test("pasted and duplicated shapes receive fresh a16:creationId GUIDs", async () => {
  const { createSlideObjectClipboard, pasteSlideObject } = await loadShapeClipboardModule();
  const input = await readDeck("simple-edit.pptx");
  const inputBuffer = toArrayBuffer(input);

  // Seed a creationId on shape 0 (cNvPr id="2") so paste/duplicate must regenerate it.
  const slidePath = "ppt/slides/slide1.xml";
  const sourceZip = await extractZip(inputBuffer);
  const seedGuid = "{11111111-1111-1111-1111-111111111111}";
  const creationExt =
    '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${seedGuid}"/>` +
    "</a:ext></a:extLst>";
  const slideXml = sourceZip.textFiles
    .get(slidePath)
    .replace('<p:cNvPr id="2" name="Slide 1 title"/>', `<p:cNvPr id="2" name="Slide 1 title">${creationExt}</p:cNvPr>`);
  assert.match(slideXml, /id="\{11111111-1111-1111-1111-111111111111\}"/, "seed creationId was injected onto shape 0");
  const seeded = await buildZip(inputBuffer, new Map([[slidePath, slideXml]]));

  // Paste the seeded shape twice and "duplicate" it (paste back onto the same slide).
  const clipboard = createSlideObjectClipboard(seeded, 0, 0);
  const firstPaste = await pasteSlideObject(seeded, clipboard, 0);
  const secondPaste = await pasteSlideObject(firstPaste.buffer, clipboard, 0);
  const duplicate = await pasteSlideObject(secondPaste.buffer, clipboard, 0);

  const exportedZip = await extractZip(duplicate.buffer);
  const guids = [];
  for (const [partPath, contents] of exportedZip.textFiles) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(partPath)) continue;
    for (const match of contents.matchAll(/<a16:creationId\b[^>]*\bid="([^"]+)"/g)) {
      guids.push(match[1]);
    }
  }

  assert.equal(guids.length, 4, `expected the seed plus three clones to each carry a creationId, found ${guids.length}`);
  assert.equal(new Set(guids).size, guids.length, "every a16:creationId GUID must remain unique");
  assert.equal(guids.filter((guid) => guid === seedGuid).length, 1, "only the source shape keeps the seed GUID");
});

test("shape transforms preserve authored a16:creationId extensions", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const input = await readDeck("simple-edit.pptx");
  const inputBuffer = toArrayBuffer(input);
  const slidePath = "ppt/slides/slide1.xml";
  const seedGuid = "{22222222-2222-2222-2222-222222222222}";
  const creationExt =
    '<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${seedGuid}"/>` +
    "</a:ext></a:extLst>";
  const sourceZip = await extractZip(inputBuffer);
  const sourceSlide = sourceZip.textFiles.get(slidePath);
  assert.ok(sourceSlide);
  const seeded = await buildZip(
    inputBuffer,
    new Map([[
      slidePath,
      sourceSlide.replace('<p:cNvPr id="2" name="Slide 1 title"/>', `<p:cNvPr id="2" name="Slide 1 title">${creationExt}</p:cNvPr>`),
    ]]),
  );

  const engine = await PresentationEngine.load(seeded);
  await engine.updateShapeTransform(0, 0, { x: 120000, y: 120000, cx: 4000000, cy: 1000000, rot: 0 });
  const exported = await engine.export();
  const validation = await validatePowerPointExportContents(seeded, exported);
  const exportedZip = await extractZip(exported);

  assert.equal(validation.ok, true, validation.errors.join(" "));
  assert.match(exportedZip.textFiles.get(slidePath) ?? "", new RegExp(`a16:creationId[^>]+id="${seedGuid}"`));
});

test("connector extension lists survive renderer conversion to ordinary shapes", async () => {
  const { preserveSlideExtensionLists } = await loadSlideExtensionPreserveModule();
  const input = await readDeck("simple-edit.pptx");
  const inputBuffer = toArrayBuffer(input);
  const slidePath = "ppt/slides/slide1.xml";
  const seedGuid = "{33333333-3333-3333-3333-333333333333}";
  const sourceSlide = [
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main">',
    '<p:cSld><p:spTree><p:cxnSp><p:nvCxnSpPr>',
    '<p:cNvPr id="9" name="Connector"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">',
    `<a16:creationId id="${seedGuid}"/>`,
    '</a:ext></a:extLst></p:cNvPr><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>',
    '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
    '</p:cxnSp></p:spTree></p:cSld></p:sld>',
  ].join("");
  const renderedSlide = [
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>',
    '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>',
    '</p:sp></p:spTree></p:cSld></p:sld>',
  ].join("");
  const authoritative = await buildZip(inputBuffer, new Map([[slidePath, sourceSlide]]));
  const rendered = await buildZip(authoritative, new Map([[slidePath, renderedSlide]]));
  const preserved = await preserveSlideExtensionLists(authoritative, rendered);
  const zip = await extractZip(preserved);

  assert.match(zip.textFiles.get(slidePath) ?? "", new RegExp(`a16:creationId[^>]+id="${seedGuid}"`));
});

test("setRunStyleForRange formats only the selected characters within a paragraph", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  const shapeIndex = 0;
  const paragraphIndex = 0;
  await engine.updateParagraphText(0, shapeIndex, paragraphIndex, "Hello world");
  await engine.setRunStyle(0, shapeIndex, { paragraphIndex, runIndex: 0 }, {
    bold: false,
    italic: false,
    underline: false,
    fontFamily: "Arial",
    fontSizePt: 28,
    color: null,
    highlight: null
  });
  await engine.setRunStyleForRange(0, shapeIndex, paragraphIndex, 6, 11, {
    bold: true,
    italic: true,
    underline: true,
    fontFamily: "Georgia",
    fontSizePt: 31,
    color: "112233",
    highlight: "FFEEDD"
  });

  const beforeSelection = engine.getRunStyle(0, shapeIndex, paragraphIndex, 0);
  const selected = engine.getRunStyle(0, shapeIndex, paragraphIndex, 1);
  assert.equal(beforeSelection?.bold, false);
  assert.equal(beforeSelection?.italic, false);
  assert.equal(beforeSelection?.underline, false);
  assert.equal(beforeSelection?.fontFamily, "Arial");
  assert.equal(beforeSelection?.fontSizePt, 28);
  assert.equal(beforeSelection?.color, null);
  assert.equal(beforeSelection?.highlight, null);
  assert.equal(selected?.bold, true);
  assert.equal(selected?.italic, true);
  assert.equal(selected?.underline, true);
  assert.equal(selected?.fontFamily, "Georgia");
  assert.equal(selected?.fontSizePt, 31);
  assert.equal(selected?.color, "112233");
  assert.equal(selected?.highlight, "FFEEDD");
  assert.equal(engine.isRangeStyled(0, shapeIndex, paragraphIndex, 6, 11, "bold"), true);
  assert.equal(engine.isRangeStyled(0, shapeIndex, paragraphIndex, 0, 5, "bold"), false);
  assert.equal(
    engine.getRangesFontSizePt(0, shapeIndex, [{ paragraphIndex, start: 6, end: 11 }]),
    31
  );
  assert.equal(
    engine.getRangesFontSizePt(0, shapeIndex, [{ paragraphIndex, start: 0, end: 11 }]),
    null
  );
});

test("setRunStyleForRanges applies every run style across paragraph selections", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const patched = await createTitleParagraphFixture([
    paragraphXml("Alpha beta"),
    paragraphXml("Gamma delta")
  ]);
  const engine = await PresentationEngine.load(patched);

  const ranges = [
    { paragraphIndex: 0, start: 6, end: 10 },
    { paragraphIndex: 1, start: 0, end: 5 }
  ];
  await engine.setRunStyleForRanges(0, 0, ranges, {
    bold: true,
    italic: true,
    underline: true,
    fontFamily: "Georgia",
    fontSizePt: 31,
    color: "112233",
    highlight: "FFEEDD"
  });

  assert.equal(engine.areRangesStyled(0, 0, ranges, "bold"), true);
  assert.equal(engine.isRangeStyled(0, 0, 0, 0, 5, "bold"), false);
  assert.equal(engine.isRangeStyled(0, 0, 1, 6, 11, "bold"), false);
  for (const [paragraphIndex, runIndex] of [[0, 1], [1, 0]]) {
    const style = engine.getRunStyle(0, 0, paragraphIndex, runIndex);
    assert.equal(style?.bold, true);
    assert.equal(style?.italic, true);
    assert.equal(style?.underline, true);
    assert.equal(style?.fontFamily, "Georgia");
    assert.equal(style?.fontSizePt, 31);
    assert.equal(style?.color, "112233");
    assert.equal(style?.highlight, "FFEEDD");
  }
  for (const [paragraphIndex, runIndex] of [[0, 0], [1, 1]]) {
    const style = engine.getRunStyle(0, 0, paragraphIndex, runIndex);
    assert.equal(style?.bold, false);
    assert.equal(style?.italic, false);
    assert.equal(style?.underline, false);
    assert.equal(style?.fontFamily, null);
    assert.equal(style?.fontSizePt, 28);
    assert.equal(style?.color, null);
    assert.equal(style?.highlight, null);
  }
});

test("font-size readers use authoritative OOXML after a lossless mutation", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const initial = await createTitleParagraphFixture([
    paragraphXml("Alpha beta"),
    paragraphXml("Gamma delta"),
  ]);
  const engine = await PresentationEngine.load(initial);
  const renderer = engine.pptxDocument.renderer;
  const staleSlideXml = renderer.getSlideOoxml(0);

  // Reproduce pptx-svg returning its pre-mutation representation while the
  // package document already owns the lossless edited slide XML.
  renderer.getSlideOoxml = () => staleSlideXml;
  const ranges = [
    { paragraphIndex: 0, start: 0, end: "Alpha beta".length },
    { paragraphIndex: 1, start: 0, end: "Gamma delta".length },
  ];

  await engine.setRunStyleForRanges(0, 0, ranges, { fontSizePt: 27 });
  assert.equal(engine.getRunStyle(0, 0, 0, 0)?.fontSizePt, 27);
  assert.equal(engine.getRangesFontSizePt(0, 0, ranges), 27);

  await engine.setRunStyleForRanges(0, 0, ranges, { fontSizePt: 26 });
  assert.equal(engine.getRunStyle(0, 0, 1, 0)?.fontSizePt, 26);
  assert.equal(engine.getRangesFontSizePt(0, 0, ranges), 26);
  assert.match(engine.pptxDocument.getAuthoritativeSlideXml(0), /\bsz="2600"/);
});

test("font-size formatting survives a later paragraph split without restoring shrink-to-fit", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const initial = await createTitleParagraphFixture([
    paragraphXml("Alpha beta"),
    paragraphXml("Gamma delta"),
  ]);
  const initialZip = await extractZip(initial);
  const slidePath = "ppt/slides/slide1.xml";
  const initialSlideXml = initialZip.textFiles.get(slidePath);
  assert.ok(initialSlideXml);
  const withShrinkAutofit = await buildZip(initial, new Map([[
    slidePath,
    initialSlideXml.replace("<a:bodyPr/>", "<a:bodyPr><a:normAutofit/></a:bodyPr>"),
  ]]));
  const engine = await PresentationEngine.load(withShrinkAutofit);

  await engine.setRunStyleForRanges(0, 0, [
    { paragraphIndex: 0, start: 0, end: "Alpha beta".length },
    { paragraphIndex: 1, start: 0, end: "Gamma delta".length },
  ], { fontSizePt: 31 });
  assert.match(engine.pptxDocument.getAuthoritativeSlideXml(0), /<a:noAutofit\s*\/>/);
  // Autosave folds pending OOXML into the package before the next text action.
  await engine.export();
  assert.match(engine.pptxDocument.getAuthoritativeSlideXml(0), /<a:noAutofit\s*\/>/);
  // pptx-svg can omit <a:noAutofit/> when it supplies slide XML for the next
  // mutation. The package document must retain the lossless version instead.
  const renderer = engine.pptxDocument.renderer;
  const getSlideOoxml = renderer.getSlideOoxml.bind(renderer);
  renderer.getSlideOoxml = (slideIndex) => getSlideOoxml(slideIndex)
    .replace("<a:noAutofit/>", "<a:normAutofit/>");
  await engine.splitParagraph(0, 0, 0, 5);

  const exported = await engine.export();
  const exportedSlideXml = (await extractZip(exported)).textFiles.get(slidePath);
  assert.ok(exportedSlideXml);
  assert.match(exportedSlideXml, /<a:noAutofit\s*\/>/);
  assert.doesNotMatch(exportedSlideXml, /<a:normAutofit\b/);
  assert.equal((exportedSlideXml.match(/\bsz="3100"/g) ?? []).length, 3);
});

test("setParagraphAlignmentForRanges aligns only selected paragraphs", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const patched = await createTitleParagraphFixture([
    paragraphXml("Alpha beta"),
    paragraphXml("Gamma delta"),
    paragraphXml("Plain ending")
  ]);
  const engine = await PresentationEngine.load(patched);

  await engine.setParagraphAlignmentForRanges(0, 0, [
    { paragraphIndex: 0, start: 6, end: 10 },
    { paragraphIndex: 1, start: 0, end: 5 }
  ], "ctr");

  assert.equal(engine.getRunStyle(0, 0, 0, 0)?.alignment, "ctr");
  assert.equal(engine.getRunStyle(0, 0, 1, 0)?.alignment, "ctr");
  assert.equal(engine.getRunStyle(0, 0, 2, 0)?.alignment, null);
});

test("updateParagraphText preserves line breaks within a paragraph", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await engine.updateParagraphText(0, 0, 0, "Line one\nLine two");

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.match(slideXml, /<a:br/);
  assert.match(slideXml, /Line one/);
  assert.match(slideXml, /Line two/);

  const reloaded = await PresentationEngine.load(exported);
  const svg = reloaded.renderSlide(0).svg;
  assert.ok(svg.includes(">one<"));
  assert.ok(svg.includes(">two<"));
  assert.ok((svg.match(/data-ooxml-para-idx="0"/g) || []).length >= 2);
});

test("splitParagraph creates a native sibling paragraph and preserves list and run styling", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const bulletParagraph = [
    '<a:p><a:pPr marL="285750" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>',
    '<a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>Bold </a:t></a:r>',
    '<a:br/><a:r><a:rPr lang="en-US" sz="2800" b="0"/><a:t>suffix</a:t></a:r>',
    '<a:endParaRPr lang="en-US"/></a:p>',
  ].join("");
  const patched = await createTitleParagraphFixture([bulletParagraph, paragraphXml("Trailing paragraph")]);
  const engine = await PresentationEngine.load(patched);

  const result = await engine.splitParagraph(0, 0, 0, 5);
  assert.deepEqual(result, {
    paragraphIndex: 1,
    beforeParagraphCount: 2,
    afterParagraphCount: 3,
    listStyle: "bullet",
    removedSoftBreaks: 1,
  });
  assert.equal(engine.getParagraphRunText(0, 0, 0), "Bold ");
  assert.equal(engine.getParagraphRunText(0, 0, 1), "suffix");
  assert.equal(engine.getParagraphRunText(0, 0, 2), "Trailing paragraph");
  assert.equal(engine.getParagraphListStyle(0, 0, 0), "bullet");
  assert.equal(engine.getParagraphListStyle(0, 0, 1), "bullet");
  assert.equal(engine.getRunStyle(0, 0, 0, 0)?.bold, true);
  assert.equal(engine.getRunStyle(0, 0, 1, 0)?.bold, false);

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.doesNotMatch(slideXml, /<a:br\b/);

  const reloaded = await PresentationEngine.load(exported);
  assert.equal(reloaded.getParagraphRunText(0, 0, 0), "Bold ");
  assert.equal(reloaded.getParagraphRunText(0, 0, 1), "suffix");
  assert.equal(reloaded.getParagraphRunText(0, 0, 2), "Trailing paragraph");
  assert.equal(reloaded.getParagraphListStyle(0, 0, 1), "bullet");
  assert.match(reloaded.renderSlide(0).svg, /^<svg\b/);
});

test("applyListStyleForRange isolates selected text as one native bullet paragraph", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const styledParagraph = [
    '<a:p><a:pPr algn="l" marL="0" indent="-285750"/><a:r><a:rPr lang="en-US" sz="2800"/><a:t>Prefix </a:t></a:r>',
    '<a:r><a:rPr lang="en-US" sz="2800" b="1"/><a:t>selected</a:t></a:r>',
    '<a:r><a:rPr lang="en-US" sz="2800" i="1"/><a:t> suffix</a:t></a:r>',
    '<a:endParaRPr lang="en-US"/></a:p>',
  ].join("");
  const patched = await createTitleParagraphFixture([styledParagraph, paragraphXml("Trailing paragraph")]);
  const engine = await PresentationEngine.load(patched);

  const result = await engine.applyListStyleForRange(0, 0, {
    paragraphIndex: 0,
    start: 7,
    end: 15,
  }, "bullet");
  assert.deepEqual(result, {
    changed: true,
    sourceParagraphIndex: 0,
    selectedParagraphIndex: 1,
    selectedRange: { paragraphIndex: 1, start: 0, end: 8 },
    beforeParagraphCount: 2,
    afterParagraphCount: 4,
    splitPrefix: true,
    splitSuffix: true,
  });
  assert.equal(engine.getParagraphRunText(0, 0, 0), "Prefix ");
  assert.equal(engine.getParagraphRunText(0, 0, 1), "selected");
  assert.equal(engine.getParagraphRunText(0, 0, 2), " suffix");
  assert.equal(engine.getParagraphRunText(0, 0, 3), "Trailing paragraph");
  assert.equal(engine.getParagraphListStyle(0, 0, 0), null);
  assert.equal(engine.getParagraphListStyle(0, 0, 1), "bullet");
  assert.equal(engine.getParagraphListStyle(0, 0, 2), null);
  assert.equal(engine.getRunStyle(0, 0, 1, 0)?.bold, true);
  assert.equal(engine.getRunStyle(0, 0, 2, 0)?.italic, true);

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.match(
    slideXml,
    /<a:pPr\b[^>]*\bmarL="285750"[^>]*\bindent="-285750"[^>]*><a:buFont\b[^>]*\btypeface="Arial"\/><a:buChar\b[^>]*\bchar="•"/,
    "selected text replaces inherited zero-margin geometry with a hanging list indent",
  );
  assert.doesNotMatch(slideXml, /<a:t>[^<]*•/);

  const reloaded = await PresentationEngine.load(exported);
  assert.equal(reloaded.getParagraphRunText(0, 0, 1), "selected");
  assert.equal(reloaded.getParagraphListStyle(0, 0, 1), "bullet");
  assert.equal(reloaded.getRunStyle(0, 0, 1, 0)?.bold, true);
  assert.equal(reloaded.getRunStyle(0, 0, 2, 0)?.italic, true);
  const rendered = reloaded.renderSlide(0).svg;
  assert.match(rendered, /^<svg\b/);
  const renderedDocument = new globalThis.DOMParser().parseFromString(rendered, "text/xml");
  const selectedParagraphContainers = Array.from(renderedDocument.getElementsByTagName("tspan")).filter(
    (element) => element.getAttribute("data-ooxml-para-idx") === "1",
  );
  const hasRun = (element) => Array.from(element.getElementsByTagName("tspan")).some(
    (child) => child.getAttribute("data-ooxml-run-idx") !== null,
  );
  const markerContainer = selectedParagraphContainers.find(
    (element) => (element.textContent || "").trim().startsWith("•") && !hasRun(element),
  );
  const runContainer = selectedParagraphContainers.find((element) => hasRun(element));
  assert.ok(markerContainer, "selected bullet renders in a marker-only container");
  assert.ok(runContainer, "selected text renders after the marker");
  assert.ok(
    Number(runContainer.getAttribute("x")) > Number(markerContainer.getAttribute("x")),
    "selected run starts to the right of its bullet marker",
  );
});

test("applyListStyleForRanges toggles every selected bullet paragraph in one mutation", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const texts = ["First line", "Second line", "Third line"];
  const patched = await createTitleParagraphFixture(texts.map((text) => paragraphXml(text)));
  const engine = await PresentationEngine.load(patched);
  const ranges = texts.map((text, paragraphIndex) => ({
    paragraphIndex,
    start: 0,
    end: text.length,
  }));

  await engine.applyListStyleForRanges(0, 0, ranges, "bullet");
  for (let paragraphIndex = 0; paragraphIndex < texts.length; paragraphIndex += 1) {
    assert.equal(engine.getParagraphListStyle(0, 0, paragraphIndex), "bullet");
    assert.equal(engine.getParagraphRunText(0, 0, paragraphIndex), texts[paragraphIndex]);
  }

  await engine.applyListStyleForRanges(0, 0, ranges, "none");
  for (let paragraphIndex = 0; paragraphIndex < texts.length; paragraphIndex += 1) {
    assert.equal(engine.getParagraphListStyle(0, 0, paragraphIndex), "none");
    assert.equal(engine.getParagraphRunText(0, 0, paragraphIndex), texts[paragraphIndex]);
  }

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.equal((slideXml.match(/<a:buChar\b/g) || []).length, 0,
    "removing every selected native bullet must remove all bullet markers from OOXML");
  assert.equal((slideXml.match(/<a:buNone\b/g) || []).length, 3,
    "removing every selected native bullet must explicitly suppress inherited markers");

  const reloaded = await PresentationEngine.load(exported);
  for (let paragraphIndex = 0; paragraphIndex < texts.length; paragraphIndex += 1) {
    assert.equal(reloaded.getParagraphListStyle(0, 0, paragraphIndex), "none");
  }
  const renderedDocument = new globalThis.DOMParser().parseFromString(reloaded.renderSlide(0).svg, "text/xml");
  const targetShape = Array.from(renderedDocument.getElementsByTagName("g")).find(
    (element) => element.getAttribute("data-ooxml-shape-idx") === "0",
  );
  assert.ok(targetShape, "expected the edited text shape in the rendered slide");
  assert.doesNotMatch(targetShape.textContent || "", /•/,
    "removing every selected native bullet must remove all rendered bullet markers");
});

test("applyListStyleForRanges preserves source indices while partial ranges split paragraphs", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const patched = await createTitleParagraphFixture([
    paragraphXml("First target"),
    paragraphXml("Second target"),
  ]);
  const engine = await PresentationEngine.load(patched);

  await engine.applyListStyleForRanges(0, 0, [
    { paragraphIndex: 0, start: 6, end: 12 },
    { paragraphIndex: 1, start: 7, end: 13 },
  ], "bullet");

  assert.deepEqual([0, 1, 2, 3].map((paragraphIndex) => (
    engine.getParagraphRunText(0, 0, paragraphIndex)
  )), ["First ", "target", "Second ", "target"]);
  assert.deepEqual([0, 1, 2, 3].map((paragraphIndex) => (
    engine.getParagraphListStyle(0, 0, paragraphIndex)
  )), [null, "bullet", null, "bullet"]);
});

test("list toggle converts an imported literal bullet without leaving a duplicate marker", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const literalBulletParagraph = [
    '<a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="2800"/><a:t>• Imported list item</a:t></a:r>',
    '<a:endParaRPr lang="en-US"/></a:p>',
  ].join("");
  const patched = await createTitleParagraphFixture([literalBulletParagraph]);
  const engine = await PresentationEngine.load(patched);

  await engine.applyListStyle(0, 0, 0, "bullet", true);
  assert.equal(engine.getParagraphListStyle(0, 0, 0), "bullet");
  assert.equal(engine.getParagraphRunText(0, 0, 0), "Imported list item");

  await engine.applyListStyle(0, 0, 0, "none", true);
  assert.equal(engine.getParagraphListStyle(0, 0, 0), "none");
  assert.equal(engine.getParagraphRunText(0, 0, 0), "Imported list item");
});

test("replaceShapeParagraphs writes native list paragraphs and survives a round trip", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  await engine.replaceShapeParagraphs(0, 0, [
    { text: "Fluent AI-Assistance can:", listStyle: "none" },
    { text: "Inspect logs and scripts", listStyle: "bullet" },
    { text: "Build review-ready artifacts", listStyle: "bullet" },
    { text: "Keep human review in the loop.", listStyle: "none" },
  ]);

  assert.equal(engine.getParagraphRunText(0, 0, 0), "Fluent AI-Assistance can:");
  assert.equal(engine.getParagraphRunText(0, 0, 1), "Inspect logs and scripts");
  assert.equal(engine.getParagraphListStyle(0, 0, 0), "none");
  assert.equal(engine.getParagraphListStyle(0, 0, 1), "bullet");
  assert.equal(engine.getParagraphListStyle(0, 0, 2), "bullet");
  assert.equal(engine.getParagraphListStyle(0, 0, 3), "none");

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.equal((slideXml.match(/<a:buChar\b[^>]*\bchar="•"/g) || []).length, 2);
  assert.doesNotMatch(slideXml, /<a:t>[^<]*•/);

  const reloaded = await PresentationEngine.load(exported);
  assert.equal(reloaded.getParagraphRunText(0, 0, 1), "Inspect logs and scripts");
  assert.equal(reloaded.getParagraphListStyle(0, 0, 1), "bullet");
  assert.match(reloaded.renderSlide(0).svg, /^<svg\b/);
});

test("updateShapeTransform allows shapes outside the slide bounds", async () => {
  const { createRequire } = await import("node:module");
  const JSZip = createRequire(import.meta.url)("jszip");
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("features.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  const fragment = await engine.updateShapeTransform(0, 0, {
    x: -9000000,
    y: 342900,
    cx: 5943600,
    cy: 685800,
    rot: 0
  });

  const exported = await engine.export();
  const zip = await JSZip.loadAsync(exported);
  const slideXml = await zip.files["ppt/slides/slide1.xml"].async("string");
  assert.match(slideXml, /x="-9000000"/);
  assert.match(fragment ?? "", /<g\b[^>]*data-ooxml-shape-idx="0"/);
});

test("slide add, reorder, and delete operations survive export", async () => {
  const { PresentationEngine } = await loadPresentationEngineModule();
  const input = await readDeck("simple-edit.pptx");
  const engine = await PresentationEngine.load(toArrayBuffer(input));

  assert.equal((await engine.addSlide(0)).slideCount, 2);
  assert.equal((await engine.addSlide(1)).slideCount, 3);
  assert.equal((await engine.moveSlide(2, -1)).slideCount, 3);
  assert.equal((await engine.deleteSlide(1)).slideCount, 2);

  const reloaded = await createRenderer(new Uint8Array(await engine.export()));
  assert.equal(reloaded.getSlideCount(), 2);
  assert.match(reloaded.renderSlideSvg(0), /^<svg\b/);
  assert.match(reloaded.renderSlideSvg(1), /^<svg\b/);
});

test("large 160-slide fixture loads, renders its bounds, and round-trips", async () => {
  const { inspectPowerPointPackage, validatePowerPointExport } = await loadPowerPointPackageModule();
  const input = await readDeck("large-deck.pptx");
  const original = inspectPowerPointPackage(toArrayBuffer(input));
  assert.equal(original.slidePaths.length, 160);

  const renderer = await createRenderer(input);
  assert.equal(renderer.getSlideCount(), 160);
  assert.match(renderer.renderSlideSvg(0), /^<svg\b/);
  assert.match(renderer.getSlideOoxml(159), /Large deck slide 160/);

  const output = await renderer.exportPptx();
  const exported = inspectPowerPointPackage(output);
  assert.deepEqual(validatePowerPointExport(original, exported, 160).errors, []);

  const reloaded = await createRenderer(new Uint8Array(output));
  assert.equal(reloaded.getSlideCount(), 160);
  assert.match(reloaded.getSlideOoxml(159), /Large deck slide 160/);
});
