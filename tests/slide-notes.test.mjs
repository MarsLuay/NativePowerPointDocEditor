import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";
import { createDeck } from "./helpers/fixture-builder.mjs";
import { createRenderer, readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const require = createRequire(import.meta.url);
let notesModulePromise;

function loadSlideNotesModule() {
  notesModulePromise ??= bundleSource(
    "src/powerpoint/slideNotes.ts",
    "slide-notes.cjs",
  ).then((outfile) => require(outfile));
  return notesModulePromise;
}

test("speaker notes text updates preserve the rest of an existing notes part and its relationships", async () => {
  const { readSlideNotesText, writeSlideNotesText } = await loadSlideNotesModule();
  const input = toArrayBuffer(await readDeck("features.potx"));
  const inputZip = await extractZip(input);
  const notesPath = "ppt/notesSlides/notesSlide1.xml";
  const notesRelsPath = "ppt/notesSlides/_rels/notesSlide1.xml.rels";
  const originalNotes = inputZip.textFiles.get(notesPath);
  const originalNotesRels = inputZip.textFiles.get(notesRelsPath);
  const originalMaster = inputZip.textFiles.get("ppt/notesMasters/notesMaster1.xml");
  assert.ok(originalNotes);
  assert.ok(originalNotesRels);
  assert.ok(originalMaster);

  const seeded = await buildZip(input, new Map([[
    notesPath,
    originalNotes
      .replace('<p:cNvPr id="2" name="Speaker notes"/>', '<p:cNvPr id="2" name="Speaker notes" descr="keep-this-note-shape"/>')
      .replace('</p:notes>', '<p:extLst><p:ext uri="{A4A7D03E-44D5-4AA7-BE76-A29394A9C746}"/></p:extLst></p:notes>'),
  ]]));

  assert.deepEqual(await readSlideNotesText(seeded, 0), {
    text: "Fixture speaker notes survive round trip.",
    notesSlidePath: notesPath,
  });

  const written = await writeSlideNotesText(seeded, 0, "First line\nSecond line");
  assert.equal(written.createdNotesSlide, false);
  assert.equal(written.createdNotesMaster, false);
  assert.equal(written.notesSlidePath, notesPath);
  assert.deepEqual(await readSlideNotesText(written.buffer, 0), {
    text: "First line\nSecond line",
    notesSlidePath: notesPath,
  });

  const outputZip = await extractZip(written.buffer);
  const outputNotes = outputZip.textFiles.get(notesPath);
  assert.ok(outputNotes);
  assert.match(outputNotes, /descr="keep-this-note-shape"/);
  assert.match(outputNotes, /<p:ext uri="\{A4A7D03E-44D5-4AA7-BE76-A29394A9C746\}"\/>/);
  assert.match(outputNotes, /<a:rPr lang="en-US"\/>/);
  assert.equal(outputZip.textFiles.get(notesRelsPath), originalNotesRels);
  assert.equal(outputZip.textFiles.get("ppt/notesMasters/notesMaster1.xml"), originalMaster);
});

test("speaker notes creation adds a notes slide and preserves the existing notes master", async () => {
  const { readSlideNotesText, writeSlideNotesText } = await loadSlideNotesModule();
  const input = toArrayBuffer(await createDeck({ format: "potx", slideCount: 2 }));
  const written = await writeSlideNotesText(input, 1, "Second slide presenter note");

  assert.equal(written.createdNotesSlide, true);
  assert.equal(written.createdNotesMaster, false);
  assert.equal(written.notesSlidePath, "ppt/notesSlides/notesSlide2.xml");
  assert.deepEqual(await readSlideNotesText(written.buffer, 1), {
    text: "Second slide presenter note",
    notesSlidePath: "ppt/notesSlides/notesSlide2.xml",
  });

  const zip = await extractZip(written.buffer);
  const slideRels = zip.textFiles.get("ppt/slides/_rels/slide2.xml.rels");
  const notesRels = zip.textFiles.get("ppt/notesSlides/_rels/notesSlide2.xml.rels");
  const contentTypes = zip.textFiles.get("[Content_Types].xml");
  assert.ok(slideRels);
  assert.ok(notesRels);
  assert.ok(contentTypes);
  assert.match(slideRels, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide" Target="\.\.\/notesSlides\/notesSlide2\.xml"/);
  assert.match(notesRels, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesMaster" Target="\.\.\/notesMasters\/notesMaster1\.xml"/);
  assert.match(notesRels, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide" Target="\.\.\/slides\/slide2\.xml"/);
  assert.match(contentTypes, /PartName="\/ppt\/notesSlides\/notesSlide2\.xml"/);

  const renderer = await createRenderer(new Uint8Array(written.buffer));
  assert.deepEqual(renderer.getSlideNotes(1), ["Second slide presenter note"]);
});

test("speaker notes creation supplies a notes master when a modern deck has none", async () => {
  const { writeSlideNotesText } = await loadSlideNotesModule();
  const input = toArrayBuffer(await createDeck({ format: "pptx", slideCount: 2 }));
  const inputZip = await extractZip(input);
  const contentTypes = inputZip.textFiles.get("[Content_Types].xml");
  assert.ok(contentTypes);
  const withoutNotesMaster = await buildZip(
    input,
    new Map([[
      "[Content_Types].xml",
      contentTypes
        .replace(/\s*<Override PartName="\/ppt\/notesMasters\/notesMaster1\.xml"[^>]*\/>/g, "")
        .replace(/\s*<Override PartName="\/ppt\/notesSlides\/notesSlide1\.xml"[^>]*\/>/g, ""),
    ]]),
    new Set([
      "ppt/notesMasters/notesMaster1.xml",
      "ppt/notesMasters/_rels/notesMaster1.xml.rels",
      "ppt/notesSlides/notesSlide1.xml",
      "ppt/notesSlides/_rels/notesSlide1.xml.rels",
    ]),
  );

  const written = await writeSlideNotesText(withoutNotesMaster, 1, "Fresh notes master");
  assert.equal(written.createdNotesSlide, true);
  assert.equal(written.createdNotesMaster, true);

  const zip = await extractZip(written.buffer);
  const presentation = zip.textFiles.get("ppt/presentation.xml");
  const presentationRels = zip.textFiles.get("ppt/_rels/presentation.xml.rels");
  const contentTypesAfter = zip.textFiles.get("[Content_Types].xml");
  assert.ok(presentation);
  assert.ok(presentationRels);
  assert.ok(contentTypesAfter);
  assert.ok(zip.textFiles.has("ppt/notesMasters/notesMaster1.xml"));
  assert.ok(zip.textFiles.has("ppt/notesMasters/_rels/notesMaster1.xml.rels"));
  assert.match(presentation, /<p:notesMasterIdLst>/);
  assert.match(presentationRels, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesMaster" Target="notesMasters\/notesMaster1\.xml"/);
  assert.match(contentTypesAfter, /PartName="\/ppt\/notesMasters\/notesMaster1\.xml"/);
});

test("speaker notes writes reject macro-enabled packages without executing macros", async () => {
  const { writeSlideNotesText } = await loadSlideNotesModule();
  const macro = toArrayBuffer(await readDeck("macro-view-only.pptm"));
  await assert.rejects(
    writeSlideNotesText(macro, 0, "Never execute macros"),
    /PPTX and POTX packages only/,
  );
});
