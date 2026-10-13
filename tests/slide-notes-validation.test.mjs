import test from 'node:test';
import * as assert from 'node:assert';
import { extractZip, buildZip } from 'pptx-svg';
import { readDeck, toArrayBuffer } from './helpers/renderer.mjs';
import { bundleSource } from './helpers/load-plugin-modules.mjs';
import { pathToFileURL } from 'node:url';

test('MUTABLE_PART_PATTERNS correctly allows notes and theme parts to be removed or modified', async () => {
  const outfile = await bundleSource('src/PowerPointPackage.ts', 'powerpoint-package.cjs');
  const mod = await import(pathToFileURL(outfile).href);

  const originalBytes = await readDeck('features.pptx');
  const originalBuffer = toArrayBuffer(originalBytes);
  const originalInspection = mod.inspectPowerPointPackage(originalBuffer);

  const zip = await extractZip(originalBuffer);

  const originalNotes = zip.textFiles.get('ppt/notesSlides/notesSlide1.xml');
  const originalNotesMaster = zip.textFiles.get('ppt/notesMasters/notesMaster1.xml');
  const originalTheme = zip.textFiles.get('ppt/theme/theme1.xml');

  const exportedBuffer = await buildZip(originalBuffer, new Map([
    ['ppt/notesSlides/notesSlide1.xml', originalNotes + "<!-- test mutation -->"],
    ['ppt/notesMasters/notesMaster1.xml', originalNotesMaster + "<!-- test mutation -->"],
    ['ppt/theme/theme1.xml', originalTheme + "<!-- test mutation -->"]
  ]));

  const exportedInspection = mod.inspectPowerPointPackage(exportedBuffer);

  const result = mod.validatePowerPointExport(originalInspection, exportedInspection, 1);
  assert.ok(result.ok, `Expected export to be valid, but got errors: ${result.errors.join(', ')}`);
  assert.deepEqual(result.errors, []);
});
