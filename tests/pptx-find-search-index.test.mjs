import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test from 'node:test';
import { DOMParser } from '@xmldom/xmldom';
import { build } from 'esbuild';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
globalThis.DOMParser = DOMParser;

async function loadSearchIndexModule() {
  const outdir = await mkdtemp(path.join(tmpdir(), 'npde-find-index-test-'));
  const outfile = path.join(outdir, 'find-search-index.cjs');
  await build({
    absWorkingDir: projectRoot,
    entryPoints: ['src/powerpoint/findSearchIndex.ts'],
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile,
    platform: 'node',
    target: 'node22',
  });
  return import(pathToFileURL(outfile).href);
}

test('find search index filters shape text and falls back once per matching slide', async () => {
  const { collectFindMatchesFromSearchIndex } = await loadSearchIndexModule();
  const index = [
    {
      slideIndex: 0,
      shapeMatches: [{ slideIndex: 0, shapeIndex: 4, text: 'Alpha beta' }],
      fallbackText: 'Alpha beta speaker notes',
    },
    {
      slideIndex: 1,
      shapeMatches: [],
      fallbackText: 'Gamma alpha',
    },
  ];

  assert.deepEqual(collectFindMatchesFromSearchIndex(index, 'ALPHA'), [
    { slideIndex: 0, shapeIndex: 4, text: 'Alpha beta' },
    { slideIndex: 1, shapeIndex: null, text: 'Gamma alpha' },
  ]);
  assert.deepEqual(collectFindMatchesFromSearchIndex(index, 'beta'), [
    { slideIndex: 0, shapeIndex: 4, text: 'Alpha beta' },
  ]);
});

test('find search index extracts shape text from OOXML without rendering SVG', async () => {
  const { createFindSearchIndexSlideFromOoxml } = await loadSearchIndexModule();
  const indexSlide = createFindSearchIndexSlideFromOoxml(2, `
    <p:sld xmlns:p="urn:p" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree>
        <p:sp><p:txBody><a:p><a:r><a:t>Alpha beta</a:t></a:r></a:p></p:txBody></p:sp>
        <p:pic />
      </p:spTree></p:cSld>
    </p:sld>
  `);

  assert.deepEqual(indexSlide.shapeMatches, [
    { slideIndex: 2, shapeIndex: 0, text: 'Alpha beta' },
  ]);
  assert.equal(indexSlide.fallbackText, 'Alpha beta');
});

test('find highlights create a detached SVG rect before inserting it', async () => {
  const source = await readFile(
    path.join(projectRoot, 'src/powerpoint/findReplaceController.ts'),
    'utf8',
  );
  const method = source.slice(
    source.indexOf('private renderFindHighlightRects('),
    source.indexOf('private updateFindStatus()', source.indexOf('private renderFindHighlightRects(')),
  );

  assert.match(method, /ownerDocument\.createElementNS\(SVG_NAMESPACE, 'rect'\)/);
  assert.doesNotMatch(method, /createSvg\('rect'\)/);
  assert.match(method, /parent\.insertBefore\(rect, textElement\)/);
});
