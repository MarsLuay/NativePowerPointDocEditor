import assert from 'node:assert/strict';
import { test } from 'node:test';

import { bundleSource } from './helpers/load-plugin-modules.mjs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const source = await bundleSource('src/powerpoint/slideLayouts.ts', 'slide-layouts.cjs');
const { listSlideLayouts } = require(source);

const rels = (body) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const layout = ({ name, type, matchingName = '' }) => `<?xml version="1.0"?>
  <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="${type}" matchingName="${matchingName}">
    <p:cSld name="${name}"/>
  </p:sldLayout>`;

test('lists each attached custom layout once and keeps its insertion source slide', () => {
  const contents = {
    textFiles: new Map([
      ['ppt/slides/_rels/slide1.xml.rels', rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/cover.xml"/>')],
      ['ppt/slides/_rels/slide2.xml.rels', rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/cover.xml"/>')],
      ['ppt/slides/_rels/slide3.xml.rels', rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/section.xml"/>')],
      ['ppt/slideLayouts/cover.xml', layout({ name: 'Cover fallback', type: 'title', matchingName: 'Cover slide' })],
      ['ppt/slideLayouts/section.xml', layout({ name: 'Section header', type: 'secHead' })],
    ]),
  };

  assert.deepEqual(listSlideLayouts(contents, 3), [
    {
      id: 'ppt/slideLayouts/cover.xml',
      layoutPath: 'ppt/slideLayouts/cover.xml',
      name: 'Cover slide',
      type: 'title',
      representativeSlideIndex: 0,
    },
    {
      id: 'ppt/slideLayouts/section.xml',
      layoutPath: 'ppt/slideLayouts/section.xml',
      name: 'Section header',
      type: 'secHead',
      representativeSlideIndex: 2,
    },
  ]);
});

test('ignores missing and external slide-layout relationships', () => {
  const contents = {
    textFiles: new Map([
      ['ppt/slides/_rels/slide1.xml.rels', rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="https://example.invalid/layout.xml" TargetMode="External"/>')],
      ['ppt/slides/_rels/slide2.xml.rels', rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/missing.xml"/>')],
    ]),
  };

  assert.deepEqual(listSlideLayouts(contents, 2), []);
});
