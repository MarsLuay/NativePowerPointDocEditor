import assert from "node:assert/strict";
import { test } from "node:test";
import { loadSlideInsertionsModule } from "./helpers/load-plugin-modules.mjs";

test("buildDuplicateSlideOrder builds correct order array", async () => {
  const { buildDuplicateSlideOrder } = await loadSlideInsertionsModule();

  // slideCount, sourceIndex, insertedIdx
  // e.g., slides 0, 1, 2. We duplicate 1, inserting it at index 2.
  // The slides were [0, 1, 2], so the new order should refer to old slides: [0, 1, 1, 2] -> 0 is old 0, 1 is old 1, 2 is old 1, 3 is old 2.
  assert.deepEqual(buildDuplicateSlideOrder(4, 1, 2), [0, 1, 1, 2]);

  // Duplicating the first slide (index 0) and inserting at index 1
  assert.deepEqual(buildDuplicateSlideOrder(2, 0, 1), [0, 0]);

  // Duplicating the last slide in a 3-slide deck (source=2, inserted at 3), total count will be 4
  assert.deepEqual(buildDuplicateSlideOrder(4, 2, 3), [0, 1, 2, 2]);
});

test("countSlideTopLevelShapes returns correct number of top level shapes", async () => {
  const { countSlideTopLevelShapes } = await loadSlideInsertionsModule();

  const xml0 = `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
        </p:spTree>
      </p:cSld>
    </p:sld>
  `;
  assert.equal(countSlideTopLevelShapes(xml0, "ppt/slides/slide1.xml"), 0);

  const xml2 = `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 1"/></p:nvSpPr></p:sp>
          <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Shape 2"/></p:nvCxnSpPr></p:cxnSp>
        </p:spTree>
      </p:cSld>
    </p:sld>
  `;
  assert.equal(countSlideTopLevelShapes(xml2, "ppt/slides/slide2.xml"), 2);

  const xmlWithGroups = `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr/>
          <p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 1"/></p:nvSpPr></p:sp>
          <p:grpSp>
             <p:nvGrpSpPr><p:cNvPr id="3" name="Group 1"/></p:nvGrpSpPr>
             <p:sp><p:nvSpPr><p:cNvPr id="4" name="Inner Shape 1"/></p:nvSpPr></p:sp>
          </p:grpSp>
          <p:pic><p:nvPicPr><p:cNvPr id="5" name="Pic 1"/></p:nvPicPr></p:pic>
        </p:spTree>
      </p:cSld>
    </p:sld>
  `;
  // Expected to count <p:sp>, <p:grpSp>, and <p:pic> as top-level children of the spTree
  assert.equal(countSlideTopLevelShapes(xmlWithGroups, "ppt/slides/slide3.xml"), 3);
});
