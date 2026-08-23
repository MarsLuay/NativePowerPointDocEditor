import assert from "node:assert/strict";
import { extractZip } from "pptx-svg";
import test from "node:test";
import { loadEngine } from "./helpers/pptx-action-harness.mjs";
import { loadPowerPointPackageModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

test("deleting an image prunes its media part and keeps save validation green", async () => {
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const source = toArrayBuffer(await readDeck("features.pptx"));
  const beforeZip = await extractZip(source);
  const mediaBefore = [...beforeZip.binaryFiles.keys()].filter((path) => path.startsWith("ppt/media/"));
  assert.ok(mediaBefore.length > 0, "fixture must embed media");

  const engine = await loadEngine("features.pptx");
  engine.renderSlide(0);
  let imageIdx = -1;
  for (let index = 0; index < 32; index += 1) {
    if (engine.isImageShape(0, index)) {
      imageIdx = index;
      break;
    }
  }
  assert.ok(imageIdx >= 0);

  await engine.deleteShape(0, imageIdx);
  const exported = await engine.export();
  const afterZip = await extractZip(exported);
  const mediaAfter = [...afterZip.binaryFiles.keys()].filter((path) => path.startsWith("ppt/media/"));
  assert.ok(
    mediaAfter.length < mediaBefore.length,
    `expected media prune: before=${mediaBefore.length} after=${mediaAfter.length}`,
  );

  const allowed = await validatePowerPointExportContents(source, exported, {
    allowedMarkerRemovals: engine.getProtectedSlideMarkerRemovalAllowance(),
    allowedPartRemovals: engine.getPrunedPackageParts(),
  });
  assert.equal(allowed.ok, true, JSON.stringify(allowed.errors));
});

test("deleting a chart prunes chart and embedding parts", async () => {
  const engine = await loadEngine("features.pptx");
  engine.renderSlide(0);
  let chartIdx = -1;
  for (let index = 0; index < 32; index += 1) {
    try {
      const svg = engine.renderShape(0, index);
      if (svg?.includes('data-ooxml-shape-type="chart"')) {
        chartIdx = index;
        break;
      }
    } catch {
      break;
    }
  }
  assert.ok(chartIdx >= 0, "features.pptx must include a chart");

  const before = await extractZip(await engine.export());
  const chartsBefore = [...before.textFiles.keys()].filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path));
  assert.ok(chartsBefore.length > 0);

  await engine.deleteShape(0, chartIdx);
  const after = await extractZip(await engine.export());
  const chartsAfter = [...after.textFiles.keys()].filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path));
  assert.ok(
    chartsAfter.length < chartsBefore.length,
    `expected chart prune: before=${chartsBefore.length} after=${chartsAfter.length}`,
  );
});

test("deleting a shape with an external relationship validates successfully", async () => {
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const source = toArrayBuffer(await readDeck("features.pptx"));

  const engine = await loadEngine("features.pptx");
  engine.renderSlide(0);
  const titleSvg = engine.renderShape(0, 0);
  assert.match(
    String(titleSvg).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(),
    /Native PowerPoint fixture/,
    "features.pptx shape 0 must be the hyperlink title",
  );

  await engine.deleteShape(0, 0);
  const exported = await engine.export();

  const unallowed = await validatePowerPointExportContents(source, exported);
  assert.equal(unallowed.ok, false);
  assert.ok(
    unallowed.errors.some((error) => error.includes("external relationship")),
    `expected an external-relationship drop without an allowance, got ${JSON.stringify(unallowed.errors)}`,
  );

  const allowance = engine.getExternalRelationshipRemovalAllowance();
  assert.ok(
    (allowance["https://example.com/native-powerpoint"] ?? 0) > 0,
    `expected hyperlink target allowance, got ${JSON.stringify(allowance)}`,
  );

  const withoutExternalAllowance = await validatePowerPointExportContents(source, exported, {
    allowedMarkerRemovals: engine.getProtectedSlideMarkerRemovalAllowance(),
    allowedPartRemovals: engine.getPrunedPackageParts(),
    allowedUnknownElementRemovals: engine.getUnknownSlideElementRemovalAllowance(),
  });
  assert.equal(withoutExternalAllowance.ok, false);
  assert.ok(
    withoutExternalAllowance.errors.some((error) => error.includes("external relationship")),
    `save-path options without external-relationship allowance should still fail, got ${JSON.stringify(withoutExternalAllowance.errors)}`,
  );

  const allowed = await validatePowerPointExportContents(source, exported, {
    allowedMarkerRemovals: engine.getProtectedSlideMarkerRemovalAllowance(),
    allowedPartRemovals: engine.getPrunedPackageParts(),
    allowedUnknownElementRemovals: engine.getUnknownSlideElementRemovalAllowance(),
    allowedExternalRelationshipRemovals: allowance,
  });

  assert.equal(allowed.ok, true, `Validation failed: ${allowed.errors.join(", ")}`);
});
