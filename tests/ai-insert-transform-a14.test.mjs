import assert from "node:assert/strict";
import { access } from "node:fs/promises";
import { test } from "node:test";
import { buildZip, extractZip } from "pptx-svg";
import {
  loadPowerPointPackageModule,
  loadPresentationEngineModule,
} from "./helpers/load-plugin-modules.mjs";
import { getDocxRuntimeAliases } from "./helpers/docx-runtime-aliases.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const FIXTURE = "features.pptx";
const TRANSFORM = { x: 914400, y: 685800, cx: 2286000, cy: 1524000, rot: 0 };
const WORKSHOP_DECK =
  "/Users/mars/ObsidianNotes/Job/PNNL/Abstract Info/Abstract Workshop_CMG_July2026.pptx";
const SPPR_HIDDEN_FILL =
  '<a:extLst><a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">' +
  '<a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">' +
  "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill>" +
  "</a14:hiddenFill></a:ext></a:extLst>";

async function injectSpPrHiddenFill(buffer) {
  const zip = await extractZip(buffer);
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = zip.textFiles.get(slidePath);
  assert.ok(slideXml, "fixture must include slide1.xml");
  assert.match(slideXml, /<p:pic\b/, "fixture must include a picture for spPr injection");

  const updated = slideXml.replace(
    /(<p:pic\b[\s\S]*?<p:spPr\b[\s\S]*?<a:prstGeom[\s\S]*?<\/a:prstGeom>)(\s*<a:noFill\/>)?/,
    `$1<a:noFill/>${SPPR_HIDDEN_FILL}`,
  );
  assert.notEqual(updated, slideXml, "expected to inject a14:hiddenFill into picture spPr");

  return buildZip(buffer, new Map([[slidePath, updated]]));
}

test("insert shape with AI transform preserves spPr a14:hiddenFill", async () => {
  const sourceBuffer = await injectSpPrHiddenFill(toArrayBuffer(await readDeck(FIXTURE)));
  const { PresentationEngine } = await loadPresentationEngineModule();
  const engine = await PresentationEngine.load(sourceBuffer);

  const shapeIndex = await engine.addShapeGeometry(0, "rect");
  await engine.applyInsertedShapeTransform(0, shapeIndex, TRANSFORM);

  const exported = await engine.export();
  const { validatePowerPointExportContents } = await loadPowerPointPackageModule();
  const validation = await validatePowerPointExportContents(sourceBuffer, exported);
  assert.equal(
    validation.ok,
    true,
    `export validation failed after insert+transform: ${validation.errors.join("; ")}`,
  );

  const slideXml = (await extractZip(exported)).textFiles.get("ppt/slides/slide1.xml");
  assert.match(slideXml, /<a14:hiddenFill\b/, "exported slide should keep spPr a14:hiddenFill");
});

test("workshop deck AI insert+transform saves when deck is available", async () => {
  try {
    await access(WORKSHOP_DECK);
  } catch {
    return;
  }

  const { build } = await import("esbuild");
  const { mkdtemp, mkdir, readFile, writeFile } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const path = await import("node:path");
  const { createRequire } = await import("node:module");
  const { patchPptxRendererSource } = await import("../scripts/lib/patch-pptx-renderer.mjs");

  const projectRoot = path.resolve(import.meta.dirname, "..");
  const require = createRequire(import.meta.url);
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-workshop-ai-"));
  const obsidianModuleDirectory = path.join(outputDirectory, "node_modules", "obsidian");
  await mkdir(obsidianModuleDirectory, { recursive: true });
  await writeFile(path.join(obsidianModuleDirectory, "package.json"), '{"name":"obsidian","main":"index.js"}');
  await writeFile(
    path.join(obsidianModuleDirectory, "index.js"),
    'class TFile { constructor(filePath, extension) { this.path = filePath; this.extension = extension; } } module.exports = { TFile };',
  );
  const { TFile } = require(path.join(obsidianModuleDirectory, "index.js"));
  const vaultPath = WORKSHOP_DECK.replace(/^.*ObsidianNotes\//, "");
  const bytes = await readFile(WORKSHOP_DECK);
  const store = new Map([[vaultPath, bytes]]);
  const vault = {
    store,
    getAbstractFileByPath(filePath) {
      return store.has(filePath) ? new TFile(filePath, "pptx") : null;
    },
    async readBinary(file) {
      return Buffer.from(store.get(file.path)).buffer.slice(0);
    },
    async modifyBinary(file, buffer) {
      store.set(file.path, Buffer.from(buffer));
    },
  };
  const outfile = path.join(outputDirectory, "pptx-service.cjs");
  await build({
    absWorkingDir: outputDirectory,
    alias: await getDocxRuntimeAliases(projectRoot),
    entryPoints: [path.join(projectRoot, "src/ai/pptxDocumentService.ts")],
    bundle: true,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
    external: ["obsidian"],
    loader: { ".wasm": "binary" },
    plugins: [
      {
        name: "inline-pptx-svg-wasm",
        setup(buildContext) {
          buildContext.onLoad({ filter: /pptx-renderer\.js$/ }, async ({ path: modulePath }) => ({
            contents: patchPptxRendererSource(await readFile(modulePath, "utf8")),
            loader: "js",
          }));
        },
      },
    ],
  });
  const { PptxDocumentService } = require(outfile);
  const service = new PptxDocumentService({
    vault,
    normalizePath: (value) => value,
    findOpenPptxView: () => null,
    findOpenDocxView: () => null,
  });

  assert.equal((await service.describe(vaultPath)).ok, true);
  assert.equal(
    (
      await service.apply(vaultPath, [
        { op: "pptx.updateShapeText", slideIndex: 0, shapeIndex: 0, text: "AI workshop test" },
      ])
    ).ok,
    true,
  );
  assert.equal((await service.save(vaultPath)).ok, true);
  assert.equal(
    (
      await service.apply(vaultPath, [
        { op: "pptx.addShape", slideIndex: 31, geometry: "rect", transform: TRANSFORM },
      ])
    ).ok,
    true,
  );
  const saveResult = await service.save(vaultPath);
  assert.equal(saveResult.ok, true, JSON.stringify(saveResult.errors));
  assert.equal((await service.undo(vaultPath)).ok, true);
  assert.equal((await service.redo(vaultPath)).ok, true);
});
