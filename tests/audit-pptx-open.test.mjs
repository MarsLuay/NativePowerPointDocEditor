import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { loadPowerPointPackageModule } from "./helpers/load-plugin-modules.mjs";
import { readDeck, toArrayBuffer } from "./helpers/renderer.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

// extensions.ts is pure and dependency-free; bundle it on the fly the same way
// tests/helpers/load-plugin-modules.mjs bundles other source modules, so the
// predicates under test are the real source rather than a re-implementation.
let extensionsModulePromise;
function loadExtensionsModule() {
  extensionsModulePromise ??= (async () => {
    const outputDirectory = await mkdtemp(path.join(tmpdir(), "native-powerpoint-audit-open-"));
    const outfile = path.join(outputDirectory, "extensions.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/extensions.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      target: "node22",
    });
    return require(outfile);
  })();
  return extensionsModulePromise;
}

const EDITABLE = ["pptx", "ppsx", "potx"];
const MACRO = ["pptm", "ppsm", "potm"];
const LEGACY = ["ppt", "pps", "pot"];
const UNSUPPORTED = ["docx", "key", "odp", "txt", "pdf", ""];

test("extension predicates classify every PowerPoint extension correctly", async () => {
  const {
    isPowerPointExtension,
    isModernPowerPointExtension,
    isEditablePowerPointExtension,
    isMacroEnabledPowerPointExtension,
  } = await loadExtensionsModule();

  for (const ext of EDITABLE) {
    assert.equal(isPowerPointExtension(ext), true, `${ext} is a PowerPoint extension`);
    assert.equal(isModernPowerPointExtension(ext), true, `${ext} is modern`);
    assert.equal(isEditablePowerPointExtension(ext), true, `${ext} is editable`);
    assert.equal(isMacroEnabledPowerPointExtension(ext), false, `${ext} is not macro-enabled`);
  }

  for (const ext of MACRO) {
    assert.equal(isPowerPointExtension(ext), true, `${ext} is a PowerPoint extension`);
    assert.equal(isModernPowerPointExtension(ext), true, `${ext} is modern`);
    assert.equal(isEditablePowerPointExtension(ext), false, `${ext} is view-only (macro)`);
    assert.equal(isMacroEnabledPowerPointExtension(ext), true, `${ext} is macro-enabled`);
  }

  for (const ext of LEGACY) {
    assert.equal(isPowerPointExtension(ext), true, `${ext} is a PowerPoint extension`);
    assert.equal(isModernPowerPointExtension(ext), false, `${ext} is not modern (legacy)`);
    assert.equal(isEditablePowerPointExtension(ext), false, `${ext} is not editable`);
    assert.equal(isMacroEnabledPowerPointExtension(ext), false, `${ext} is not macro-enabled`);
  }

  for (const ext of UNSUPPORTED) {
    assert.equal(isPowerPointExtension(ext), false, `${ext || "(empty)"} is not a PowerPoint extension`);
    assert.equal(isModernPowerPointExtension(ext), false, `${ext || "(empty)"} is not modern`);
    assert.equal(isEditablePowerPointExtension(ext), false, `${ext || "(empty)"} is not editable`);
    assert.equal(isMacroEnabledPowerPointExtension(ext), false, `${ext || "(empty)"} is not macro-enabled`);
  }
});

test("extension predicates are case-insensitive", async () => {
  const { isPowerPointExtension, isModernPowerPointExtension, isEditablePowerPointExtension } =
    await loadExtensionsModule();

  assert.equal(isPowerPointExtension("PPTX"), true);
  assert.equal(isModernPowerPointExtension("Pptm"), true);
  assert.equal(isEditablePowerPointExtension("PPSX"), true);
});

test("an editable .pptx fixture inspects and validates as openable and not view-only", async () => {
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();
  const { isEditablePowerPointExtension, isMacroEnabledPowerPointExtension } =
    await loadExtensionsModule();

  const inspection = inspectPowerPointPackage(toArrayBuffer(await readDeck("features.pptx")));
  assert.equal(inspection.slidePaths.length, 1);
  assert.equal(inspection.hasVbaProject, false, "editable deck has no VBA project");

  const validation = validatePowerPointPackageStructure(inspection, 1);
  assert.equal(validation.ok, true, validation.errors.join("; "));
  assert.deepEqual(validation.errors, []);

  // Mirrors NativePowerPointView.shouldOpenViewOnly: extension editable + no VBA => editable.
  assert.equal(isEditablePowerPointExtension("pptx"), true);
  const viewOnly = isMacroEnabledPowerPointExtension("pptx") || inspection.hasVbaProject;
  assert.equal(viewOnly, false, "editable .pptx should open editable, not view-only");
});

test("a macro-enabled .pptm fixture is detected as view-only via extension and VBA project", async () => {
  const { inspectPowerPointPackage, validatePowerPointPackageStructure } =
    await loadPowerPointPackageModule();
  const { isMacroEnabledPowerPointExtension, isEditablePowerPointExtension } =
    await loadExtensionsModule();

  const inspection = inspectPowerPointPackage(toArrayBuffer(await readDeck("macro-view-only.pptm")));

  // The package still parses and validates structurally (it opens, view-only).
  const validation = validatePowerPointPackageStructure(inspection, 1);
  assert.equal(validation.ok, true, validation.errors.join("; "));
  assert.ok(
    validation.warnings.some((warning) => /vbaProject\.bin|macro/i.test(warning)),
    "structure validation warns about the embedded macro project",
  );

  // VBA detection flags the embedded macro project.
  assert.equal(inspection.hasVbaProject, true, "macro fixture exposes ppt/vbaProject.bin");

  // Mirrors NativePowerPointView.shouldOpenViewOnly / getViewOnlyReason.
  assert.equal(isMacroEnabledPowerPointExtension("pptm"), true);
  assert.equal(isEditablePowerPointExtension("pptm"), false);
  const viewOnly = isMacroEnabledPowerPointExtension("pptm") || inspection.hasVbaProject;
  assert.equal(viewOnly, true, "macro-enabled .pptm must open view-only");
});
