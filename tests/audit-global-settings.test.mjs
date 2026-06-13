// Audit test for the "Global · Settings & commands" feature group.
//
// Targets the three pure helpers that gate how persisted settings are
// normalized before they reach the editor views and the PowerPoint runtime:
//
//   1. normalizeDefaultZoom (src/settings.ts ~76-84): clamps the saved DOCX
//      zoom into 0.5-2 and snaps it to the 0.05 slider step; non-numeric or
//      non-finite input must fall back to the 1.0 default.
//   2. getNativePowerPointSettings (src/settings.ts ~63-74): maps the flat
//      DocxidianSettings.powerPoint* fields onto the NativePowerPointSettings
//      shape the view consumes, and threads the setOpenWithYoloMode callback.
//   3. normalizeDocxidianLanguage (src/locales.ts ~40-42): keeps a known BCP-47
//      code and falls back to the default ('en') for unknown/non-string input.
//
// settings.ts statically `extends PluginSettingTab` from 'obsidian', so the
// module cannot be required without an 'obsidian' implementation. We mark
// 'obsidian' external in the esbuild output and install a minimal Module._load
// shim (mirroring tests/helpers/load-plugin-modules.mjs) for the require().
// locales.ts has no 'obsidian' dependency, so it bundles cleanly on its own.

import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import Module, { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { test } from "node:test";
import { build } from "esbuild";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

let tempDirectoryPromise;
function getTempDirectory() {
  tempDirectoryPromise ??= mkdtemp(path.join(tmpdir(), "native-powerpoint-settings-audit-"));
  return tempDirectoryPromise;
}

// Mirror of the (un-exported) bundleSource pattern in
// tests/helpers/load-plugin-modules.mjs.
async function bundleSource(entry, outputName, external = []) {
  const outputDirectory = await getTempDirectory();
  const outfile = path.join(outputDirectory, outputName);
  await build({
    entryPoints: [path.join(projectRoot, entry)],
    bundle: true,
    external,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  return outfile;
}

// Minimal 'obsidian' stub: settings.ts only needs PluginSettingTab to be a
// constructable base class at module-evaluation time. Setting/Notice/App are
// referenced inside DocxidianSettingTab.renderSettings(), which the pure
// helpers never invoke, but we provide harmless stubs anyway.
function loadWithObsidianStub(outfile) {
  const originalLoad = Module._load;
  Module._load = function load(request, parent, isMain) {
    if (request === "obsidian") {
      return {
        App: class App {},
        Notice: class Notice {},
        PluginSettingTab: class PluginSettingTab {
          constructor(app, plugin) {
            this.app = app;
            this.plugin = plugin;
          }
        },
        Setting: class Setting {},
      };
    }
    return originalLoad.call(this, request, parent, isMain);
  };
  try {
    return require(outfile);
  } finally {
    Module._load = originalLoad;
  }
}

let settingsModulePromise;
function loadSettingsModule() {
  settingsModulePromise ??= bundleSource("src/settings.ts", "settings.cjs", ["obsidian"]).then(
    loadWithObsidianStub,
  );
  return settingsModulePromise;
}

let localesModulePromise;
function loadLocalesModule() {
  localesModulePromise ??= bundleSource("src/locales.ts", "locales.cjs").then((outfile) =>
    require(outfile),
  );
  return localesModulePromise;
}

test("normalizeDefaultZoom clamps to 0.5-2 and snaps to the 0.05 step", async () => {
  const { normalizeDefaultZoom } = await loadSettingsModule();

  // Below min / above max clamp to the bounds.
  assert.equal(normalizeDefaultZoom(0), 0.5, "0 clamps up to min 0.5");
  assert.equal(normalizeDefaultZoom(0.1), 0.5, "below min clamps to 0.5");
  assert.equal(normalizeDefaultZoom(5), 2, "above max clamps to 2");
  assert.equal(normalizeDefaultZoom(-3), 0.5, "negative clamps to min");

  // Bounds pass through unchanged.
  assert.equal(normalizeDefaultZoom(0.5), 0.5);
  assert.equal(normalizeDefaultZoom(2), 2);

  // Snap-to-step rounding (0.05 increments). Allow float epsilon.
  assert.ok(Math.abs(normalizeDefaultZoom(1.02) - 1.0) < 1e-9, "1.02 -> 1.00");
  assert.ok(Math.abs(normalizeDefaultZoom(1.03) - 1.05) < 1e-9, "1.03 -> 1.05");
  assert.ok(Math.abs(normalizeDefaultZoom(1.234) - 1.25) < 1e-9, "1.234 -> 1.25");
  assert.ok(Math.abs(normalizeDefaultZoom(0.77) - 0.75) < 1e-9, "0.77 -> 0.75");
});

test("normalizeDefaultZoom falls back to default 1 for non-numeric / non-finite input", async () => {
  const { normalizeDefaultZoom } = await loadSettingsModule();

  assert.equal(normalizeDefaultZoom("not a number"), 1, "non-numeric string -> NaN -> default");
  assert.equal(normalizeDefaultZoom(undefined), 1, "undefined -> NaN -> default");
  assert.equal(normalizeDefaultZoom(NaN), 1, "NaN -> default");
  assert.equal(normalizeDefaultZoom(Infinity), 1, "Infinity -> not finite -> default");
  assert.equal(normalizeDefaultZoom(-Infinity), 1, "-Infinity -> not finite -> default");
  assert.equal(normalizeDefaultZoom({}), 1, "object -> NaN -> default");

  // NOTE (audit finding): null and "" are NOT treated as "non-numeric".
  // Number(null) === 0 and Number("") === 0, both finite, so they clamp to the
  // 0.5 minimum rather than the 1.0 default. Harmless (settings always store a
  // number) but documents the actual coercion behavior.
  assert.equal(normalizeDefaultZoom(null), 0.5, "null -> Number(null)=0 -> clamps to min 0.5");
  assert.equal(normalizeDefaultZoom(""), 0.5, "empty string -> Number('')=0 -> clamps to min 0.5");

  // Numeric strings ARE coerced (Number("1.5") === 1.5) then normalized.
  assert.equal(normalizeDefaultZoom("1.5"), 1.5, "numeric string coerces");
});

test("getNativePowerPointSettings maps DocxidianSettings -> NativePowerPoint shape", async () => {
  const { getNativePowerPointSettings, DEFAULT_SETTINGS } = await loadSettingsModule();

  const settings = {
    ...DEFAULT_SETTINGS,
    powerPointAutosaveEnabled: false,
    powerPointHideUnsupportedSvgContent: true,
    powerPointOpenWithYoloMode: true,
    powerPointShowInspector: true,
  };

  let captured;
  const setOpenWithYoloMode = async (value) => {
    captured = value;
  };

  const result = getNativePowerPointSettings(settings, setOpenWithYoloMode);

  assert.equal(result.autosaveEnabled, false, "autosaveEnabled <- powerPointAutosaveEnabled");
  assert.equal(
    result.hideUnsupportedSvgContent,
    true,
    "hideUnsupportedSvgContent <- powerPointHideUnsupportedSvgContent",
  );
  assert.equal(result.openWithYoloMode, true, "openWithYoloMode <- powerPointOpenWithYoloMode");
  assert.equal(result.showInspector, true, "showInspector <- powerPointShowInspector");

  // The callback is threaded straight through.
  assert.equal(typeof result.setOpenWithYoloMode, "function");
  await result.setOpenWithYoloMode(true);
  assert.equal(captured, true, "setOpenWithYoloMode callback invoked with the passed value");

  // The mapped shape only carries the five expected keys.
  assert.deepEqual(
    Object.keys(result).sort(),
    ["autosaveEnabled", "hideUnsupportedSvgContent", "openWithYoloMode", "setOpenWithYoloMode", "showInspector"],
  );
});

test("getNativePowerPointSettings supplies a no-op callback default", async () => {
  const { getNativePowerPointSettings, DEFAULT_SETTINGS } = await loadSettingsModule();

  const result = getNativePowerPointSettings(DEFAULT_SETTINGS);
  assert.equal(typeof result.setOpenWithYoloMode, "function");
  // Default callback resolves without throwing.
  await assert.doesNotReject(() => result.setOpenWithYoloMode(true));

  // DEFAULT_SETTINGS map straight through.
  assert.equal(result.autosaveEnabled, true);
  assert.equal(result.hideUnsupportedSvgContent, false);
  assert.equal(result.openWithYoloMode, false);
  assert.equal(result.showInspector, false);
});

test("normalizeDocxidianLanguage keeps known codes and falls back to 'en'", async () => {
  const { normalizeDocxidianLanguage, DEFAULT_LANGUAGE } = await loadLocalesModule();

  assert.equal(DEFAULT_LANGUAGE, "en");

  // Every supported code is preserved verbatim.
  for (const code of ["en", "pl", "pt-BR", "tr", "he", "zh-CN"]) {
    assert.equal(normalizeDocxidianLanguage(code), code, `${code} preserved`);
  }

  // Unknown / non-string input falls back to the default.
  assert.equal(normalizeDocxidianLanguage("fr"), "en", "unknown code");
  assert.equal(normalizeDocxidianLanguage("EN"), "en", "wrong case is not a known code");
  assert.equal(normalizeDocxidianLanguage(""), "en", "empty string");
  assert.equal(normalizeDocxidianLanguage(undefined), "en", "undefined");
  assert.equal(normalizeDocxidianLanguage(null), "en", "null");
  assert.equal(normalizeDocxidianLanguage(42), "en", "number");
  assert.equal(normalizeDocxidianLanguage({}), "en", "object");
});
