import assert from "node:assert/strict";
import { test } from "node:test";
import { loadSettingsModule } from "./helpers/load-plugin-modules.mjs";

test("normalizeDefaultZoom normalizes inputs correctly", async () => {
    const { normalizeDefaultZoom, MIN_DEFAULT_ZOOM, MAX_DEFAULT_ZOOM, DEFAULT_ZOOM } = await loadSettingsModule();

    // Valid zoom values inside range
    assert.equal(normalizeDefaultZoom(1), 1);
    assert.equal(normalizeDefaultZoom(1.25), 1.25);

    // Boundary values
    assert.equal(normalizeDefaultZoom(0.5), 0.5);
    assert.equal(normalizeDefaultZoom(2), 2);

    // Below MIN bounds
    assert.equal(normalizeDefaultZoom(0.1), MIN_DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(0.49), MIN_DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(0), MIN_DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(-1), MIN_DEFAULT_ZOOM);

    // Above MAX bounds
    assert.equal(normalizeDefaultZoom(2.1), MAX_DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(5), MAX_DEFAULT_ZOOM);

    // Rounding to nearest DEFAULT_ZOOM_STEP (0.05)
    assert.equal(normalizeDefaultZoom(1.02), 1);
    assert.equal(normalizeDefaultZoom(1.03), 1.05);

    // Type coercion
    assert.equal(normalizeDefaultZoom("1.5"), 1.5);

    // Invalid/non-finite inputs
    assert.equal(normalizeDefaultZoom(NaN), DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(undefined), DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(null), MIN_DEFAULT_ZOOM); // Number(null) is 0 -> MIN_DEFAULT_ZOOM
    assert.equal(normalizeDefaultZoom("invalid"), DEFAULT_ZOOM);
    assert.equal(normalizeDefaultZoom(Infinity), DEFAULT_ZOOM);
});

test("formatZoom formats zoom factor to percentage string", async () => {
    const { formatZoom } = await loadSettingsModule();

    assert.equal(formatZoom(1), "100%");
    assert.equal(formatZoom(0.5), "50%");
    assert.equal(formatZoom(1.25), "125%");
    assert.equal(formatZoom(2), "200%");

    // Rounding logic in formatZoom (Math.round(value * 100))
    assert.equal(formatZoom(1.004), "100%");
    assert.equal(formatZoom(1.006), "101%");
});

test("normalizeEditorThemePreference normalizes theme preferences", async () => {
    const { normalizeEditorThemePreference } = await loadSettingsModule();
    // Assuming DEFAULT_SETTINGS.editorTheme is 'system' or similar, let's verify fallback

    // Valid values
    assert.equal(normalizeEditorThemePreference("light"), "light");
    assert.equal(normalizeEditorThemePreference("dark"), "dark");
    assert.equal(normalizeEditorThemePreference("system"), "system");

    // Invalid values
    assert.equal(normalizeEditorThemePreference("invalid"), "system"); // fallback is typically system
    assert.equal(normalizeEditorThemePreference(undefined), "system");
    assert.equal(normalizeEditorThemePreference(null), "system");
});

test("resolveEditorThemePreference resolves 'system' to systemTheme", async () => {
    const { resolveEditorThemePreference } = await loadSettingsModule();

    // Explicit light/dark
    assert.equal(resolveEditorThemePreference("light", "dark"), "light");
    assert.equal(resolveEditorThemePreference("light", "light"), "light");
    assert.equal(resolveEditorThemePreference("dark", "light"), "dark");

    // 'system' resolves to the second argument
    assert.equal(resolveEditorThemePreference("system", "light"), "light");
    assert.equal(resolveEditorThemePreference("system", "dark"), "dark");

    // Invalid inputs fall back to 'system', which resolves to the second argument
    assert.equal(resolveEditorThemePreference("invalid", "light"), "light");
    assert.equal(resolveEditorThemePreference("invalid", "dark"), "dark");
});

test("mergeNativePowerPointDocEditorSettings merges properties correctly", async () => {
    const { mergeNativePowerPointDocEditorSettings } = await loadSettingsModule();

    // 1. Empty saved settings
    const resultEmpty = mergeNativePowerPointDocEditorSettings(null, "light");
    assert.equal(resultEmpty.settings.editorTheme, "system");
    assert.equal(resultEmpty.settings.defaultZoom, 1);
    assert.equal(resultEmpty.settings.debugLogging, false);
    assert.equal(resultEmpty.settings.powerPointAutosaveEnabled, true);
    assert.equal(resultEmpty.settings.powerPointShowInspector, false);
    assert.equal(resultEmpty.hadLegacyEditorLanguage, false);
    // When saved is null, shouldPersistSettings is false because we don't need to overwrite defaults right away (or true if defaults differ from undefined)
    // Looking at the code: it checks raw.editorTheme !== normalizedEditorTheme -> undefined !== 'system' -> true
    assert.equal(resultEmpty.shouldPersistSettings, true);

    // 2. Exact match to defaults
    const exactDefaults = {
        authorName: "Obsidian User",
        editorTheme: "system",
        showRuler: false,
        autosave: true,
        createBackupsBeforeSave: false,
        defaultZoom: 1,
        enableDocxSearchIndex: false,
        autoIndexDocxSearch: false,
        debugLogging: false,
        powerPointAutosaveEnabled: true,
        powerPointHideUnsupportedSvgContent: false,
        powerPointOpenWithYoloMode: false,
        powerPointShowInspector: false,
        disableDocxFiles: false,
        disablePowerPointFiles: false,
        enableAiInterfacing: false,
        addAiSkill: false,
    };
    const resultExact = mergeNativePowerPointDocEditorSettings(exactDefaults, "light");
    assert.equal(resultExact.shouldPersistSettings, false);

    // 3. Legacy settings conversion
    const legacySaved = {
        powerPointRemoveUnsupportedSvgContent: true, // Should map to powerPointHideUnsupportedSvgContent
        powerPointYoloMode: true, // Should map to powerPointOpenWithYoloMode
        editorLanguage: "en", // Sets hadLegacyEditorLanguage
    };

    const resultLegacy = mergeNativePowerPointDocEditorSettings(legacySaved, "light");
    assert.equal(resultLegacy.hadLegacyEditorLanguage, true);
    assert.equal(resultLegacy.settings.powerPointHideUnsupportedSvgContent, true);
    assert.equal(resultLegacy.settings.powerPointOpenWithYoloMode, true);
    assert.equal(resultLegacy.shouldPersistSettings, true);

    // 4. Overriding defaults
    const overrideSaved = {
        editorTheme: "dark",
        defaultZoom: 1.5,
        debugLogging: true,
        powerPointShowInspector: true,
        powerPointAutosaveEnabled: false,
    };

    const resultOverride = mergeNativePowerPointDocEditorSettings(overrideSaved, "light");
    assert.equal(resultOverride.settings.editorTheme, "dark");
    assert.equal(resultOverride.settings.defaultZoom, 1.5);
    assert.equal(resultOverride.settings.debugLogging, true);
    assert.equal(resultOverride.settings.powerPointShowInspector, true);
    assert.equal(resultOverride.settings.powerPointAutosaveEnabled, false);
    // Missing fields in overrideSaved will be normalized, e.g. undefined !== normalized -> shouldPersistSettings = true
    assert.equal(resultOverride.shouldPersistSettings, true);
});
