# Contributing to Native PowerPoint Doc Editor

Thanks for helping improve Native PowerPoint Doc Editor. This plugin is maintained by Mars and is built to make DOCX and PowerPoint files usable directly inside Obsidian.

## Before You Start

- Open an issue for larger changes so the scope is clear before implementation.
- Keep pull requests focused on one feature, fix, or cleanup at a time.
- Avoid committing generated artifacts such as `main.js`, `test-results/`, or `scripts/visual-output/`.
- Do not include private vault content, documents, presentations, screenshots, or logs unless they are intentionally sanitized.

## Local Setup

Install dependencies:

```bash
npm install
```

Run a production build:

```bash
npm run build
```

Run the core checks:

```bash
npm run lint
npm run smoke
```

Run focused PowerPoint and DOCX checks when touching rendering, editing, clipboard, chart, font, or export code:

```bash
npm run smoke:generated-text
npm run smoke:chart-data
npm run smoke:objects
npm run smoke:fonts
npm run smoke:pptx-js
npm run smoke:mobile-pptx
npm run visual:caret
```

## PowerPoint engine backends

PPTX rendering uses the `pptx-svg` WebAssembly (wasm-gc) engine, with a pure-JS
engine fallback (`src/vendor/pptx-js-engine.mjs`) for runtimes that lack
WebAssembly GC (Obsidian installer < 1.5.8 / Chromium < 119). Because the
fallback normally only runs on old installers, it is easy to break without
noticing. Two guards keep it honest:

- `npm run smoke:pptx-js` renders a deck through the JS engine (runs in CI and
  before every release).
- To exercise the fallback on a modern machine, open Obsidian's developer
  console and run
  `app.plugins.plugins['native-powerpoint-doc-editor'].setForceJsBackendDevOverride(true)`,
  then reopen the PPTX. Call it with `false` to return to the Wasm engine. In
  Node/tests, call `setForceJsBackendOverride(true)` on the bundled
  `PresentationEngine` module instead.

### Regenerating the JS fallback

The `pptx-svg` npm package ships only the wasm-gc binary, so the pure-JS engine
is built from source and vendored. When you bump `pptx-svg`, regenerate the
fallback so it matches the installed version:

```bash
npm run regen:pptx-js
```

This one command clones the `pptx-svg` source at the tag matching the installed
version, patches its `moon.pkg` to emit a JS build, compiles it with MoonBit,
rewraps the output into `src/vendor/pptx-js-engine.mjs`, and verifies it with
`smoke:pptx-js`. Commit the regenerated `src/vendor/pptx-js-engine.mjs` alongside
the dependency bump.

Notes:

- It requires the MoonBit toolchain. If `moon` is not installed, the script
  prints the one-line installer; re-run with `INSTALL_MOONBIT=1` to install it
  automatically.
- To regenerate from a specific ref instead of the version-matched tag, set
  `PPTX_SVG_REF` (for example `PPTX_SVG_REF=v0.6.0 npm run regen:pptx-js`).
- The lower-level `node scripts/build-pptx-js-engine.mjs <moonbit-js-output>` step
  is still available if you already have a MoonBit JS build in hand.

## Mobile (iOS / Android)

`manifest.json` sets `isDesktopOnly: false`, so the plugin is expected to load on
Obsidian Mobile. Mobile uses WKWebView (iOS) or the Android System WebView — no
Node.js or Electron APIs. The build runs `check-mobile-compat.mjs` after every
`npm run build` to reject static `require("electron")` / Node builtin imports in
`main.js`.

### What CI verifies automatically

| Check | What it proves |
| --- | --- |
| `npm run check:mobile` | Plugin bundle won't crash on load due to desktop-only `require()` calls |
| `npm run smoke:mobile-pptx` | `PresentationEngine` can open and render a deck through the **JS fallback** — the path mobile WebViews use when WebAssembly GC is unavailable |
| `npm run smoke:pptx-js` | The vendored JS engine itself renders SVG in isolation |

The JS fallback activates automatically when Wasm init fails (common on older
mobile WebViews). On modern devices with WasmGC, the faster Wasm engine is used
instead.

### Clipboard behavior

- **DOCX** (`DocxReactView.tsx`): tries Electron's clipboard on desktop via
  `window.require('electron')`, then falls back to `navigator.clipboard`. Mobile
  always uses the web clipboard APIs.
- **PPTX**: shape copy/paste is in-memory inside the plugin. Plain-text paste
  uses `navigator.clipboard.readText()` (may require a user gesture on mobile).

### Manual verification (required before claiming mobile support)

CI cannot run Obsidian on a device. Before a release that touches PPTX loading,
rendering, touch interaction, or save paths, spot-check on at least one iOS and
one Android device:

1. Open `tests/fixtures/decks/features.pptx` (or any small deck in the vault).
2. Confirm slides render (not a blank slide or runtime error).
3. Tap a text box, edit, and confirm autosave.
4. Pinch/zoom or toolbar zoom if available; switch slides via thumbnails.

To force the JS fallback on a modern desktop install (same path many mobile
devices use), see [PowerPoint engine backends](#powerpoint-engine-backends).

## Development Notes

- Prefer Obsidian APIs and browser-safe DOM patterns over Node-only APIs.
- Use `activeDocument` for DOM access that should work in Obsidian popout windows.
- Use cross-window-safe element checks for DOM nodes created outside the main window.
- Keep CSS compatible with the minimum Obsidian version in `manifest.json`.
- Keep user-facing text concise and consistent with the existing plugin wording.
- Preserve file safety behavior: conflict checks, validation, recovery copies, and view-only fallbacks should stay conservative.

## Testing Changes

For narrow documentation-only changes, a build is not usually required.

For source, styling, or dependency changes, run:

```bash
npm run lint
npm run build
npm run smoke
```

Then run the focused smoke commands that match the files you changed.

## Release Notes

Release notes should explain user-visible changes and review fixes clearly. Community plugin release assets should only include:

- `main.js`
- `manifest.json`
- `styles.css`

## Reporting Security or Data-Safety Issues

If you find a bug that could corrupt files, expose private vault data, or bypass Obsidian security expectations, open a minimal issue without private samples. Share sensitive reproduction files only after they have been sanitized.
