# Contributing to Native PowerPoint Doc Editor

Thanks for helping improve Native PowerPoint Doc Editor. This plugin opens and edits `.docx` and `.pptx` files directly inside Obsidian.

## Before you open a pull request

- Open an issue first for larger changes so scope is clear.
- Keep each pull request focused on one feature, fix, or cleanup.
- Do not commit generated artifacts such as `main.js`, `test-results/`, or `scripts/visual-output/`.
- Do not include private vault files, documents, screenshots, or logs unless they are sanitized.

## Local setup

```bash
npm install
npm run build
npm run lint
npm run smoke
```

For changes to rendering, editing, clipboard, charts, fonts, or export code, also run the focused smoke commands that match your change:

```bash
npm run smoke:generated-text
npm run smoke:chart-data
npm run smoke:objects
npm run smoke:fonts
npm run smoke:pptx-js
npm run smoke:mobile-pptx
```

Documentation-only changes usually do not need a full build.

## Project conventions

- Prefer Obsidian APIs and browser-safe DOM patterns over Node-only APIs.
- Use `activeDocument` for DOM access that should work in Obsidian popout windows.
- Use cross-window-safe element checks from `src/domGuards.ts` for DOM nodes created outside the main window.
- Keep user-facing text concise and consistent with existing plugin wording.
- Preserve conservative file-safety behavior: conflict checks, validation, recovery copies, and view-only fallbacks.

## PowerPoint rendering

PPTX files are rendered with the `pptx-svg` WebAssembly engine when the host supports WebAssembly GC. Older Chromium builds use a pure-JS fallback in `src/powerpoint/backend/pptxJsEngine.mjs`.

If you change PPTX rendering or bump `pptx-svg`, run:

```bash
npm run smoke:pptx-js
npm run smoke:mobile-pptx
```

After bumping `pptx-svg`, regenerate the local JS engine and commit it:

```bash
npm run regen:pptx-js
```

See `patches/README.md` for upstream patch notes.

## Mobile support

The plugin is not desktop-only. `npm run build` runs a mobile compatibility check to ensure `main.js` does not statically require Node or Electron APIs.

If your change affects PPTX loading, rendering, touch interaction, or saving, verify it on Obsidian Mobile when you can.

## Releases

Community plugin releases should include only:

- `main.js`
- `manifest.json`
- `styles.css`

Bump `manifest.json`, `package.json`, and `versions.json` together, tag the release to match the manifest version exactly, and update `CHANGELOG.md`.

## Security and data safety

If you find a bug that could corrupt files, expose private vault data, or bypass Obsidian security expectations, open a minimal issue without private samples. Share sensitive reproduction files only after they have been sanitized.
