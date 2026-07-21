# Changelog

All notable changes to the Native PowerPoint Doc Editor plugin are documented in
this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.54] - 2026-07-21

### Changed

- Release-branch DOCX runtime is vendor-only (`vendor/docx-editor-runtime` from `docx-editor-source` via `npm run vendor:docx`). Removed in-repo `docx-editor/` from `nightly-releases`. Documented `docx-editor-source` → `nightly-releases` → `main` flow in AGENTS.md. Removed bogus `versions.json` `"undefined"` key.

### Fixed

- AI/DOCX unit tests resolve vendored `@npde/docx-editor-*` aliases under `vendor/docx-editor-runtime`.
- Release-branch DOCX font/IME harness scripts skip cleanly when harness entries are absent.
- Enable `obsidianmd/prefer-active-doc`; document `DOCX_EDITOR_SOURCE_DIR`; bump `dompurify` to 3.4.12.

## [1.0.53] - 2026-07-21

### Fixed

- PowerPoint cropped-image resize handles track the visible clip frame instead of the expanded source image bounds.
- PowerPoint Delete no longer no-ops after filmstrip focus leaves a stale multi-shape selection.
- PowerPoint deleting charts/tables/groups records protected-marker allowances so save validation succeeds.
- PowerPoint cut/duplicate honor multi-select; nested group hit-testing prefers the top-level shape.
- PowerPoint AI multi-`deleteShape` ops coalesce to one batched delete before index renumber.
- PowerPoint shape delete prunes orphan media/chart/embedding parts and blocks them from resurrecting on save merge.
- PowerPoint Cmd+A / whole-box text clear hides bullet and number markers in the live SVG preview (not only after blur).
- PowerPoint overlapping Delete while a shape delete is in flight queues a fresh-selection rerun instead of dropping the second request.

## [1.0.52] - 2026-07-17

### Fixed

- PowerPoint selection outline edge strips move the shape; only corner and mid-edge resize dots stretch it.
- Multi-select south (and pure vertical) group resize stretches all rows instead of translating the bottom via text preview inverse-compensation.
- Shift/Ctrl/Meta while a multi-select union outline is active hit-tests shapes underneath instead of forcing group drag.
- Per-shape dashed outlines appear again as multi-select grows (including marquee preview).

## [1.0.51] - 2026-07-17

### Fixed

- PowerPoint empty/centered text boxes wrap typed characters horizontally instead of stacking one glyph per line.
- PowerPoint text color/highlight/font popovers keep formatting context after selection clears.
- PowerPoint Backspace on an empty first paragraph no longer drops the text box out of edit.
- PowerPoint color-picker swatches stay filled in dark chrome (not blank until hover).

## [1.0.50] - 2026-07-17

### Fixed

- Multi-select move now live-previews selected shapes with the selection outline (not outline-only).


## [1.0.49] - 2026-07-17

### Fixed

- PowerPoint whole-shape text commits no longer skip OOXML writes after a live SVG preview update (authoritative baseline text).
- Inline text preview measurement uses a detached `Window.createEl('canvas')` so pop-out / XML slide documents do not throw during reflow.

## [1.0.48] - 2026-07-16

### Fixed

- Catalog export now removes stale analyzer caches, DOCX test fixtures, source tooling, and TypeScript from the public runtime mirror; its strict surface guard prevents them from returning.
- PowerPoint **Send forward** and **Send backward** now move a shape exactly one z-order position, including around pictures.
- PowerPoint selection, inline text previews, shape-fill swatches, and text-box resize behavior stay visually stable while editing.

## [1.0.47] - 2026-07-16

### Added

- AI: `docx.replaceBodyParagraphs` for writing multi-paragraph DOCX bodies without chaining paragraph breaks.
- AI: `ai.createDocument({ path, kind, paragraphs?, overwrite? })` to create blank vault DOCX/PPTX packages (optional DOCX paragraphs).

### Fixed

- PowerPoint filmstrip thumbnails preserve nonstandard slide aspect ratios instead of forcing a standard format.
- PowerPoint two-finger gestures pan horizontally and vertically; they no longer invert into pinch zoom, and zoomed canvases can reach the upper and left bounds.
- PowerPoint rich-text editing: reliable caret/word/paragraph selection, `Ctrl+A` within a text box, and increment/decrement font-size synchronization for mixed-format text.
- PowerPoint text editing preserves blank paragraphs and reflows content after inserting paragraph breaks instead of overflowing a text box.
- New PowerPoint text boxes use the right-click position.
- DOCX editors route a typed space through the active editor once, preventing duplicate spacing.

## [1.0.46] - 2026-07-15

### Fixed

- Catalog mirror verify: comment-marker tests load `docx-editor` package **dist** (no package `src/` in public tree).

## [1.0.45] - 2026-07-15

### Fixed

- DOCX comments: overlapping threads keep marks; empty range repair; flush/save races (hydrate dirty, getComments ref, strip empty comments.xml).
- Prefer Obsidian DOM helpers (`createEl` / `createDiv` / `createSvg`) over `createElement`; inject host/print CSS via `adoptedStyleSheets`.

### Added

- Settings tab `getSettingDefinitions()` (Obsidian 1.13+ settings search) with dual-support `display()`.
- Code-analysis ESLint policy requires `obsidianmd/prefer-create-el` and `obsidianmd/settings-tab/prefer-setting-definitions` (`eslint-plugin-obsidianmd` 0.4.1).

## [1.0.43] - 2026-07-15

### Fixed

- DOCX chrome in Obsidian dark mode: `theme-system` no longer paints light `--npde-chrome-*` over `theme-resolved-dark`; document surface remaps stay on white pages.
- Caret stays page ink (`#000000`) on non-inverted DOCX pages in dark Obsidian (vendor `--doc-caret` was light).
- Formatting-bar vendor tooltips (`Insert link (Ctrl+K)` etc.): `ToolbarButton` / `Tooltip` no longer inject titles; plugin chrome stamps only.

### Changed

- Code-analysis: `css/theme-system-light-chrome` and `obsidian/vendor-floating-toolbar-tooltip` (plus theme-css mirror check).
- Catalog mirror ESLint: when `docx-editor/CATALOG_SURFACE.md` is present, disable type-checked TypeScript rules so JS-only package dist verifies (fixes failed 1.0.42 release CI).

## [1.0.42] - 2026-07-15

### Changed

- Catalog mirror Option A: public packages are **JS-only** (no package `.d.ts` / `types` fields). Sync drops decls; types stay in the vault. Code-analysis fails catalog-shaped trees that still ship package declarations (`catalog/dts-not-excluded`). Sanitized public `.d.ts` retired.
- Catalog surface build skips `tsc` (`scripts/typecheck-for-surface.mjs`) so clean clones typecheck-free via esbuild against package JS.
- Drop leftover Option A / agent hygiene: ambient-stub `rmSync`, `docs/AGENT-API.md` exclude/gitignore, hazard scanners + async catalog check wrapper, empty `dist/agent`, stale `./agent` docs.
- Remove always-error MCP mutation stubs (`docx_insert_text` / replace / delete / format / apply_style, `docx_insert_variable`). Drop `@npde/docx-editor-agents` package rows from local READMEs / changeset fixed set.

## [1.0.41] - 2026-07-15

### Changed

- Unbundle Eigenpal `DocumentAgent` / AgentPanel from the Obsidian DOCX editor path: save uses `Document` + `exportDocxBuffer` (selective / repack / create).
- Remove `@npde/docx-editor-agents` stub/aliases; plugin AI remains `src/ai` only.
- Re-home content-control types on `@npde/docx-editor-core/contentControls`.
- Public catalog mirror sanitizes package `*.d.ts` for Obsidian catalog ESLint (`globalThis`, Identifier `document`, `#private`, etc.). Code-analysis gates the same patterns.

### Added

- AI OOXML ops: `docx.insertText`, `docx.deleteRange`, `docx.insertHyperlink`, `docx.removeHyperlink`, `docx.insertParagraphBreak`.

## [1.0.40] - 2026-07-15

### Fixed

- DOCX editor a11y: screen-reader labels on header/footer, response preview, and table-options controls.
- Print path builds the print document with DOM APIs (no `document.write`).
- Reliability: await empty-DOCX/arrayBuffer helpers; catch unhandled dynamic-import / composition promise chains.
- Clear empty `innerHTML` assignments and agent-layer text-helper duplication without changing agent offset contracts.

### Changed

- Deduplicate shading parse/serialize and consolidate agent navigation helpers onto `text-utils`.
- Vault code-analysis scans Obsidian-runtime `docx-editor/packages/{core,react,i18n}` and skips demos / unused framework packages.

## [1.0.39] - 2026-07-15

### Fixed

- Unblock Obsidian Community catalog checks: public mirror ships dist-only `docx-editor/packages/{core,react,i18n}` (no monorepo TypeScript sources / agents / vue / nuxt). Catalog ESLint scans public `.ts`/`.tsx` regardless of local eslint ignores.
- Rename command id `copy-native-powerpoint-doc-editor-debug-log` → `copy-debug-log` (plugin id must not appear in command ids).

### Changed

- Add `scripts/sync-obsidian-catalog-mirror.mjs` and make `build:docx-editor` verify committed dist on catalog-shaped trees.

## [1.0.38] - 2026-07-14

### Fixed

- Clear nested editor CSS guardrails without suppressions: build-time Tailwind inject, class hooks instead of `:has`, CSS2 `text-decoration`, JS-stamped indent, no page-break fragmentation CSS, no nested `!important`.
- Deduplicate React / ProseMirror identities across the in-repo `docx-editor` bun tree (tsconfig paths + esbuild `react` aliases) so harness verifies and `tsc` stay green.

### Changed

- Pin plugin ProseMirror packages to the docx-editor versions and document the dual-copy pin in `AGENTS.md`.

## [1.0.37] - 2026-07-14

### Changed

- Document intentional DOCX editor CSS guardrail exceptions (`@tailwind`, Word layout `:has`/page-break cues, selection and find-highlight `!important`, track-changes decoration color).
- Point `package.json` `repository` at `MarsLuay/NativePowerPointDocEditor` so release tooling resolves the authoritative GitHub repo from the vault subtree.

## [1.0.36] - 2026-07-14

### Added

- Embed the Eigenpal docx-editor source monorepo under `docx-editor/` (1.9.0 pin) and wire package dist into the plugin build; Obsidian downloads remain `main.js` only.
- Add `npm run build:docx-editor` and require a fresh monorepo package rebuild on publish.

### Changed

- Replace `src/vendor/eigenpal` committed packages with in-repo `docx-editor/packages/{core,react,i18n}` plus `agentsStub` for AgentPanel.
- Move the pure-JS PPTX engine to `src/powerpoint/backend/pptxJsEngine.mjs`.

### Fixed

- Restore DOCX IME zoom-wrapper and font-roundtrip harness targeting by recognizing Eigenpal `.paged-editor__pages` / `.paged-editor__hidden-pm` and stamping matching plugin chrome markers.

## [1.0.35] - 2026-07-14

### Added

- Add a dedicated GitHub Suggestion issue form with the `[Suggestion]` title prefix.

## [1.0.34] - 2026-07-14

### Fixed

- Select a word on a double-click and all text in a text box on a triple-click.
- Keep text geometry stable while resizing a text box, then reflow it when the resize is committed.
- Add a fill-color action to shape context menus.
- Save every dirty open presentation before development plugin reloads, and abort the reload if any save fails.
- Handle background DOCX/PPTX rendering, search, locale, clipboard, and navigation promise failures without unhandled rejections.

### Changed

- Resolve DOCX editor packages through the vendored source and remove direct Eigenpal package dependencies where local equivalents exist.

## [1.0.30] - 2026-06-24

### Fixed

- Strip React 19 `react-dom` `createElement("script")` patterns from the production
  bundle so Obsidian's automated plugin review no longer flags dynamic script
  injection.
- Remove an unnecessary `ArrayBuffer` type assertion in chart workbook buffer
  handling.

### Added

- Post-build guard that fails when `createElement("script")` leaks into `main.js`.

## [1.0.29] - 2026-06-24

### Changed

- Upgrade the vendored DOCX editor packages to 1.9.0 for improved Word-compatible
  pagination and layout fidelity.

### Fixed

- Prevent a terminal empty DOCX paragraph from creating an extra preview page
  when the visible document content already fits on the preceding page.

### Added

- Add a repeatable DOCX pagination audit that compares preview page counts with
  LibreOffice-rendered PDF page counts, distinguishes preview over-pagination
  from renderer-specific reference differences, and reports tab-heavy source
  paragraphs that can paginate differently across office suites.
- Add live DOCX pagination diagnostics to the development log, including page
  counts, page geometry, explicit page breaks, and tab-heavy paragraph counts.

## [1.0.28] - 2026-06-23

### Changed

- Bump `js-yaml` to 4.2.0 (security), `prosemirror-view` to 1.41.9, and
  `@npde/docx-editor-i18n` to 1.9.0.
- Bump React and React DOM to 19.2.7 with matching type packages.
- Bump selected dev dependencies (`@types/node`, `globals`, `jiti`, `tslib`,
  `typescript-eslint`, `electron`).

### Fixed

- Restore TypeScript build compatibility with `@types/node` 26 and `Array.at`
  usage (`ES2022` lib, chart workbook buffer cast).

## [1.0.27] - 2026-06-23

### Added

- Publish `docs/privacy-policy.md` and `docs/terms-of-service.md` for the community
  plugin distribution.
- Add Dependabot configuration for weekly npm dependency update PRs.

### Changed

- Link the README privacy section to the new legal documents.

### Fixed

- Resolve ESLint `no-unsafe-assignment` in logger prototype normalization.

## [1.0.26] - 2026-06-23

### Fixed

- Preserve direct table-cell font sizes across DOCX save when eigenpal drops
  `w:sz` run properties.
- Improve Japanese IME candidate anchoring with caret-aware hidden-editor
  positioning and stronger zoom-transform neutralization.

### Changed

- Add feature diagnostics logging across DOCX and PowerPoint workflows.
- Expand debug log retention and add log statistics for the copy-debug-log command.
- Add DOCX font-roundtrip verification harness and table font-size fixture tests.

## [1.0.25] - 2026-06-23

### Fixed

- DOCX bullet and numbered lists now get default hanging indent when toggled on,
  and list layout remeasures after indent normalization.

## [1.0.24] - 2026-06-22

### Fixed

- Insert link dialog centering when the file explorer sidebar is expanded.
- Gray out and block the comments sidebar toggle when a document has no comments.
- PPTX bullet lists now get default hanging indent and bullet font on insert.

## [1.0.22] - 2026-06-22

### Fixed

- Bundle DOCX and PPTX support into `main.js` so community releases ship only
  Obsidian-supported assets (`main.js`, `manifest.json`, `styles.css`).
- Community-plugin lint fixes: cross-window DOM `instanceOf` checks, IME transform
  neutralizer style updates, dynamic imports instead of `require()`, deprecated
  `activeLeaf`, and floating promise handling.

## [1.0.21] - 2026-06-22

### Fixed

- DOCX Japanese IME candidate window placement on multi-monitor setups by
  neutralizing eigenpal zoom/outline CSS transforms (`scale` / `translateX` →
  `zoom` / `margin-left`).

### Changed

- Lazy-load DOCX and PPTX runtime chunks to reduce plugin startup cost.
- Defer vault-wide DOCX search indexing, locale loading, filmstrip rendering,
  and inspector UI until idle or first use.

## [1.0.19] - 2026-06-13

### Security

- Bump dev-only `esbuild` to 0.28.1 and `electron` to 39.8.10 to clear Dependabot
  alerts (build/smoke tooling only; not bundled in the release plugin).

## [1.0.18] - 2026-06-12

### Fixed

- Community plugin review: stop reading Chromium version from `navigator.userAgent`;
  desktop builds now resolve it via `Platform.isDesktop` and Electron
  `process.versions.chrome`, wired through `configureChromiumVersionReader`.

## [1.0.17] - 2026-06-12

### Fixed

- Community plugin review: replace unreachable LinkedIn `authorUrl` with GitHub profile.
- Remove disallowed `eslint-disable` for `obsidianmd/no-global-this`; parse Chromium
  version from `navigator.userAgent` instead of `process.versions`.
- Replace raw `localStorage` / `globalThis` JS-engine dev overrides with
  `setForceJsBackendOverride` (tests) and `setForceJsBackendDevOverride` (plugin).
- ESLint warnings: safe `unknown` error handling, unnecessary type assertion,
  `setInterval` async callback, and cross-window `instanceOf` for flip wrappers.
- CSS lint: drop `!important` swatch override and `text-decoration` on failed-save
  badge (use `border-bottom` instead).

## [1.0.16] - 2026-06-12

### Added

- Shape flip display pass (`shapeFlipTransforms.ts`) so horizontal/vertical flip edits
  are visible in the canvas and thumbnails even though pptx-svg only records flip
  flags in OOXML data attributes.
- Click-to-retry on the **Save failed** status badge; failed autosaves also
  schedule a follow-up attempt after five seconds.
- Headless coverage for flip OOXML toggling and mirror-transform math
  (`tests/shape-flip.test.mjs`).

### Fixed

- **Flip horizontal / flip vertical** now work for pictures, autoshapes, text boxes,
  and groups (not only images in the context menu). Grouped-shape edits resolve via
  `getShapeElementByRendererIndex`.
- **Save failed** no longer leaves the editor stuck without context: the last error
  is stored, surfaced in the status tooltip, and autosave retries when still dirty.
- Removed duplicate `applyShapeFlipTransforms` definition that blocked `npm run build`.
- Cleared remaining oxlint warnings (`no-useless-spread`, `no-this-alias`) in the
  PowerPoint host adapters and paragraph-clear helper.

### Changed

- Pure-JS PowerPoint engine regeneration script (`npm run regen:pptx-js`) that
  clones the matching `pptx-svg` source, builds the JS target with MoonBit, and
  rewraps it into the vendored fallback — documented in `CONTRIBUTING.md`.
- Mobile PPTX smoke test (`npm run smoke:mobile-pptx`) that renders a deck
  through the JS fallback via `PresentationEngine`, wired into lint and release
  CI to keep the iOS/Android path honest.
- `authorUrl` in the manifest.
- Mobile-aware runtime error message: on iOS/Android it points to the App
  Store / Play Store instead of the desktop installer download.
- Bumped `@types/node` from `^16` to `^22` to match the CI Node runtime (22.x/24.x).
- Split the large `NativePowerPointView.ts` into focused `src/powerpoint/`
  modules (extensions, constants, types, runtime-compat, text/SVG utils, find and
  history controllers).

## [1.0.14] - 2026-06-11

### Added

- **Lazy-loaded pure-JS PPTX engine fallback** for older Obsidian installers and
  mobile WebViews that lack WebAssembly GC (Chromium < 119). PPTX files now open
  on these runtimes instead of failing with a "Wasm init failed" error; the
  faster Wasm engine is still used wherever WebAssembly GC is available.
- Friendly, actionable runtime-compatibility message when neither engine can run,
  explaining that the Obsidian installer (not the in-app updater) controls the
  bundled Chromium version.
- Developer/test override to force the JS backend
  (`localStorage` key `native-powerpoint-force-js-engine`, or
  `globalThis.__NATIVE_PPTX_FORCE_JS__` in Node) plus a `smoke:pptx-js` test.

## [1.0.13] - 2026-06-10

### Fixed

- DOCX and PPTX editor reliability issues.

## [1.0.12] - 2026-06-10

### Fixed

- Text selection loss when pressing modifier keys.
- Toolbar now shows inherited fonts for runs without an explicit font.

## [1.0.11] - 2026-06-10

### Fixed

- Find-highlight offset for bulleted text.

## [1.0.9] - 2026-06-10

### Added

- UI screenshot in the README.

### Fixed

- General bug fixes.

## [1.0.8] - 2026-06-10

### Added

- iOS support.

### Fixed

- PPTX rendering fidelity fixes.

## [1.0.7] - earlier

### Changed

- CSS specificity refinements for the editor views without relying on
  `!important`.

## [1.0.4] - earlier

### Fixed

- Obsidian plugin review issues addressed for the community release.

## [1.0.0] - earlier

### Added

- Initial community release: open, view, and edit DOCX and PPTX files directly
  inside the Obsidian vault.

[Unreleased]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.47...HEAD
[1.0.47]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.46...1.0.47
[1.0.28]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.27...1.0.28
[1.0.27]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.26...1.0.27
[1.0.26]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.25...1.0.26
[1.0.25]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.24...1.0.25
[1.0.24]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.22...1.0.24
[1.0.22]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.21...1.0.22
[1.0.21]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.19...1.0.21
[1.0.19]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.18...1.0.19
[1.0.18]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.17...1.0.18
[1.0.17]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.16...1.0.17
[1.0.16]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.15...1.0.16
[1.0.15]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.15
[1.0.14]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.14
[1.0.13]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.13
[1.0.12]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.12
[1.0.11]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.11
[1.0.9]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.9
[1.0.8]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.8
[1.0.7]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.7
[1.0.4]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.4
[1.0.0]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.0
