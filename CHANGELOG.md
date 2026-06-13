# Changelog

All notable changes to the Native PowerPoint Doc Editor plugin are documented in
this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Pure-JS PowerPoint engine regeneration script (`npm run regen:pptx-js`) that
  clones the matching `pptx-svg` source, builds the JS target with MoonBit, and
  rewraps it into the vendored fallback — documented in `CONTRIBUTING.md`.
- Mobile PPTX smoke test (`npm run smoke:mobile-pptx`) that renders a deck
  through the JS fallback via `PresentationEngine`, wired into lint and release
  CI to keep the iOS/Android path honest.
- `authorUrl` in the manifest.

### Changed

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

[Unreleased]: https://github.com/MarsLuay/NativePowerPointDocEditor/compare/1.0.14...HEAD
[1.0.14]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.14
[1.0.13]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.13
[1.0.12]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.12
[1.0.11]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.11
[1.0.9]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.9
[1.0.8]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.8
[1.0.7]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.7
[1.0.4]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.4
[1.0.0]: https://github.com/MarsLuay/NativePowerPointDocEditor/releases/tag/1.0.0
