# Vendored Eigenpal DOCX Editor

This directory vendors the open-source Eigenpal DOCX editor packages used by the DOCX side of the plugin.

## Source

- Upstream repository: https://github.com/eigenpal/docx-editor
- Upstream packages: `@eigenpal/docx-editor-react`, `@eigenpal/docx-editor-core`, `@eigenpal/docx-editor-i18n`, `@eigenpal/docx-editor-agents`
- Vendored version: `1.9.0`
- License: Apache-2.0

Each package directory keeps its upstream `package.json`, `README.md`, `LICENSE`, and published `dist/` artifacts.

The vendored `package.json` files intentionally set `sideEffects: true` so esbuild preserves package CSS and bare chunk imports when bundling from local files. Keep this local bundling metadata unless a focused bundle audit proves a narrower allowlist is safe.

## Architecture

The plugin has no npm dependency on `@eigenpal/*`. Locale loaders import committed artifacts directly. Package specifiers remain where plugin components and published artifacts rely on the upstream export-map identities, but `esbuild.config.mjs` maps every exported runtime path to this committed vendor tree. `tsconfig.json` maps the same package names to committed declaration files.

This keeps the DOCX dependency boundary explicit while letting the PPTX implementation remain independent. PPTX code should not import from this vendor tree unless a shared plugin-owned abstraction is created first.

Do not add an `@eigenpal/*` package to the root `package.json`. If a new upstream artifact imports an export that is not vendored and aliased, add that artifact to this tree or fail the upgrade rather than silently resolving it from `node_modules`.

## Update Process

1. Confirm the target Eigenpal release is open source and license-compatible.
2. Replace the package directories here with the matching npm package artifacts.
3. Keep package `LICENSE` files with the vendored code.
4. Keep the root package manifest and lockfile free of `@eigenpal/*` npm dependencies.
5. Run typecheck/build and DOCX/PPTX smoke checks.

## Local DOM Contract

The plugin-owned integration points are exposed through `data-native-powerpoint-doc-editor-*` markers added to the vendored artifacts. Runtime DOCX glue should use the constants in `src/docxEditorChromeMarkers.ts` instead of Eigenpal classes or test ids for editor chrome, toolbars, table toolbars, page containers, rendered pages, rendered paragraphs, list markers, tab runs, hyperlink popups, hidden ProseMirror roots, carets, and Eigenpal tooltips.

Eigenpal's own internal code may still use its private classes such as `.layout-page` and `.paged-editor__pages` inside the vendored renderer. Treat those as upstream implementation details; plugin code should only add new dependencies through owned markers or explicit vendored APIs.
