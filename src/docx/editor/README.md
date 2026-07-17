# DOCX runtime and source worktree

`main` contains a generated, allowlisted DOCX runtime snapshot at `vendor/docx-editor-runtime/` (Apache-2.0). It contains only package JavaScript, CSS, required JSON, license material, and `provenance.json`; it never contains DOCX editor TypeScript or declarations.

The full editable editor monorepo is preserved on branch `docx-editor-source` in the adjacent worktree `/Users/mars/NPDE-docx-editor-source`. Obsidian users download only the bundled plugin assets, not either worktree's source tree.

## Layout

| Path | Role |
|------|------|
| `src/docx/runtime/` | Plugin-owned TypeScript contract, bridge declaration, runtime bridge, and CSS boundary |
| `vendor/docx-editor-runtime/` | Generated JS/CSS runtime snapshot used by this branch |
| `/Users/mars/NPDE-docx-editor-source/` | Full editable DOCX monorepo on `docx-editor-source` |

Plugin TypeScript imports the local `src/docx/runtime` facade. Only `src/docx/runtime/bridge.mjs` imports `@npde/docx-editor-*`; `scripts/lib/docx-editor-aliases.mjs` resolves those runtime imports to `vendor/docx-editor-runtime/` and keeps one root React/ReactDOM copy.

**Do not** add `@npde/docx-editor-*` to the plugin root `package.json`, import package paths directly from plugin TypeScript, or hand-edit the generated snapshot.

## Editing plugin code

Run `npm run verify:review`, then `npm run dev` or `npm run build`. This branch is independently type-checkable and does not require Bun.

## Editing the DOCX editor runtime

1. In `/Users/mars/NPDE-docx-editor-source`, edit the monorepo and run its Bun build, typecheck, and tests.
2. In this worktree, run `npm run vendor:docx` (or set `DOCX_EDITOR_SOURCE_DIR` for a different source location).
3. Run `npm run verify:review` and commit the resulting runtime snapshot with its updated provenance.

## AI and save

There is no `@npde/docx-editor-agents` package and no core `DocumentAgent` tree. Plugin AI lives under `src/ai` only. React DOCX save uses `Document` + packers (`exportDocxBuffer`) directly.
