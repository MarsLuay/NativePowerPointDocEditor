# Local DOCX editor packages

The editable Eigenpal DOCX editor **source monorepo** lives at `docx-editor/` (branch `npde-mirror-1.9.0`, commit `66d74702…`, Apache-2.0). Package **dist** under `docx-editor/packages/{core,react,i18n}/dist` is what esbuild bundles into `main.js`.

Obsidian users only download release `main.js` (+ manifest/css) — **not** this monorepo. Size/performance of that bundle is unchanged as long as you ship the same dist inputs.

## Layout

| Path | Role |
|------|------|
| `docx-editor/` | Full clone (TypeScript source + seeded/generated `dist`) |
| `docx-editor/packages/{core,react,i18n}/` | Wired into the plugin via `scripts/lib/docx-editor-aliases.mjs` |
| `Projects/docx-editor-mirror-1.9.0/` (vault) | Optional npm `.tgz` insurance outside this plugin |

Plugin imports use `@npde/docx-editor-*`; in-dist imports still use `@eigenpal/*` (compat aliases).

**Do not** add `@eigenpal/*` to the plugin root `package.json`.

## Edit → reflect in Obsidian

1. Edit TypeScript under `docx-editor/packages/*/src/`.
2. Rebuild package dist: `npm run build:docx-editor` (requires [bun](https://bun.sh)).
3. Rebuild / reload the plugin: `npm run dev` or `npm run build`.

Until you run step 2, the plugin uses the committed/seeded `dist/` (same bytes as the previous standalone packages).

**Publish:** the publish skill always runs `npm run build:docx-editor` then `npm run build` for this project before release assets — never ship stale dist when that script exists.

## AI and save

There is no `@eigenpal/docx-editor-agents` package in this tree. Plugin AI lives under `src/ai` only. React DOCX save uses `Document` + packers directly — not `DocumentAgent`.
