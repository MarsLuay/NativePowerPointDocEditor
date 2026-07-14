# NPDE in-repo docx-editor mirror

This monorepo is **inside** `Native PowerPoint Doc Editor/docx-editor/`.

| Field | Value |
|-------|-------|
| Branch | `npde-mirror-1.9.0` |
| Commit | `66d74702b8fa26fa3a309e93ea5cd617e09cb7c8` |
| Upstream | gone private; cloned from community archives |
| License | Apache-2.0 |

Plugin esbuild consumes `packages/{core,react,i18n}/dist` only. Rebuild after source edits:

```bash
npm run build:docx-editor   # from plugin root; needs bun
npm run build               # or npm run dev
```

Obsidian release artifacts remain `main.js` / `manifest.json` / `styles.css` — users do not download this tree.

## Former git remotes (nested .git removed so ObsidianNotes tracks this tree)
mhur	https://github.com/mhurhangee/docx-editor.git (fetch) [blob:none]
mhur	https://github.com/mhurhangee/docx-editor.git (push)
origin	https://github.com/KanRule/docx-editor.git (fetch) [blob:none]
origin	https://github.com/KanRule/docx-editor.git (push)
soren	https://github.com/sorenlouv/docx-editor.git (fetch) [blob:none]
soren	https://github.com/sorenlouv/docx-editor.git (push)
Detached from commit: 66d74702b8fa26fa3a309e93ea5cd617e09cb7c8
