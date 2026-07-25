# NPDE DOCX source branch

This monorepo exists only on the `docx-editor-source` branch.

| Field | Value |
|-------|-------|
| Branch | `npde-mirror-1.9.0` |
| Commit | `66d74702b8fa26fa3a309e93ea5cd617e09cb7c8` |
| Upstream | gone private; cloned from community archives |
| License | Apache-2.0 |

Build and test source here:

```bash
npm run build:docx-editor   # from plugin root; needs bun
npm run build               # or npm run dev
```

## Release handoff

Do not merge or copy `docx-editor/` into `nightly-releases` or `main`.

After source changes are built and tested, a release branch runs `npm run vendor:docx` with this worktree as `DOCX_EDITOR_SOURCE_DIR`. That command copies only the allowlisted runtime output into `vendor/docx-editor-runtime/`, writes provenance, and rejects TypeScript, declarations, workspace metadata, and package build tooling.

Obsidian release artifacts remain `main.js` / `manifest.json` / `styles.css`. Release branches contain the vendor snapshot, not this source tree.

### Branch roles

- `docx-editor-source`: editable DOCX source and Bun package builds.
- `nightly-releases`: latest vendor-only plugin release candidate.
- `main`: vendor-only stable plugin release.

## Former git remotes (nested .git removed so ObsidianNotes tracks this tree)
mhur	https://github.com/mhurhangee/docx-editor.git (fetch) [blob:none]
mhur	https://github.com/mhurhangee/docx-editor.git (push)
origin	https://github.com/KanRule/docx-editor.git (fetch) [blob:none]
origin	https://github.com/KanRule/docx-editor.git (push)
soren	https://github.com/sorenlouv/docx-editor.git (fetch) [blob:none]
soren	https://github.com/sorenlouv/docx-editor.git (push)
Detached from commit: 66d74702b8fa26fa3a309e93ea5cd617e09cb7c8
