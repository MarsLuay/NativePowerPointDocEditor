# Module — src-docx-runtime

**Paths:** `src/docx/runtime/bridge.mjs`, `vendor/docx-editor-runtime/`

**Purpose:** Provide the vendored DOCX editor/runtime implementation used by
the plugin for ProseMirror selection, paragraph commands, pagination, and
serialization.

**Public surface:** The bridge imported by plugin source and the runtime APIs
used by `DocxReactView.tsx`.

**Depends on:** The `docx-editor-source` runtime snapshot and the plugin’s
single `@npde/*` bridge boundary.

**Invariants:** Runtime vendor output is generated; source changes belong in
the source worktree and are then refreshed into the plugin vendor snapshot.
Keyboard movement must retain empty paragraphs in the document model.

**Related functions:** [functions/src-docx-editor.md](../functions/src-docx-editor.md)

**Related types:** [types/src-docx-editor.md](../types/src-docx-editor.md)
