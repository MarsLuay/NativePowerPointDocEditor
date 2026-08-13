# Cross-cutting behavior

## Selection and focus

The hidden ProseMirror editor is the selection authority. Toolbar focus can
temporarily move DOM focus away from it, so non-empty selections are preserved
until editor input or a new visible-page selection resets them. Keyboard input
must update the same selection state rather than relying on the previous text
paragraph.

## Persistence

DOCX edits mark the open view dirty, schedule autosave, serialize through the
DOCX runtime, and write the vault file. The deployed editor also maintains a
developer debug log when debug logging is enabled.

## Diagnostics

Use the project logger in `src/logger.ts` and the deployed
`.obsidian/plugins/native-powerpoint-doc-editor/dev-debug.log`. Relevant areas
are `editor`, `save`, `load`, `render`, and `font-preservation`. Logs should
record paragraph/range identifiers and summary values, never document bodies.

## Generated artifacts

Build output is deployed to `.obsidian/plugins/native-powerpoint-doc-editor/`.
Source changes require a build and a reload/hot-reload before live behavior is
considered verified.
