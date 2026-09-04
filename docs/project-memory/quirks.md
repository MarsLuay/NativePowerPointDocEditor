# Quirks

- The build may deploy directly to `.obsidian/plugins/native-powerpoint-doc-editor` when that directory exists; otherwise deployment is skipped.
- The root contains generated/bundled artifacts such as `main.js`, `pptx-js-engine.mjs`, `pptx-wasm-renderer.mjs`, `heic-decode.mjs`, and `styles.css`; source changes belong in `src/` or the documented runtime source branch.
- Word paragraph borders (`w:pBdr`) keep `w:space` plus stroke width in the paragraph's layout height. The paginator must reserve that inset and the overlay must paint inside it; a negative overlay inset paints through the next paragraph.
- PPTX action logging uses the shared `debugLog`/`warnLog`/`errorLog` facilities, with `logPptxAction` preferred for user-triggered operation starts and small payloads.
- The optional DOCX search index, recovery copies, plugin settings, and development debug log are local data paths; the debug log is enabled only through the documented debug/hot-reload behavior.
