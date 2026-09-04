# Known Failures

- Documentation drift: `Native PowerPoint Doc Editor.md` records version `1.0.14`, while `manifest.json` and `package.json` record `1.1.15`. The project note identifies the manifest/CHANGELOG as the version source of truth.
- Release-instructions ambiguity: `README.md` says manual setup requires “all six release files”, while the local project guidance says GitHub releases should contain `main.js`, `manifest.json`, and `styles.css`; generated `.mjs` runtime artifacts are materialized on load and should not be attached.
- Heading underline vs following text (resume DOCX): paginated layout used line-box height only for `w:pBdr` paragraphs, and the painter overlay used negative `top`/`bottom` equal to `w:space`. The stroke sat on the next paragraph's glyphs (SUMMARY / EDUCATION / PROJECTS). Fixed in `docx-editor-source` `531ca08` by reserving space+width in fragment height and painting the overlay inside that box. Reload the plugin after vendoring; Obsidian keeps the previous `main.js` in memory until reload.
- No runtime or test failure was established during this bounded memory pass.
