# Known Failures

- Documentation drift: `Native PowerPoint Doc Editor.md` records version `1.0.14`, while `manifest.json` and `package.json` record `1.1.15`. The project note identifies the manifest/CHANGELOG as the version source of truth.
- Release-instructions ambiguity: `README.md` says manual setup requires “all six release files”, while the local project guidance says GitHub releases should contain `main.js`, `manifest.json`, and `styles.css`; generated `.mjs` runtime artifacts are materialized on load and should not be attached.
- No runtime or test failure was established during this bounded memory pass.
