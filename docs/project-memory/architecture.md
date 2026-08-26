# Architecture

- This project is an Obsidian community plugin. `manifest.json` identifies it as `native-powerpoint-doc-editor`; `src/main.ts` is the TypeScript entry point and the build emits root `main.js`.
- `src/main.ts` loads settings and locale services, resolves the editor theme, configures logging and runtime artifact loading, optionally enables the AI interface, and independently loads DOCX and PowerPoint support according to settings. It registers commands/settings and flushes open views during unload.
- DOCX and PowerPoint functionality are separated behind `src/docx/` and `src/powerpoint/` modules. The supported PowerPoint extensions are `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, and `.potm`.
- `esbuild.config.mjs` bundles the source and, when the vault plugin directory exists, deploys `main.js`, runtime artifacts, `styles.css`, `manifest.json`, and `locales/`/`ai/` outputs there. `tsconfig.json` includes `src/**/*.ts`, `src/**/*.tsx`, and `src/**/*.json`.
- Release branches vendor the DOCX runtime under `vendor/docx-editor-runtime/`; generated runtime CSS is not edited in this plugin root. PPTX/HEIC runtime artifacts are materialized as sibling `.mjs` files at load and are not GitHub release attachments.
