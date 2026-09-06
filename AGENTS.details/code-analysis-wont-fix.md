## Code analysis — wont-fix

obsidian: allow missing-release

- `main.js` (`perf/bundle-size`): Soft 90% warning under the Sync Standard 5 MB budget. Gzip-embedded PPTX/HEIC runtimes keep community releases to Obsidian-supported assets while staying under the hard Sync limit; further shrinking would drop engines or reintroduce unsupported release sidecars.
- `src/powerpoint/measureCanvas.ts` (`grouped-diagnostics.recommendation`): Detached measure canvas must use native `Document.createElement` when Obsidian `Window.createEl` is unavailable; `Document.createEl` can throw HierarchyRequestError on SVG/XML owner documents. Factory exempts `prefer-create-el` / `no-deprecated`.
- `src/powerpoint/measureCanvas.ts` (`grouped-diagnostics.warning`): Same detached factory — plugin-review / `obsidianmd/prefer-create-el` is intentional.
- `src/docxEditorChromeDom.ts` (`grouped-diagnostics.warning`): Detached DOCX chrome nodes must use `ownerDocument.createElement` so Obsidian helpers do not append a second document root before sidebar mount; `prefer-create-el` is intentional.
- `src/powerpoint/backend/pptxJsEngine.mjs` (`repo/large-file`): generated pure-JS fallback of `pptx-svg` MoonBit JS backend (`npm run regen:pptx-js`). Must stay Git-tracked for offline Obsidian installs without WASM GC; size is inherent to the engine, not compressible without losing the fallback. External source of truth is the pptx-svg package + regen script.
- `docx-editor/packages/core/src/layout-bridge/footnoteLayout.ts` (`completeness-audit.todo-marker`): upstream deferred footnote layout marker in the vendored DOCX monorepo; do not “finish” it in the plugin layer.
- `docx-editor/**/editor.css` (`mobile-web.no-media-queries` / `mobile-web.small-input-font`): DOCX editor CSS targets Obsidian desktop chrome; mobile media-query / 16px input rules are not product requirements for this plugin surface.
- `docx-editor/packages/i18n/dist/**` (`obsidian-semantic.analysis-degraded`): generated ambient `.d.ts` emits value initializers that TypeScript rejects in declaration files; semantic analysis quality is reduced by that upstream emit, not by plugin source. Fix on `docx-editor-source` / package build, not in the Obsidian plugin layer.
- `src/ai/registerAiCommands.ts` (`obsidian-semantic.dynamic-identifier-unresolved`): command ids come from `AI_COMMAND_IDS.*` constants; static Semgrep cannot resolve the indirection by design.
- `scripts/smoke-selection-geometry.mjs` (`completeness-audit.todo-marker`): The "leaf-total fix" comment documents a past fix to `getNumberOfChars()` under-counting lines and capping the offset; it is a regression test rationale, not a pending action item.
- `src/DocxReactView.tsx` (`code-health.long-function`): `getPlainTextFromInputEvent` is short (6 lines) and clear; do not introduce unnecessary switch statements.
