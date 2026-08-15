# Obsidian community plugin

> **Public GitHub branches:** `docx-editor-source` (edit DOCX monorepo) → `nightly-releases` (latest plugin + `vendor:docx`) → `main` (promote when satisfied). Release branches are vendor-only — no in-repo `docx-editor/`. See section **Branch workflow** below.
>
> **This vault tree** may still keep a local `docx-editor/` for analysis/dev; do not push that monorepo onto `nightly-releases` / `main`.


## Project overview

- Target: Obsidian Community Plugin (TypeScript → bundled JavaScript).
- Entry point: `main.ts` compiles to `main.js`, Obsidian loads.
- Release artifacts required: `main.js`, `manifest.json`, optional `styles.css`. Optional PPTX/HEIC engines are gzip-embedded in `main.js` and materialized as sibling `.mjs` files on load — do not attach those `.mjs` files to GitHub releases (Obsidian will not download them).

## Native PowerPoint Doc Editor architecture

- Use `src/menuControls.ts` for popovers, menus, menu items, menu sections, select rows, checkbox rows, action rows. Hand-build rows only when patching unreplacable third-party editor DOM.
- Render DOCX settings from shared descriptors in `src/settings.ts`; do not duplicate labels, descriptions, defaults, option lists between Obsidian settings tab and in-editor DOCX settings menu.
- Resolve editor theme via plugin-level path in `src/main.ts`. DOCX/PPTX roots consume `resolvedEditorTheme`; views must not inspect `document.body` or call `resolveEditorThemePreference()` locally.
- Theme colors use `--npde-*` tokens. Hardcoded color literals live in token definitions, not component rules.
- Do not hand-edit generated DOCX runtime CSS under `vendor/docx-editor-runtime/`. Apply DOCX CSS changes on `docx-editor-source`; root plugin `styles.css` may use line-scoped `obsidian: allow css-important` only for documented third-party / PDF isolation.
- Do not use CSS `:has()`, stylesheet `text-indent`, `break-before`/`page-break-*`, or `@tailwind` in scanned source. Stamp class hooks / indent in JS; inject Tailwind at build time; page breaks stay in the layout engine.
- Prefer CSS2 single-keyword `text-decoration` only. Tint deletes with `box-shadow` / `background` / `color`.
- After theme/menu/settings changes, run guards: `npm run check:theme-architecture`, `npm run check:theme-css`, `npm run check:shared-ui-patterns`.

### Branch workflow (`docx-editor-source` → `nightly-releases` → `main`)

Release branches (`nightly-releases`, then `main`) must **not** contain an in-repo `docx-editor/` monorepo. Latest plugin work lands on **`nightly-releases`**; when that branch is good, promote it to **`main`**.

| Branch | Role |
|--------|------|
| `docx-editor-source` | Editable DOCX editor monorepo only (Bun build/test). Not the Obsidian plugin release surface. |
| `nightly-releases` | Latest plugin updates. Builds against `vendor/docx-editor-runtime` refreshed from `docx-editor-source` via `npm run vendor:docx`. |
| `main` | Stable promote target. Same vendor-only layout as nightly once promoted. Catalog / default-branch review uses this shape. |

Flow:

1. Edit DOCX runtime on `docx-editor-source` (worktree e.g. `/Users/mars/NPDE-docx-editor-source` or `DOCX_EDITOR_SOURCE_DIR`). Build/test there with Bun.
2. On `nightly-releases`, run `npm run vendor:docx` to refresh `vendor/docx-editor-runtime/` + `provenance.json` from that source commit.
3. Implement plugin changes on `nightly-releases`; verify with `npm run verify:review` (no Bun / no in-repo monorepo).
4. When satisfied with nightly, merge or fast-forward **`nightly-releases` → `main`** (keep vendor-only; never reintroduce `docx-editor/`).

### DOCX runtime snapshot (vendor-only on release branches)

- Release branches vendor only generated JS/CSS at `vendor/docx-editor-runtime/{core,react,i18n}` plus `provenance.json` (source branch + commit). No DOCX TypeScript or package `.d.ts` on these branches.
- Plugin TypeScript imports DOCX facilities from `src/docx/runtime` only. `src/docx/runtime/bridge.mjs` is the sole `@npde/*` import boundary; `src/docx/runtime/styles.ts` is the sole vendored-CSS boundary.
- `scripts/lib/docx-editor-aliases.mjs` resolves `@npde/docx-editor-*` to `vendor/docx-editor-runtime` and keeps one root `react` / `react-dom` pair. Keep plugin ProseMirror versions aligned with the source worktree.
- Plugin AI remains `src/ai`; do **not** add `@npde/*` to root `package.json`. Details: `src/docx/editor/README.md`.

### Agent DOCX editing (vault)

- Prefer plugin AI bridge (`describe` / `apply` / `openSession().save()`), not Computer Use.
- Multi-run template bodies: put full text on first run, clear siblings.
- Missing capability → `src/ai` catalog + executor + tests + `npm run ai:generate` + build (not one-off ZIP hacks).
- Vault playbook: `.agents/skills/01-personal-vault/native-docx-plugin-edit/SKILL.md`.

### PPTX action logging

- Use `debugLog(area, message, data?)` for PPTX actions; `warnLog` and `errorLog` are always on. Prefer `logPptxAction(area, op, data?)` for user-triggered action starts.
- Reuse areas: `save`, `insert`, `arrange`, `selection`, `text-format`, `text-edit`, `text-select`, `search`, `history`, `export`, `inspector`, `slide`, `render`, `clipboard`, `view`. Add a specific area only when none fit (for example, `snap`, `menu`, `engine`, `mutate`).
- Keep payloads small: `{ file?, slide?, shapeIndexes?, op, ... }`; never include large buffers or XML.
- Log user-triggered operation start; log completed/failed only for async or fallible operations.
- Never log every `pointermove`; log drag start, end/cancel, and commit only.

## Environment & tooling

- Node.js: current LTS; Node 18+ recommended.
- **Package manager: npm** required here; `package.json` defines scripts/deps.
- **Bundler: esbuild** required here; `esbuild.config.mjs` and build scripts depend on it. Other projects may use Rollup/webpack if all external deps bundle into `main.js`.
- Types: `obsidian` type definitions.

**Note**: This sample depends on npm/esbuild. From scratch, choose other tools only if replacing build config.

### Install

```bash
npm install
```

### Dev (watch)

```bash
npm run dev
```

### Production build

```bash
npm run build
```

## Linting

- Install eslint: `npm install -g eslint`
- Analyze project: `eslint main.ts`
- eslint creates report with improvement suggestions by file/line.
- If source is in folder like `src`, analyze all files with: `eslint ./src/`

## File & folder conventions

- **Organize code into multiple files**: Split modules; avoid everything in `main.ts`.
- Source lives in `src/`. Keep `main.ts` small: lifecycle only (loading, unloading, commands).
- **Example file structure**:
  ```
  src/
    main.ts           # Plugin entry point, lifecycle management
    settings.ts       # Settings interface and defaults
    commands/         # Command implementations
      command1.ts
      command2.ts
    ui/              # UI components, modals, views
      modal.ts
      view.ts
    utils/           # Utility functions, helpers
      helpers.ts
      constants.ts
    types.ts         # TypeScript interfaces and types
  ```
- **Do not commit build artifacts**: Never commit `node_modules/`, `main.js`, or generated files.
- Keep plugin small. Avoid large deps. Prefer browser-compatible packages.
- Generated output goes plugin root or `dist/` per build. Release artifacts must be top-level in vault plugin folder: `main.js`, `manifest.json`, `styles.css`.

## Manifest rules (`manifest.json`)

- Must include (non-exhaustive):
  - `id` (plugin ID; local dev should match folder name)
  - `name`
  - `version` (Semantic Versioning `x.y.z`)
  - `minAppVersion`
  - `description`
  - `isDesktopOnly` (boolean)
  - Optional: `author`, `authorUrl`, `fundingUrl` (string or map)
- Never change `id` after release. Stable API.
- Keep `minAppVersion` accurate for newer APIs.
- Canonical requirements here: https://github.com/obsidianmd/obsidian-releases/blob/master/.github/workflows/validate-plugin-entry.yml

## Testing

- Manual install: copy `main.js`, `manifest.json`, `styles.css` (if any) to:
  ```
  <Vault>/.obsidian/plugins/<plugin-id>/
  ```
- Reload Obsidian; enable plugin in **Settings → Community plugins**.

## Commands & settings

- Add user-facing commands via `this.addCommand(...)`.
- If config exists, provide settings tab and sane defaults.
- Persist settings using `this.loadData()` / `this.saveData()`.
- Use stable command IDs; avoid renaming after release.

## Versioning & releases

- Bump `version` in `manifest.json` (SemVer) and update `versions.json` mapping plugin version → minimum app version.
- Create GitHub release tag exactly matching `manifest.json`'s `version`. No leading `v`.
- Attach `manifest.json`, `main.js`, and `styles.css` (if present) as individual assets.
- After first release, follow community catalog add/update process.

## Security, privacy, and compliance

Follow Obsidian's **Developer Policies** and **Plugin Guidelines**. Key rules:

- Default local/offline. Network only when feature needs it.
- No hidden telemetry. Optional analytics/third-party calls require explicit opt-in and clear docs in `README.md` and settings.
- Never execute remote code, fetch/eval scripts, or auto-update plugin code outside normal releases.
- Minimize scope: read/write only needed vault data. Do not access files outside vault.
- Disclose external services, data sent, risks.
- Respect privacy. Do not collect vault contents, filenames, or personal info unless essential and explicitly consented.
- Avoid deceptive patterns, ads, spammy notifications.
- Register/clean up DOM, app, interval listeners with `register*` helpers so unload is safe.

## UX & copy guidelines (for UI text, commands, settings)

- Prefer sentence case for headings, buttons, titles.
- Use clear, action-oriented imperatives in steps.
- Use **bold** for literal UI labels. Prefer "select" for interactions.
- Use arrows for navigation: **Settings → Community plugins**.
- Keep in-app strings short, consistent, jargon-free.

## Performance

- Keep startup light. Defer heavy work.
- Avoid long-running tasks during `onload`; use lazy init.
- Batch disk access; avoid excessive vault scans.
- Debounce/throttle expensive file event work.

## Coding conventions

- TypeScript with `"strict": true` preferred.
- **Keep `main.ts` minimal**: lifecycle only (onload, onunload, addCommand calls). Feature logic goes separate modules.
- **Split large files**: If file exceeds ~200-300 lines, consider smaller focused modules.
- **Use clear module boundaries**: One file, one clear responsibility.
- Bundle all into `main.js` (no unbundled runtime deps).
- Avoid Node/Electron APIs for mobile compatibility; set `isDesktopOnly` accordingly.
- Prefer `async/await` over promise chains; handle errors gracefully.

## Mobile

- Test iOS/Android where feasible.
- Do not assume desktop-only unless `isDesktopOnly` is `true`.
- Avoid large in-memory structures; mind memory/storage limits.

## Agent do/don't

**Do**
- Add commands with stable IDs; do not rename after release.
- Provide defaults and validation in settings.
- Write idempotent reload/unload paths; no listener/interval leaks.
- Use `this.register*` helpers for cleanup.

**Don't**
- Add network calls without obvious user-facing reason and docs.
- Ship cloud-required features without clear disclosure and explicit opt-in.
- Store/transmit vault contents unless essential and consented.

## Common tasks

### Organize code across multiple files

**main.ts** (minimal, lifecycle only):
```ts
import { Plugin } from "obsidian";
import { MySettings, DEFAULT_SETTINGS } from "./settings";
import { registerCommands } from "./commands";

export default class MyPlugin extends Plugin {
  settings: MySettings;

  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    registerCommands(this);
  }
}
```

**settings.ts**:
```ts
export interface MySettings {
  enabled: boolean;
  apiKey: string;
}

export const DEFAULT_SETTINGS: MySettings = {
  enabled: true,
  apiKey: "",
};
```

**commands/index.ts**:
```ts
import { Plugin } from "obsidian";
import { doSomething } from "./my-command";

export function registerCommands(plugin: Plugin) {
  plugin.addCommand({
    id: "do-something",
    name: "Do something",
    callback: () => doSomething(plugin),
  });
}
```

### Add a command

```ts
this.addCommand({
  id: "your-command-id",
  name: "Do the thing",
  callback: () => this.doTheThing(),
});
```

### Persist settings

```ts
interface MySettings { enabled: boolean }
const DEFAULT_SETTINGS: MySettings = { enabled: true };

async onload() {
  this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  await this.saveData(this.settings);
}
```

### Register listeners safely

```ts
this.registerEvent(this.app.workspace.on("file-open", f => { /* ... */ }));
this.registerDomEvent(window, "resize", () => { /* ... */ });
this.registerInterval(window.setInterval(() => { /* ... */ }, 1000));
```

## Troubleshooting

- Plugin doesn't load after build: ensure `main.js` and `manifest.json` are top-level in plugin folder under `<Vault>/.obsidian/plugins/<plugin-id>/`.
- Build issues: if `main.js` missing, run `npm run build` or `npm run dev` to compile TypeScript.
- Commands missing: verify `addCommand` runs after `onload` and IDs unique.
- Settings not persisting: ensure `loadData`/`saveData` awaited and UI re-renders after changes.
- Mobile-only issues: confirm no desktop-only APIs; check `isDesktopOnly` and adjust.

## References

- Obsidian sample plugin: https://github.com/obsidianmd/obsidian-sample-plugin
- API documentation: https://docs.obsidian.md
- Developer policies: https://docs.obsidian.md/Developer+policies
- Plugin guidelines: https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines
- Style guide: https://help.obsidian.md/style-guide

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
