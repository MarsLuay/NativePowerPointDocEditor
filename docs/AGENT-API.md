# AI-Interfacing Design (Agent API)

> Local-only design doc. Gitignored — not shipped in releases.
> Status: **Phases 1–6 implemented.** PPTX + DOCX read/write via headless AI API with per-op schemas, examples, round-trip save validation, and smoke tests.

## Goal

Let any AI extension (Cursor, an in-Obsidian AI panel, or a local script) **read** the
structure of a DOCX/PPTX file and **edit any part of it** through a stable, discoverable
API — with zero external installation. The only thing the user turns on is a setting:

- **Settings → Native PowerPoint Doc Editor → Enable AI-Interfacing** (default: off).

When off: no API surface, no manifest, no commands exposed. When on: the plugin exposes
the surfaces described below.

## Non-goals

- No separate MCP server the user has to install/run. Discovery happens through the plugin
  itself (in-process API + on-disk manifest + Obsidian commands).
- Agents never drive the UI (no toolbar clicks, no DOM selectors, no React internals).
- `describe()` never returns multi-MB base64 blobs; it returns structure. Raw bytes only
  on an explicit `export` request.

## Decisions locked in

| Topic | Decision |
|-------|----------|
| Consumer | Any AI extension; gated behind **Enable AI-Interfacing** setting |
| Formats | Both PPTX and DOCX (PPTX first, it is ~70% ready via `PresentationEngine`) |
| Transport | Zero-config: plugin API object + on-disk `capabilities.json` + JSON commands |
| Headless | Yes — edit any file by path, whether or not it is open in a tab |
| Op granularity | Fine-grained ops per feature; agent composes them for any requested edit |
| Safety | **Auto-apply as a single undo step**; `dryRun` preview available but not required |
| Validation | **Full round-trip validation on the real save**; light checks on `dryRun` |
| Docs | This file (`docs/AGENT-API.md`) + generated `capabilities.json` |

### Safety model (as chosen)

- Agent edits **apply immediately**.
- An entire `apply(path, ops[])` call collapses into **one undo step** (one Ctrl+Z reverts
  the whole batch).
- Any op may be sent with `dryRun: true`: the plugin computes the change and returns a
  diff-style preview **without writing**. Preview is optional, for agents that want to
  confirm before committing.
- No forced two-step confirm, no automatic backup copy (can be added later behind a
  setting if wanted).

### Validation model (as chosen)

- Real saves run the **full round-trip validation** (same as manual save): rebuild bytes,
  re-parse/re-open the whole file to prove it is not corrupt before writing to the vault.
- `dryRun` calls run only **light structural checks** (they never write, so the expensive
  round-trip is skipped).

## Architecture

```
Setting: Enable AI-Interfacing
        │  (gates everything below)
        ▼
   ┌─────────────────────────────────────────────┐
   │                 AI Core (src/ai/)            │
   │  - setting gate                              │
   │  - op registry + JSON-schema validation      │
   │  - structured results                        │
   │  - `agent` log area (apply traces)           │
   └───────────────┬──────────────┬───────────────┘
                   │              │
        ┌──────────▼───┐   ┌──────▼──────────┐
        │ PptxDocument │   │ DocxDocument     │
        │ Service      │   │ Service          │
        └──────┬───────┘   └──────┬───────────┘
               │                  │
        PresentationEngine   DOCX OOXML model
                                (new)
```

Surfaces exposed when the setting is on:

1. **Plugin API** — `app.plugins.plugins['native-powerpoint-doc-editor'].ai`
   - `getInfo()` → `{ apiVersion, pluginId, enabled }`
   - `listCapabilities()`
   - `validateOps(ops[])`
   - `describe(path)` → snapshot JSON
   - `apply(path, ops[], { dryRun? })` → result JSON
   - `openSession(path)` / `session.describe()` / `session.apply()` / `session.save()` / `session.close()`
2. **On-disk manifest** — `<vault>/.obsidian/plugins/native-powerpoint-doc-editor/ai/capabilities.json`
   - Generated from the op registry (like the i18n key generation step).
   - Lets an agent discover ops/schemas by reading a file, no code introspection needed.
   - Includes `surfaces.clipboardCommands` payload hints for command-only agents.
3. **Obsidian commands** (JSON in/out via clipboard) — for agents that can only trigger commands:
   - `npde-ai-capabilities` — copy live manifest (input ignored)
   - `npde-ai-describe` — `{ "path"?: "vault/file.pptx" }` → `DescribeResult` (omit `path` to use active file)
   - `npde-ai-apply` — `{ "path"?, "ops": DocumentOp[], "dryRun"?: boolean }` → `ApplyResult`
   - `npde-ai-validate` — `{ "ops": DocumentOp[] }` → `{ ok, errors }`
   - `npde-ai-save` — `{ "path"?: "vault/file.pptx" }` → `{ ok, errors }`
   - Legacy aliases remain registered: `npde-ai-list-capabilities`, `npde-ai-describe-document`, `npde-ai-apply-operations`.

## Stable IDs (the core contract)

Agents address content by stable IDs that survive reload:

- PPTX shape: `slide:<slideIndex>/shape:<shapeIndex>`
- PPTX paragraph/run: `slide:2/shape:3/p:0/r:1`
- DOCX block: `body/p[12]`, `body/tbl[3]/tr[1]/tc[2]`
- DOCX run: `body/p[12]/r[0]`

Rules:
- Negative PPTX shape indices are **inherited/synthetic** (master/layout placeholders) and
  are **not editable** — surfaced as `editable: false` in `describe()` (mirrors
  `isEditableShapeIndex`).
- IDs are 0-based and match the renderer's `data-ooxml-shape-idx`.

## Read model — `describe(path)`

### PPTX (example)

```json
{
  "format": "pptx",
  "file": "deck.pptx",
  "slideCount": 35,
  "slides": [
    {
      "index": 2,
      "shapes": [
        {
          "id": "slide:2/shape:3",
          "kind": "image",
          "editable": true,
          "text": null,
          "transform": { "x": 6972300, "y": 1257300, "cx": 5715000, "cy": 5715000, "rot": 0 },
          "style": null
        },
        {
          "id": "slide:2/shape:2",
          "kind": "textbox",
          "editable": true,
          "text": "Abstract Workshop",
          "paragraphs": [
            { "id": "slide:2/shape:2/p:0", "text": "Abstract Workshop", "align": "ctr" }
          ]
        }
      ]
    }
  ]
}
```

### DOCX (example)

```json
{
  "format": "docx",
  "file": "doc.docx",
  "blocks": [
    {
      "id": "body/p[12]",
      "kind": "paragraph",
      "style": "Heading1",
      "text": "Abstract",
      "runs": [ { "id": "body/p[12]/r[0]", "text": "Abstract", "bold": true } ]
    }
  ]
}
```

## Write model — `apply(path, ops[])`

Each op is a typed, schema-validated object. Result reports what changed.

```jsonc
// request
{
  "file": "deck.pptx",
  "dryRun": false,
  "ops": [
    { "op": "pptx.updateShapeText", "slideIndex": 2, "shapeIndex": 2, "text": "New title" },
    { "op": "pptx.updateTransform", "slideIndex": 2, "shapeIndex": 3,
      "transform": { "x": 7861300, "y": 1828800, "cx": 5715000, "cy": 5715000, "rot": 0 } }
  ]
}
```

```jsonc
// result
{
  "ok": true,
  "changed": ["slide:2/shape:2", "slide:2/shape:3"],
  "undoLabel": "AI edit",
  "warnings": [],
  "errors": []
}
```

```jsonc
// dryRun result (nothing written)
{
  "ok": true,
  "dryRun": true,
  "preview": [
    { "id": "slide:2/shape:2", "field": "text", "before": "Abstract Workshop", "after": "New title" },
    { "id": "slide:2/shape:3", "field": "transform.x", "before": 6972300, "after": 7861300 }
  ]
}
```

## Op coverage (map from existing feature areas)

Every diagnostics area in `main.ts` becomes an op namespace. Full coverage target so an
agent can perform any edit the user asks for by composing ops.

| Area (main.ts) | Example ops |
|----------------|-------------|
| PPTX text-editing | `pptx.updateShapeText`, `pptx.updateParagraphText`, `pptx.updateTextRun`, `pptx.replaceText` |
| PPTX text-formatting | `pptx.setRunStyle`, `pptx.setParagraphAlignment`, `pptx.applyListStyle` |
| PPTX inspector/arrange | `pptx.updateTransform`, `pptx.reorderShapes`, `pptx.groupShapes`, `pptx.ungroupShapes`, `pptx.flipShape` |
| PPTX insert | `pptx.addImage`, `pptx.addShape`, `pptx.addTextBox`, `pptx.addTable`, `pptx.addChart` |
| PPTX slide-operations | `pptx.addSlide`, `pptx.deleteSlide`, `pptx.moveSlide`, `pptx.duplicateSlide`, `pptx.reorderSlides`, `pptx.setSlideBackground` |
| PPTX image | `pptx.setImageCrop`, `pptx.resetImage`, `pptx.replaceImage` |
| PPTX charts | `pptx.updateChartData` |
| DOCX text/font | `docx.setRunText`, `docx.setRunStyle`, `docx.setParagraphStyle` |
| DOCX table | `docx.insertTable`, `docx.setCellText`, `docx.setCellStyle` |
| DOCX image | `docx.insertImage`, `docx.replaceImage` |
| DOCX find-replace | `docx.replaceText` |

## Headless sessions

- `openSession(path)` loads the file into a `PptxDocumentService` / `DocxDocumentService`
  without requiring a tab.
- The existing `NativePowerPointView` / `DocxView` become thin hosts over the same service,
  so UI edits and agent edits share one code path (no forked logic, same history + save).
- Concurrency: if a file is open in a view, the session reuses that engine to avoid two
  writers; otherwise it loads its own and releases on `close()`.

## PPTX vs DOCX status

| | PPTX | DOCX |
|---|------|------|
| Engine | `PresentationEngine` mostly ready | Eigenpal is DOM-bound |
| Read | Shape tree from engine + slide XML | **Done** — `word/document.xml` OOXML parser (`docxDescribe.ts`) |
| Write | Wrap existing engine methods | **Done (Phase 5)** — headless OOXML patch layer (`docxPatchSession`, `docxOpExecutor`, `docxSave`) |
| Priority | Phase 2 | Phase 4–5 **done** |

## Error codes

- `AI_DISABLED` — setting is off.
- `FILE_NOT_FOUND` / `UNSUPPORTED_FORMAT`
- `SHAPE_NOT_FOUND` / `BLOCK_NOT_FOUND`
- `OBJECT_NOT_EDITABLE` — negative/synthetic index or view-only file.
- `VALIDATION_FAILED` — round-trip validation rejected the rebuilt file (nothing written).
- `SCHEMA_INVALID` — op failed JSON-schema validation.

## Observability

- New `agent` log area records every `apply` (ops, changed IDs, dryRun, ms, validation
  outcome). Surfaces through the existing Copy PPTX/DOCX log flow.

## Phased implementation

1. **Done** — AI Core (`src/ai/`): setting gate, op registry, schema validation, `agent` log area,
   `capabilities.json` generation.
2. **Done** — `PptxDocumentService.describe()` + `apply()` + headless `save()` over `PresentationEngine`.
3. **Done** — Plugin API (`.ai`) + JSON commands + this doc.
4. **Done** — `DocxDocumentService.describe()` — structured OOXML read model from `word/document.xml`.
5. **Done** — DOCX write ops via OOXML patch layer.
6. **Done** — Per-op JSON Schemas + `example` payloads in `capabilities.json` (schema v2), DOCX round-trip save validation, `check:ai-capabilities`, `smoke:ai-agent`.

## Open items to revisit later

- Optional backup-before-edit setting (currently off per decision).
- Optional forced-preview mode per agent (currently dryRun is opt-in).
- Rate/size limits for very large batches.
