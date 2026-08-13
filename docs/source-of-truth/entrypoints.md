# Entrypoints

| Surface | Source | Purpose |
| --- | --- | --- |
| Obsidian plugin load | `src/main.ts` | Registers plugin settings, views, commands, logger, and runtime loading. |
| DOCX view | `src/DocxView.tsx` | Opens a vault DOCX and owns the document session/lifecycle. |
| DOCX React editor | `src/DocxReactView.tsx` | Renders pages and connects visible interaction to the hidden editor. |
| AI bridge | `src/ai/pluginApi.ts`, `src/ai/aiCore.ts` | Exposes describe/apply/save/export operations to the NPDE AI interface. |
| Deployed plugin | `.obsidian/plugins/native-powerpoint-doc-editor/main.js` | Installed artifact currently loaded by Obsidian. |

## Focused user actions

- Arrow-key navigation moves the ProseMirror selection through document
  paragraphs, including empty paragraphs.
- Font-size controls invoke the editor formatting path and must preserve the
  selected empty paragraph as the target.
- Autosave serializes the editor state back to the vault DOCX.
