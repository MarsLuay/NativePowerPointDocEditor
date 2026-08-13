# Architecture

## DOCX editing flow

1. `src/main.ts` loads the plugin and registers the DOCX view.
2. `src/DocxView.tsx` owns the Obsidian leaf, toolbar, save/export lifecycle,
   and the DOCX document session.
3. `src/DocxReactView.tsx` owns the hidden ProseMirror editor and rendered-page
   interaction. The editor holds the canonical caret/selection; rendered pages
   provide the visible document surface.
4. `src/docxTextSelectionMemory.ts` preserves a non-empty selection across
   toolbar focus changes and clears it on editor input.
5. `src/docxFontSizeTarget.ts` resolves whether a font-size command targets the
   current selection, preserved selection, current paragraph, rendered
   paragraph, or an empty paragraph.
6. `src/docx/runtime/bridge.mjs` is the only `@npde/*` import boundary. The
   vendored runtime supplies ProseMirror commands, selection movement, and
   DOCX rendering.

## Important invariant

The visible rendered page and hidden editor must resolve to the same paragraph.
An empty paragraph is still a formatting target: its paragraph-level default
text formatting must be mutated rather than redirected to the preceding text
paragraph.

## Ownership boundary

- Keyboard navigation belongs to the runtime/editor selection layer.
- Toolbar font-size target resolution belongs to `DocxReactView.tsx` plus
  `docxFontSizeTarget.ts`.
- Selection preservation belongs to `docxTextSelectionMemory.ts` and its event
  wiring in `DocxReactView.tsx`.
- Serialization and autosave belong to the DOCX view/runtime save path.
- `main.js` and `ai/capabilities.json` are generated/deployed artifacts, not
  source owners.
