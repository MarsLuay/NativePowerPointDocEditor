# Module — src-docx-editor

**Paths:** `src/DocxView.tsx`, `src/DocxReactView.tsx`,
`src/docxFontSizeTarget.ts`, `src/docxTextSelectionMemory.ts`

**Purpose:** Open and edit DOCX files through a hidden ProseMirror editor while
painting a visible paginated document surface.

**Public surface:** DOCX Obsidian view lifecycle, rendered-page selection,
toolbar formatting, keyboard editing, save/export actions.

**Depends on:** Obsidian view APIs, ProseMirror, the DOCX runtime bridge,
project logger, and vault persistence.

**Invariants:** Stable paragraph identity survives visible-page clicks;
empty paragraphs remain addressable formatting targets; toolbar focus does not
silently retarget a command to the previous text paragraph; DOCX saves use the
current editor state.

**Related functions:** [functions/src-docx-editor.md](../functions/src-docx-editor.md)

**Related types:** [types/src-docx-editor.md](../types/src-docx-editor.md)
