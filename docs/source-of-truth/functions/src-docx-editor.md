# Functions — src-docx-editor

Source roots: `src/DocxReactView.tsx`, `src/docxFontSizeTarget.ts`,
`src/docxTextSelectionMemory.ts`

| Symbol | File | Signature | Purpose | Side effects / errors |
| --- | --- | --- | --- | --- |
| `resolveDocxFontSizeTarget` | `src/docxFontSizeTarget.ts` | `(input: ResolveDocxFontSizeTargetInput) => ResolvedDocxFontSizeTarget` | Chooses the range for a font-size command, including collapsed empty-paragraph targets. | Returns a range/source pair; does not mutate the editor. |
| `resolveDocxFontSizeStepBase` | `src/docxFontSizeTarget.ts` | `(selectionFontSizePoints: readonly number[], controlFontSizePoints: number) => number` | Chooses the displayed step base for font-size increment/decrement controls. | Filters non-finite values. |
| `prepareFontSizeSelection` | `src/DocxReactView.tsx` | `(view, preservedRange, renderedParagraphRange, preferCurrentSelection) => PreparedFontSizeSelection` | Maps the current UI state to an editor selection before formatting. | Dispatches a selection transaction and seeds empty-paragraph marks. |
| `seedEmptyParagraphStoredMarks` | `src/DocxReactView.tsx` | `(view: EditorView) => void` | Carries paragraph defaults into stored marks for empty paragraphs. | Dispatches stored marks only for an empty paragraph with defaults. |
| `applyFontFamilyToEditorView` | `src/DocxReactView.tsx` | `(view, fontFamily, preservedRange) => { applied, range }` | Applies font-family formatting to text or empty paragraphs. | Dispatches formatting and logs empty-paragraph results. |
| `updateDocxPreservedTextSelection` | `src/docxTextSelectionMemory.ts` | `(previous, current, resetForEditorInput) => DocxTextSelectionRange | null` | Maintains or clears selection memory across focus transitions. | Clears on configured editor input. |
| `shouldResetDocxPreservedTextSelection` | `src/docxTextSelectionMemory.ts` | `(context: DocxTextSelectionEventContext) => boolean` | Decides whether an event invalidates preserved selection. | Treats hidden-editor keydown as reset input. |
