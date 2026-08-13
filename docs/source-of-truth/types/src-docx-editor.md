# Types — src-docx-editor

| Type | File | Shape / role |
| --- | --- | --- |
| `DocxTextRange` | `src/docxFontSizeTarget.ts` | `{ from: number; to: number }`, a ProseMirror document range. |
| `DocxFontSizeTargetSource` | `src/docxFontSizeTarget.ts` | `selection`, `preserved-selection`, `caret-paragraph`, `rendered-paragraph`, or `empty-paragraph`. |
| `ResolveDocxFontSizeTargetInput` | `src/docxFontSizeTarget.ts` | Current selection plus optional preserved, caret-paragraph, rendered-paragraph, and preference inputs. |
| `ResolvedDocxFontSizeTarget` | `src/docxFontSizeTarget.ts` | `{ range, source }`, the resolved formatting target. |
| `DocxTextSelectionRange` | `src/docxTextSelectionMemory.ts` | Preserved non-empty selection range. |
| `DocxTextSelectionEventContext` | `src/docxTextSelectionMemory.ts` | Event type plus whether the event occurred in rendered pages or hidden ProseMirror. |
