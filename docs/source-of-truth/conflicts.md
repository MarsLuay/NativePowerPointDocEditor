# Open conflicts and duplicates

Last updated: 2026-08-10

| ID | Kind | Severity | Symbols / paths | Evidence | Canonical owner (proposed) | Status |
| --- | --- | --- | --- | --- | --- | --- |
| C001 | ownership seam | block | Hidden ProseMirror selection vs rendered DOCX page selection | Empty-paragraph sizing now invalidates relayout signatures, and visual navigation now invalidates cached line state after external selection changes. Focused runtime tests and plugin tests pass. | Keep ProseMirror as the selection authority; rendered-page mapping and runtime keyboard movement are explicit and testable. | fixed |
