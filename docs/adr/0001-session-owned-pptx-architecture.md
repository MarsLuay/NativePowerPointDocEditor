---
status: accepted
---

# Session-owned PPTX architecture

`PresentationSession` owns PPTX domain state: selection, current slide, dirty/save state, and history. UI controllers only bind UI events to `session.applyCommand` and selection APIs; the lossless OOXML package is authoritative, the renderer is derived, and every edit uses one mutation-transaction path.

Vault persistence for DOCX and PPTX belongs to `DocumentSaveCoordinator`. Eigenpal remains behind `DocxEditorAdapter`, with a future test fake; we reject a forever host-adapter pattern and a permanent hybrid architecture because both preserve competing state and mutation owners.

## Considered options

- Forever host-adapter pattern — rejected.
- Permanent hybrid architecture — rejected.
