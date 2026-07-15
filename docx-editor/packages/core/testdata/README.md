# Core package test fixtures

DOCX (and related) files for `@npde/docx-editor-core` unit tests under `src/**`.

Formerly lived at `e2e/fixtures/` (Playwright). Playwright e2e was removed from this Obsidian-trimmed tree; fixtures used by package tests moved here.

- Top level: fixtures imported by tests in `packages/core/src/**`
- `manual/`: issue/layout samples not currently wired to a unit test (kept for local repro / regen scripts)

Resolve from tests with paths relative to the test file, e.g. `join(import.meta.dir, '../../../testdata/<file>.docx')`.
