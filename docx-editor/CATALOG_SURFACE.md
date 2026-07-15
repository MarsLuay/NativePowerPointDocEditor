# Catalog surface (public mirror)

This checkout ships **dist-only** `docx-editor/packages/{core,react,i18n}`.
Full TypeScript sources live in the ObsidianNotes vault authoritative tree.

Package `*.d.ts` are sanitized for Obsidian catalog ESLint (`globalThis`,
Identifier `document`, `#private` stubs, etc.). Runtime JS is unchanged.

Rebuild from vault: `npm run build:docx-editor` then re-sync.
