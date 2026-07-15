# Catalog surface (public mirror)

This checkout ships **dist-only** `docx-editor/packages/{core,react,i18n}`.
Full TypeScript sources live in the ObsidianNotes vault authoritative tree.

Obsidian Community catalog ESLint scans public `.ts`/`.tsx`; unbundled
monorepo source fails that gate (style injection, static styles, etc.).

Rebuild from vault: `npm run build:docx-editor` then re-sync.
