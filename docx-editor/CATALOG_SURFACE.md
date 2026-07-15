# Catalog surface (public mirror)

This checkout ships **JS-only** `docx-editor/packages/{core,react,i18n}`
(`.js` / `.mjs` / `.cjs` / `.css`). No package `.d.ts` and no `types`
fields in those `package.json` files.

Full TypeScript sources and package typings live in the ObsidianNotes vault
authoritative tree only. Runtime JS is unchanged by sync.

Rebuild from vault: `npm run build:docx-editor` then re-sync.
