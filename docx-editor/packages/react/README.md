<p align="center">
  <a href="https://www.docx-editor.dev/">
    <img src="https://raw.githubusercontent.com/MarsLuay/NativePowerPointDocEditor/main/.github/assets/header.png" alt="DOCX Editor — .docx in, .docx out. Open source, agent ready, client-side." width="500" />
  </a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/@npde/docx-editor-react"><img src="https://img.shields.io/npm/v/@npde/docx-editor-react.svg?style=flat-square&color=3B5BDB" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/@npde/docx-editor-react"><img src="https://img.shields.io/npm/dm/@npde/docx-editor-react.svg?style=flat-square&color=3B5BDB" alt="npm downloads" /></a>
  <a href="https://github.com/MarsLuay/NativePowerPointDocEditor/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-Apache_2.0-blue.svg?style=flat-square&color=3B5BDB" alt="license" /></a>
  <a href="https://docx-editor.dev/editor"><img src="https://img.shields.io/badge/Live_Demo-3B5BDB?style=flat-square&logo=vercel&logoColor=white" alt="Demo" /></a>
  <a href="https://www.docx-editor.dev/docs"><img src="https://img.shields.io/badge/Docs-3B5BDB?style=flat-square&logo=readthedocs&logoColor=white" alt="Documentation" /></a>
</p>

# @npde/docx-editor-react

React adapter for the [docx-editor](https://docx-editor.dev). WYSIWYG `.docx` editing with canonical OOXML, tracked changes, comments, real-time collaboration, and an AI agent bridge.

## Quick Start

```bash
npm install @npde/docx-editor-react
```

```tsx
import { useState } from 'react';
import { DocxEditor } from '@npde/docx-editor-react';
import '@npde/docx-editor-react/styles.css';

export function App() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);

  return (
    <>
      <input
        type="file"
        accept=".docx"
        onChange={async (e) => setBuffer((await e.target.files?.[0]?.arrayBuffer()) ?? null)}
      />
      {buffer && <DocxEditor documentBuffer={buffer} mode="editing" />}
    </>
  );
}
```

> **Next.js / SSR:** Use dynamic import. The editor requires the DOM.

## Start with a blank document

Skip the file picker for new documents. `createEmptyDocument` returns a fresh `Document` model you can pass straight to the editor:

```tsx
import { DocxEditor, createEmptyDocument } from '@npde/docx-editor-react';
import '@npde/docx-editor-react/styles.css';

const doc = createEmptyDocument();
// Or with options:
// createEmptyDocument({ initialText: 'Untitled', pageWidth: 12240 })

<DocxEditor document={doc} mode="editing" />;
```

`createDocumentWithText(text, options?)` is the same idea with a starting paragraph already typed. Both helpers are re-exported from `@npde/docx-editor-core` so you don't need a separate dependency.

## Customize File > Open

By default, the built-in `File > Open` item and Cmd/Ctrl+O prompt for a `.docx` file and load it into the local editor view. Pass `onOpen` to keep the native file picker but route the selected `File` through your own import pipeline instead:

```tsx
<DocxEditor
  document={doc}
  externalContent
  externalPlugins={plugins}
  onOpen={async (file) => {
    await importIntoBackend(file);
  }}
/>
```

Set `showFileOpen={false}` to hide the built-in Open item and leave Cmd/Ctrl+O for your own menu or toolbar.

## Packages

| Package                                                                                      | Description                                                                                                                                |
| -------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| [`@npde/docx-editor-react`](https://www.npmjs.com/package/@npde/docx-editor-react)   | <img src="https://cdn.simpleicons.org/react/61DAFB" width="20" align="middle" /> &nbsp; React adapter. Toolbar, paged editor, plugins.     |
| [`@npde/docx-editor-vue`](https://www.npmjs.com/package/@npde/docx-editor-vue)       | <img src="https://cdn.simpleicons.org/vuedotjs/4FC08D" width="20" align="middle" /> &nbsp; Vue 3 adapter. Toolbar, paged editor, plugins.  |
| [`@npde/docx-editor-core`](https://www.npmjs.com/package/@npde/docx-editor-core)     | Framework-agnostic core: OOXML parser, serializer, layout engine, ProseMirror schema. Depend on this if you fork the React or Vue adapter. |
| [`@npde/docx-editor-i18n`](https://www.npmjs.com/package/@npde/docx-editor-i18n)     | Shared locale strings and types consumed by both adapters.                                                                                 |

> **Forking the adapter?** Keep your fork thin. Depend on `@npde/docx-editor-core` directly so parser, serializer, and rendering fixes land in your build automatically, without backporting each upstream change by hand.

## Imperative mounting

```ts
import { renderAsync } from '@npde/docx-editor-react';

const editor = await renderAsync(file, document.getElementById('editor')!, { mode: 'editing' });
await editor.save();
editor.destroy();
```

## Subpaths

- `@npde/docx-editor-react` — `DocxEditor`, `renderAsync`, public types
- `@npde/docx-editor-react/ui` — toolbar primitives, pickers, sidebars, dialogs
- `@npde/docx-editor-react/hooks` — `useAutoSave`, `useTableSelection`, ...
- `@npde/docx-editor-react/dialogs` — dialog components barrel
- `@npde/docx-editor-react/plugin-api` — plugin host and plugin-facing types
- `@npde/docx-editor-react/styles` — style constants (`EDITOR_CSS_PATH`, z-index)

## Plugins

```tsx
import { DocxEditor } from '@npde/docx-editor-react';
import { PluginHost, templatePlugin } from '@npde/docx-editor-react/plugin-api';

<PluginHost plugins={[templatePlugin]}>
  <DocxEditor documentBuffer={buffer} />
</PluginHost>;
```

## Component API

Full props and ref reference: **[docx-editor.dev/docs/props](https://www.docx-editor.dev/docs/props)**. `DocxEditor` and `DocxEditorRef` mirror the Vue adapter, so docs apply with just the import path swapped.

`@npde/docx-editor-core` is installed transitively. Add it to your `package.json` only if your own code imports core APIs directly. Strict installers like pnpm with peer auto-install disabled may also need the ProseMirror peers listed in `package.json`.

Examples: [Vite](https://github.com/MarsLuay/NativePowerPointDocEditor/tree/main/examples/vite) · [Next.js](https://github.com/MarsLuay/NativePowerPointDocEditor/tree/main/examples/nextjs) · [Remix](https://github.com/MarsLuay/NativePowerPointDocEditor/tree/main/examples/remix) · [Astro](https://github.com/MarsLuay/NativePowerPointDocEditor/tree/main/examples/astro)

## Contributing

Contributions welcome. See [CONTRIBUTING.md](https://github.com/MarsLuay/NativePowerPointDocEditor/blob/main/CONTRIBUTING.md) for setup, tests, and the one-time CLA signature.

## Commercial Support

> [!TIP]
> Questions or custom features? Email **[noreply@users.noreply.github.com](mailto:noreply@users.noreply.github.com)**.
