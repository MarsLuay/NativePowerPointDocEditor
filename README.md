# Native PowerPoint Doc Editor

Native PowerPoint Doc Editor is an Obsidian plugin for opening, searching, and editing `.docx` and `.pptx` files directly inside your vault.

The plugin keeps Office files in place instead of converting them to Markdown. It is designed for school, work, and research vaults where Word documents and PowerPoint decks need small edits, search, review, or quick inspection without leaving Obsidian.

| DOCX editor | PowerPoint editor |
| --- | --- |
| ![Native PowerPoint Doc Editor DOCX screen](screenshot.png) | ![Native PowerPoint Doc Editor PPTX screen](screenshot-pptx.png) |

## Features

- Open DOCX files in a native editor view
- Open PPTX files in a PowerPoint-style slide editor view
- Edit and save DOCX files back to the original vault file
- Edit PowerPoint text, tables, charts, shapes, slide objects, and chart data for supported `.pptx` decks
- Search inside DOCX files from Obsidian
- Search within opened PowerPoint decks
- Duplicate, export, and save-as supported documents
- Detect possible save conflicts when a file changes on disk while it is open
- Scan DOCX files for hidden or suspicious text
- Keep DOCX and PPTX handling optional so another plugin can take over those extensions

## Installation

### Community plugin directory

1. Open Obsidian Settings.
2. Go to Community plugins.
3. Search for `Native PowerPoint Doc Editor`.
4. Install and enable the plugin.

### Manual install or beta testing

1. Download the latest release assets from GitHub:
   - `main.js`
   - `manifest.json`
   - `styles.css`
2. Create this folder in your vault:

   ```text
   .obsidian/plugins/native-powerpoint-doc-editor
   ```

3. Copy the release files into that folder.
4. Reload Obsidian and enable `Native PowerPoint Doc Editor` from Community plugins.

The `run-to-import` folder also contains local Windows and macOS installers for manual vault installation.

## Usage

- Open a `.docx` file in the file explorer to use the DOCX editor.
- Open a `.pptx`, `.pptm`, `.ppsx`, `.ppsm`, `.potx`, or `.potm` file to use the PowerPoint view.
- Use the toolbar and command palette actions for save, export, duplicate, search, and document diagnostics.
- Use plugin settings to turn DOCX or PowerPoint handling on or off.

## Privacy and network

Native PowerPoint Doc Editor is designed to work **offline inside your vault**.

- **No telemetry or analytics.** The plugin does not phone home, collect usage data, or require an account.
- **No self-updating code.** Updates come only through Obsidian's normal community-plugin install flow or manual release files you place in your vault.
- **No ads.** The plugin does not load advertising scripts or remote UI content.

### When network access can happen

Most use is fully local. Network access is limited to cases below:

- **Links you choose to open.** Settings may show outbound links (for example presentation template sites, GitHub issue reporting, or the Obsidian download page shown when an older browser engine cannot render PPTX). Opening those links is handled by your browser or operating system.
- **Links inside your documents.** If a DOCX or PPTX file contains a hyperlink and you activate it, Obsidian or your OS handles that navigation.
- **DOCX PDF export with remote images.** If a document references images hosted on the web, export may ask the browser to load those image URLs so they can be rendered into the PDF.

The plugin does not upload vault contents to a server as part of normal editing, search, save, or export.

### Local data the plugin stores

- **DOCX search index** — optional, **off by default**. When enabled in settings, builds a local text cache at `.obsidian/plugins/native-powerpoint-doc-editor/docx-search-index.json` for vault-wide DOCX search.
- **Recovery copies** — if a PowerPoint save fails or autosave is disabled, unsaved edits may be written to a new file in your vault.
- **Plugin settings** — stored in Obsidian's normal plugin data file for this plugin.
- **Debug log file** — only when a developer `.hotreload` marker is present in the plugin folder; otherwise logging stays in memory and the developer console.

### Files outside the vault

- **Import font** and **insert image** actions use the system file picker. The file you choose may live outside the vault; only the content needed for the document is used.
- **Copy debug log** copies diagnostics you request to the clipboard. It does not send them anywhere automatically.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for release history, including the WasmGC fallback added in 1.0.14.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines, local setup notes, and release expectations.

## License

Released under the MIT license.
