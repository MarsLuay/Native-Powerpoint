# Native PowerPoint Doc Editor

Native PowerPoint Doc Editor is an Obsidian plugin for opening, searching, and editing `.docx` and `.pptx` files directly inside your vault.

The plugin keeps Office files in place instead of converting them to Markdown. It is designed for school, work, and research vaults where Word documents and PowerPoint decks need small edits, search, review, or quick inspection without leaving Obsidian.

![Native PowerPoint Doc Editor](screenshot.png)

A DOCX file open in the editor view, with the formatting toolbar and the document rendered inline in Obsidian.

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

## Safety

Native PowerPoint Doc Editor edits binary Office files. Keep backups of important documents, especially before making large changes to complex decks or documents. The plugin includes export validation and conflict checks, but Office file formats are broad and some advanced content may remain view-only.

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

## Desktop Support

This plugin is desktop-only. It uses desktop Obsidian capabilities for Office editing, clipboard fidelity, local diagnostics, and larger binary file workflows.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines, local setup notes, and release expectations.

```bash
npm install
npm run build
npm run lint
npm run smoke
npm run smoke:chart-data
npm run smoke:generated-text
npm run smoke:halos
npm run smoke:objects
npm run smoke:fonts
```

PowerPoint smoke tests generate local fixture decks under `test-results/native-powerpoint-fixtures`, so they can run without checked-in sample decks. Real PPTX files can be passed with the `NATIVE_POWERPOINT_*_SAMPLE` environment variables where supported.

## License

Released under the 0BSD license.
