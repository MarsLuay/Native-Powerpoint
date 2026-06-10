# Native PowerPoint/Doc Editor

Native PowerPoint/Doc Editor opens `.docx` and `.pptx` files directly inside Obsidian. It can edit DOCX files, view and edit PPTX files, search across DOCX text, and keep the original vault files in place.

It is not reviewed by Obsidian community plugin moderators. I made it for my own school/work document workflow, so use it carefully with important files.

## What this thing does

- Opens DOCX files in a native editor view
- Opens PPTX files in a native PowerPoint-style view
- Saves DOCX edits back to the original vault file
- Supports Save as, Duplicate, and Export as
- Adds conflict detection when a DOCX changes on disk while open
- Can insert images into DOCX files from the top Insert menu
- Adds File, Format, Insert, and Search menu controls that match the editor UI
- Lets you scan a DOCX for hidden or suspicious prompt-injection-style text
- Indexes DOCX text so vault-wide DOCX search can work
- Lets you turn off DOCX or PPTX handling when another plugin should take over

## Easiest way to install it in this vault

### Windows

Double click the .bat in run-to-import

### Mac

Double click the .app in run-to-import

### Linux

Nothing yet

## Working on it

```bash
cd "Projects/Native PowerPoint Doc Editor"
npm install
npm run build
npm run lint
npm run smoke
npm run smoke:chart-data
npm run smoke:halos
npm run smoke:objects
```

The smoke test loads the installed plugin copy and checks that DOCX/PPTX registration and diagnostics still work.
The PowerPoint chart/table smoke tests generate local fixture decks under `test-results/native-powerpoint-fixtures`, so they run without checked-in sample decks. You can still pass explicit deck paths with the `NATIVE_POWERPOINT_*_SAMPLE` environment variables when testing real files.

## Useful paths

- Source: `Projects/Native PowerPoint Doc Editor`
- Installed plugin: `.obsidian/plugins/native-powerpoint-doc-editor`
- Manifest id: `native-powerpoint-doc-editor`
- Display name: `Native PowerPoint/Doc Editor`

## Background

This started as a DOCX plugin, then got merged with my PowerPoint plugin so school and work files can open in Obsidian without bouncing into Word, Pages, or PowerPoint for every small edit.

The DOCX editor is based on Eigenpal's DOCX editor packages, with local Obsidian-specific glue and UI polish on top.
