# Contributing to Native PowerPoint Doc Editor

Thanks for helping improve Native PowerPoint Doc Editor. This plugin is maintained by Mars and is built to make DOCX and PowerPoint files usable directly inside Obsidian.

## Before You Start

- Open an issue for larger changes so the scope is clear before implementation.
- Keep pull requests focused on one feature, fix, or cleanup at a time.
- Avoid committing generated artifacts such as `main.js`, `test-results/`, or `scripts/visual-output/`.
- Do not include private vault content, documents, presentations, screenshots, or logs unless they are intentionally sanitized.

## Local Setup

Install dependencies:

```bash
npm install
```

Run a production build:

```bash
npm run build
```

Run the core checks:

```bash
npm run lint
npm run smoke
```

Run focused PowerPoint and DOCX checks when touching rendering, editing, clipboard, chart, font, or export code:

```bash
npm run smoke:generated-text
npm run smoke:chart-data
npm run smoke:objects
npm run smoke:fonts
npm run visual:caret
```

## Development Notes

- Prefer Obsidian APIs and browser-safe DOM patterns over Node-only APIs.
- Use `activeDocument` for DOM access that should work in Obsidian popout windows.
- Use cross-window-safe element checks for DOM nodes created outside the main window.
- Keep CSS compatible with the minimum Obsidian version in `manifest.json`.
- Keep user-facing text concise and consistent with the existing plugin wording.
- Preserve file safety behavior: conflict checks, validation, recovery copies, and view-only fallbacks should stay conservative.

## Testing Changes

For narrow documentation-only changes, a build is not usually required.

For source, styling, or dependency changes, run:

```bash
npm run lint
npm run build
npm run smoke
```

Then run the focused smoke commands that match the files you changed.

## Release Notes

Release notes should explain user-visible changes and review fixes clearly. Community plugin release assets should only include:

- `main.js`
- `manifest.json`
- `styles.css`

## Reporting Security or Data-Safety Issues

If you find a bug that could corrupt files, expose private vault data, or bypass Obsidian security expectations, open a minimal issue without private samples. Share sensitive reproduction files only after they have been sanitized.
