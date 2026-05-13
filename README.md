# Word Pattern Formatter

Word Pattern Formatter is a Microsoft Word task pane add-in for saving, editing, importing, exporting, and applying reusable text formatting patterns.

The add-in is built with Office.js, TypeScript, Webpack, and the Microsoft Office Add-ins tooling. The production build is published from the `docs` folder for GitHub Pages.

## Features

- Save reusable Word formatting patterns.
- Apply a pattern to selected text, the current paragraph, or the whole document.
- Edit, duplicate, delete, import, and export saved patterns.
- Restore default formatting presets.
- Search and sort saved patterns.
- Preview saved patterns before applying them.
- Store patterns locally in the browser/task pane storage.
- Show toast notifications after user actions.

Saved patterns can include:

- Font family and font size.
- Alignment.
- Line spacing.
- Bold, italic, and underline.
- Text color and highlight color.
- Paragraph spacing.
- Left, right, and first-line indentation.

## Project Structure

```text
.
|-- docs/
|   `-- Production build used by GitHub Pages
|-- WordPatternFormatter/
|   |-- assets/
|   |-- src/
|   |   |-- commands/
|   |   `-- taskpane/
|   |       |-- taskpane.html
|   |       |-- taskpane.css
|   |       `-- taskpane.ts
|   |-- manifest.xml
|   |-- package.json
|   `-- webpack.config.js
`-- README.md
```

## Requirements

- Node.js and npm.
- Microsoft Word desktop or Word on the web.
- A Microsoft 365 account for local debugging with Office Add-ins tooling.

## Installation

Install dependencies from the add-in project folder:

```bash
cd WordPatternFormatter
npm install
```

## Local Development

Start the local HTTPS dev server:

```bash
npm run dev-server
```

In another terminal, start debugging the add-in in Word:

```bash
npm run start
```

The development manifest uses:

```text
https://localhost:3000/taskpane.html
```

To stop the debugging session:

```bash
npm run stop
```

## Build

Create a production build:

```bash
cd WordPatternFormatter
npm run build
```

The generated files are written to:

```text
WordPatternFormatter/dist
```

## Validate

Check linting:

```bash
npm run lint
```

Validate the Office add-in manifest:

```bash
npm run validate
```

## Publish to GitHub Pages

This repository uses the root-level `docs` folder as the GitHub Pages output.

After building, copy the generated files into `docs`:

```powershell
Copy-Item -Path WordPatternFormatter\dist\* -Destination docs -Recurse -Force
```

Then commit and push the changes:

```bash
git add .
git commit -m "Update Word Pattern Formatter"
git push
```

GitHub Pages should then serve the updated add-in from:

```text
https://vasilelin.github.io/word_plugin/
```

## Production Manifest Notes

During production builds, `webpack.config.js` replaces local development URLs with the GitHub Pages URL:

```text
https://vasilelin.github.io/word_plugin/
```

The generated HTML also includes cache-busting query strings for JavaScript assets, and the task pane CSS is versioned to help Word load the latest files after deployment.

If Word still shows an older version after deployment, close and reopen the add-in task pane or restart Word.

## Pattern Storage

Patterns are saved in local storage under:

```text
word_text_patterns
```

This means saved patterns are local to the current browser or Office WebView profile. Use Export JSON and Import JSON to move patterns between environments.

## Useful Scripts

Run these commands inside `WordPatternFormatter`:

```bash
npm run build       # Build for production
npm run build:dev   # Build for development
npm run dev-server  # Start local HTTPS dev server
npm run start       # Sideload and debug in Word
npm run stop        # Stop debugging
npm run lint        # Run Office add-in lint checks
npm run lint:fix    # Auto-fix lint issues where possible
npm run validate    # Validate manifest.xml
npm run watch       # Rebuild on file changes
```
