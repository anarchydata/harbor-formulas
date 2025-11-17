# Harbor Formulas VS Code Extension

This is a VS Code extension that provides a spreadsheet application built on top of VS Code, similar to how Cursor is built on VS Code.

## Architecture

- **Extension Host**: TypeScript code running in VS Code's extension host
- **Webview**: HTML/CSS/JS running in a webview panel (isolated from VS Code)
- **Core Services**: HyperFormula (calculation engine) and Handsontable (grid UI)

## Development

1. Install dependencies:
   ```bash
   cd extension
   npm install
   ```

2. Compile TypeScript:
   ```bash
   npm run compile
   ```

3. Run extension:
   - Press F5 in VS Code
   - This opens a new Extension Development Host window
   - The extension will auto-activate and open the Harbor Formulas panel

## Future: Electron App

This extension structure is designed to be easily converted to a standalone Electron app, similar to Cursor. The webview content can be extracted and run in Electron's main window.

## Structure

```
extension/
├── src/
│   ├── extension.ts          # Extension entry point
│   └── panels/
│       └── HarborFormulasPanel.ts  # Webview panel manager
├── media/                    # Webview resources (HTML, CSS, JS)
├── out/                      # Compiled JavaScript
└── package.json              # Extension manifest

