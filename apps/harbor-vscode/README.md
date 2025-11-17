# Harbor Formulas VS Code Extension

This is a VS Code extension that provides a spreadsheet application built on top of VS Code, similar to how Cursor is built on VS Code.

## Architecture

- **Extension Host**: TypeScript code running in VS Code's extension host
- **Webview**: HTML/CSS/JS running in a webview panel (isolated from VS Code)
- **Core Services**: HyperFormula (calculation engine) and Handsontable (grid UI)

## Development

1. Install dependencies:
   ```bash
   npm install
   ```

2. Compile TypeScript:
   ```bash
   npm run compile
   ```

3. Bundle webview dependencies:
   ```bash
   npm run bundle
   ```

4. **Run extension**:
   - Open this folder (`apps/harbor-vscode/`) in VS Code
   - Press **F5** (or go to Run > Start Debugging)
   - This opens a new "Extension Development Host" window
   - The extension will auto-activate and open the Harbor Formulas panel
   - You'll see the spreadsheet grid in the webview

## Testing the Extension

**You MUST use VS Code to test VS Code extensions** - this is how the VS Code extension API works:

1. Open `apps/harbor-vscode/` in VS Code
2. Press **F5** to launch Extension Development Host
3. A new VS Code window opens with your extension loaded
4. The Harbor Formulas panel should appear automatically

## Build Scripts

- `npm run compile` - Compile TypeScript extension code
- `npm run bundle` - Bundle webview JavaScript (HyperFormula + Handsontable)
- `npm run bundle:watch` - Watch mode for bundling
- `npm run watch` - Watch mode for TypeScript compilation

## Future: Electron App

This extension structure is designed to be easily converted to a standalone Electron app, similar to Cursor. The webview content can be extracted and run in Electron's main window.

## Structure

```
apps/harbor-vscode/
├── src/
│   ├── extension.ts          # Extension entry point
│   └── panels/
│       └── HarborFormulasPanel.ts  # Webview panel manager
├── media/                    # Webview resources
│   ├── main.js               # Source (bundled to bundle.js)
│   ├── bundle.js             # Bundled with HyperFormula + Handsontable
│   ├── bundle.css            # Bundled CSS
│   └── main.css              # Custom styles
├── out/                      # Compiled JavaScript
└── package.json              # Extension manifest
