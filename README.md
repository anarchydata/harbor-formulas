# Harbor Formulas

A spreadsheet application built on VS Code, similar to how Cursor is built on VS Code. Eventually will become a standalone Electron app.

## Project Structure

```
harbor-formulas/
├── packages/
│   └── harbor-core/          # Shared TypeScript logic
│       ├── src/
│       │   ├── services/      # Core services (HyperFormula, Grid, Sheet)
│       │   ├── BaseService.ts
│       │   ├── BaseController.ts
│       │   ├── grid/          # Grid adapters
│       │   ├── utils/         # Utility functions
│       │   └── types/         # Type definitions
│       └── package.json
├── apps/
│   └── harbor-vscode/        # VS Code extension (webview inside)
│       ├── src/
│       │   ├── extension.ts
│       │   └── panels/
│       ├── media/             # Webview resources
│       └── package.json
└── package.json               # Root workspace config
```

## Development

### Setup

```bash
npm install
```

### Build

```bash
# Build all packages
npm run build

# Build specific package
npm run build:core
npm run build:vscode
```

### Run VS Code Extension

1. Open the project in VS Code
2. Go to `apps/harbor-vscode/`
3. Press F5 to launch Extension Development Host
4. The extension will auto-activate and open Harbor Formulas panel

## Architecture

- **packages/harbor-core**: Shared TypeScript logic that can be used by any app (VS Code extension, future Electron app, etc.)
- **apps/harbor-vscode**: VS Code extension that uses a webview to display the spreadsheet UI

## Future: Electron App

The structure is designed to easily extract the webview content and run it in Electron's main window, similar to Cursor.

