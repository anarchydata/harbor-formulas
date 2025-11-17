import * as vscode from 'vscode';
import { HarborFormulasPanel } from './panels/HarborFormulasPanel';

/**
 * This method is called when your extension is activated.
 * Your extension is activated the very first time the command is executed.
 */
export function activate(context: vscode.ExtensionContext) {
  console.log('Harbor Formulas extension is now active!');

  // Register command to open Harbor Formulas
  const disposable = vscode.commands.registerCommand(
    'harborFormulas.openSpreadsheet',
    () => {
      HarborFormulasPanel.createOrShow(context.extensionUri);
    }
  );

  context.subscriptions.push(disposable);

  // Auto-open on activation - use setTimeout to ensure VS Code is ready
  setTimeout(() => {
    HarborFormulasPanel.createOrShow(context.extensionUri);
  }, 100);
}

/**
 * This method is called when your extension is deactivated.
 */
export function deactivate() {
  // Cleanup if needed
}

