import * as vscode from 'vscode';
import { HarborFormulasPanel } from './panels/HarborFormulasPanel';

/**
 * This method is called when your extension is activated.
 * Your extension is activated the very first time the command is executed.
 */
export function activate(context: vscode.ExtensionContext) {
  // Use output channel for better visibility
  const outputChannel = vscode.window.createOutputChannel('Harbor Formulas');
  outputChannel.appendLine('=== Harbor Formulas Extension Activated ===');
  outputChannel.appendLine(`Extension URI: ${context.extensionUri.toString()}`);
  outputChannel.show();
  
  console.log('Harbor Formulas extension is now active!');
  console.log('Extension URI:', context.extensionUri.toString());

  // Register command to open Harbor Formulas
  const disposable = vscode.commands.registerCommand(
    'harborFormulas.openSpreadsheet',
    () => {
      console.log('Command triggered: harborFormulas.openSpreadsheet');
      HarborFormulasPanel.createOrShow(context.extensionUri);
    }
  );

  context.subscriptions.push(disposable);

  // Auto-open on activation - use setTimeout to ensure VS Code is ready
  outputChannel.appendLine('Auto-opening Harbor Formulas panel...');
  console.log('Auto-opening Harbor Formulas panel...');
  console.log('Extension URI:', context.extensionUri.toString());
  
  // Try immediate open first
  vscode.window.showInformationMessage('Harbor Formulas extension activated!');
  
  // Use requestAnimationFrame-like approach for webview
  const openPanel = () => {
    outputChannel.appendLine('Creating Harbor Formulas panel...');
    console.log('Creating Harbor Formulas panel...');
    try {
      HarborFormulasPanel.createOrShow(context.extensionUri);
      outputChannel.appendLine('Panel creation completed');
      console.log('Panel creation completed');
      vscode.window.showInformationMessage('Harbor Formulas panel opened!');
    } catch (error: any) {
      outputChannel.appendLine(`ERROR: ${error?.message || error}`);
      outputChannel.appendLine(`Stack: ${error?.stack || 'No stack trace'}`);
      console.error('Error creating panel:', error);
      console.error('Error stack:', error?.stack);
      vscode.window.showErrorMessage(`Failed to create panel: ${error?.message || error}`);
    }
  };
  
  // Try multiple times with increasing delays
  setTimeout(openPanel, 500);
  setTimeout(openPanel, 1500);
  setTimeout(openPanel, 3000);
}

/**
 * This method is called when your extension is deactivated.
 */
export function deactivate() {
  // Cleanup if needed
}

