import * as vscode from 'vscode';
import * as path from 'path';

/**
 * Manages the Harbor Formulas webview panel
 */
export class HarborFormulasPanel {
  public static currentPanel: HarborFormulasPanel | undefined;
  public static readonly viewType = 'harborFormulas';

  private readonly _panel: vscode.WebviewPanel;
  private readonly _extensionUri: vscode.Uri;
  private _disposables: vscode.Disposable[] = [];

  public static createOrShow(extensionUri: vscode.Uri) {
    const column = vscode.window.activeTextEditor
      ? vscode.window.activeTextEditor.viewColumn
      : undefined;

    // If we already have a panel, show it
    if (HarborFormulasPanel.currentPanel) {
      HarborFormulasPanel.currentPanel._panel.reveal(column);
      return;
    }

    // Otherwise, create a new panel
    const panel = vscode.window.createWebviewPanel(
      HarborFormulasPanel.viewType,
      'Harbor Formulas',
      column || vscode.ViewColumn.One,
      {
        enableScripts: true,
        localResourceRoots: [
          vscode.Uri.joinPath(extensionUri, 'media'),
          vscode.Uri.joinPath(extensionUri, 'out'),
        ],
      }
    );

    HarborFormulasPanel.currentPanel = new HarborFormulasPanel(panel, extensionUri);
  }

  private constructor(panel: vscode.WebviewPanel, extensionUri: vscode.Uri) {
    this._panel = panel;
    this._extensionUri = extensionUri;

    // Set the webview's initial html content
    this._update();

    // Listen for when the panel is disposed
    // This happens when the user closes the panel or when the panel is closed programmatically
    this._panel.onDidDispose(() => this.dispose(), null, this._disposables);

    // Handle messages from the webview
    this._panel.webview.onDidReceiveMessage(
      (message) => {
        switch (message.command) {
          case 'alert':
            vscode.window.showErrorMessage(message.text);
            return;
          case 'error':
            vscode.window.showErrorMessage(`Harbor Formulas: ${message.text}`);
            return;
          case 'initialized':
            console.log('Harbor Formulas webview initialized');
            return;
        }
      },
      null,
      this._disposables
    );
  }

  public dispose() {
    HarborFormulasPanel.currentPanel = undefined;

    // Clean up our resources
    this._panel.dispose();

    while (this._disposables.length) {
      const x = this._disposables.pop();
      if (x) {
        x.dispose();
      }
    }
  }

  private _update() {
    const webview = this._panel.webview;
    this._panel.webview.html = this._getHtmlForWebview(webview);
  }

  private _getHtmlForWebview(webview: vscode.Webview) {
    // Get paths to resources
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this._extensionUri, 'media', 'bundle.js')
    );
    const bundleCssUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this._extensionUri, 'media', 'bundle.css')
    );
    const styleUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this._extensionUri, 'media', 'main.css')
    );

    // Use a nonce to only allow specific scripts to be run
    const nonce = getNonce();
    
    // Debug: log URIs
    console.log('Webview URIs:', {
      scriptUri: scriptUri.toString(),
      bundleCssUri: bundleCssUri.toString(),
      styleUri: styleUri.toString(),
      extensionUri: this._extensionUri.toString()
    });

    return `<!DOCTYPE html>
			<html lang="en">
			<head>
				<meta charset="UTF-8">
				<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}' ${webview.cspSource};">
				<meta name="viewport" content="width=device-width, initial-scale=1.0">
				<link href="${bundleCssUri}" rel="stylesheet">
				<link href="${styleUri}" rel="stylesheet">
				<title>Harbor Formulas</title>
			</head>
			<body>
				<div class="workbench">
					<div class="grid-container">
						<div class="grid-wrapper">
							<div id="handsontableRoot" style="width: 100%; height: 100%; min-height: 400px; background: #1e1e1e;">
								<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #cccccc;">
									<div style="text-align: center;">
										<h2>Harbor Formulas</h2>
										<p>Loading spreadsheet engine...</p>
										<p style="font-size: 11px; color: #858585;">If this message persists, check the Developer Tools console (Help > Toggle Developer Tools)</p>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
				<script nonce="${nonce}">
					const vscode = acquireVsCodeApi();
					window.vscode = vscode;
					console.log('VS Code API acquired');
					console.log('Loading bundle from:', '${scriptUri}');
				</script>
				<script nonce="${nonce}" src="${scriptUri}" onerror="console.error('Failed to load bundle.js:', event)" onload="console.log('bundle.js loaded successfully')"></script>
			</body>
			</html>`;
  }
}

function getNonce() {
  let text = '';
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 32; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

