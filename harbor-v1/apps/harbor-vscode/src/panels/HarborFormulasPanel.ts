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
    const outputChannel = vscode.window.createOutputChannel('Harbor Formulas');
    outputChannel.appendLine('HarborFormulasPanel.createOrShow called');
    outputChannel.appendLine(`Extension URI: ${extensionUri.toString()}`);
    console.log('HarborFormulasPanel.createOrShow called');
    console.log('Extension URI:', extensionUri.toString());
    
    const column = vscode.window.activeTextEditor
      ? vscode.window.activeTextEditor.viewColumn
      : undefined;

    // If we already have a panel, show it
    if (HarborFormulasPanel.currentPanel) {
      outputChannel.appendLine('Panel already exists, revealing it');
      console.log('Panel already exists, revealing it');
      HarborFormulasPanel.currentPanel._panel.reveal(column);
      return;
    }

    outputChannel.appendLine('Creating new webview panel...');
    console.log('Creating new webview panel...');
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
        retainContextWhenHidden: true,
      }
    );

    outputChannel.appendLine('Panel created, initializing...');
    console.log('Panel created, initializing...');
    HarborFormulasPanel.currentPanel = new HarborFormulasPanel(panel, extensionUri);
    outputChannel.appendLine('Panel initialized successfully');
    console.log('Panel initialized successfully');
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
    
    // Add cache busting to script URI
    const scriptUriWithCache = `${scriptUri.toString()}?v=${Date.now()}`;
    
    // Debug: log URIs
    console.log('Webview URIs:', {
      scriptUri: scriptUriWithCache,
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
								<div id="statusMessage" style="display: flex; align-items: center; justify-content: center; height: 100%; color: #cccccc; flex-direction: column; gap: 10px;">
									<div style="text-align: center;">
										<h2>Harbor Formulas</h2>
										<p id="statusText">Initializing...</p>
										<p style="font-size: 11px; color: #858585;">Check the Developer Tools console (Help > Toggle Developer Tools)</p>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
				<script nonce="${nonce}">
					(function() {
						const statusText = document.getElementById('statusText');
						const updateStatus = (msg) => {
							if (statusText) statusText.textContent = msg;
							console.log('[STATUS]', msg);
						};
						
						updateStatus('Script started...');
						console.log('=== WEBVIEW SCRIPT STARTED ===');
						console.log('Document ready state:', document.readyState);
						console.log('Container exists:', !!document.getElementById('handsontableRoot'));
						console.log('Script URI:', '${scriptUriWithCache}');
						
						try {
							updateStatus('Acquiring VS Code API...');
							const vscode = acquireVsCodeApi();
							window.vscode = vscode;
							updateStatus('VS Code API acquired ✓');
							console.log('✓ VS Code API acquired');
							
						updateStatus('Loading bundle.js...');
						console.log('Loading bundle from:', '${scriptUriWithCache}');
						
						// Load bundle script with proper error handling (no inline handlers)
						const script = document.createElement('script');
						script.nonce = '${nonce}';
						script.src = '${scriptUriWithCache}';
						script.onload = function() {
							updateStatus('bundle.js loaded ✓');
							console.log('✓ bundle.js loaded successfully');
						};
						script.onerror = function(event) {
							const errorMsg = '✗ Failed to load bundle.js';
							updateStatus(errorMsg);
							console.error('✗ Failed to load bundle.js:', event);
							console.error('Script URI was:', '${scriptUriWithCache}');
							const container = document.getElementById('handsontableRoot');
							if (container) {
								container.innerHTML = '<div style="padding: 40px; color: #f48771; text-align: center;"><h2>Error</h2><p>Failed to load bundle.js</p><p style="font-size: 11px;">URI: ${scriptUriWithCache}</p><p style="font-size: 11px;">Check Developer Tools console for details</p></div>';
							}
						};
						document.head.appendChild(script);
						updateStatus('Script element added to DOM');
						console.log('Script element added to DOM');
						} catch (error) {
							const errorMsg = '✗ Error: ' + error.message;
							updateStatus(errorMsg);
							console.error('✗ Error in webview script:', error);
							const container = document.getElementById('handsontableRoot');
							if (container) {
								container.innerHTML = '<div style="padding: 40px; color: #f48771; text-align: center;"><h2>Error</h2><p>' + error.message + '</p></div>';
							}
						}
					})();
				</script>
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

