/**
 * Harbor Formulas Webview Main Script
 * Initializes HyperFormula and Handsontable in the VS Code webview
 */

(async function initializeHarborFormulas() {
  try {
    console.log('Initializing Harbor Formulas in webview...');

    const container = document.getElementById('handsontableRoot');
    if (!container) {
      throw new Error('Grid container not found');
    }

    // Show loading state
    container.innerHTML = `
      <div style="display: flex; align-items: center; justify-content: center; height: 100%; flex-direction: column; gap: 20px;">
        <h2 style="color: var(--text);">Harbor Formulas</h2>
        <p style="color: var(--text-muted);">Loading spreadsheet engine...</p>
      </div>
    `;

    // Load HyperFormula and Handsontable from CDN
    // In production, these should be bundled
    await Promise.all([
      loadScript('https://cdn.jsdelivr.net/npm/hyperformula@3.1.0/dist/hyperformula.full.min.js'),
      loadScript('https://cdn.jsdelivr.net/npm/handsontable@16.1.1/dist/handsontable.full.min.js'),
      loadStylesheet('https://cdn.jsdelivr.net/npm/handsontable@16.1.1/dist/handsontable.full.min.css')
    ]);

    // Initialize HyperFormula
    const { HyperFormula } = window.HyperFormula;
    const hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
      undoLimit: 500,
      useArrayArithmetic: true,
    });

    // Add Sheet1
    hf.addSheet('Sheet1');
    const sheetId = 0;

    // Create empty data
    const rows = 150;
    const columns = 50;
    const data = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < columns; c++) {
        row.push('');
      }
      data.push(row);
    }

    // Initialize Handsontable
    const hot = new Handsontable(container, {
      data: data,
      rowHeaders: true,
      colHeaders: (index) => {
        let label = '';
        let num = index;
        do {
          label = String.fromCharCode(65 + (num % 26)) + label;
          num = Math.floor(num / 26) - 1;
        } while (num >= 0);
        return label;
      },
      readOnly: false,
      licenseKey: 'non-commercial-and-evaluation',
      themeName: 'ht-theme-main-dark',
      selectionMode: 'multiple',
      outsideClickDeselects: false,
      autoRowSize: false,
      autoColumnSize: false,
    });

    // Store references globally for debugging
    window.hf = hf;
    window.hot = hot;
    window.hfSheetId = sheetId;

    console.log('Harbor Formulas initialized successfully');
    console.log('HyperFormula:', hf);
    console.log('Handsontable:', hot);

    // Send success message to extension host
    if (window.vscode) {
      window.vscode.postMessage({
        command: 'initialized',
        text: 'Harbor Formulas initialized'
      });
    }
  } catch (error) {
    console.error('Failed to initialize Harbor Formulas:', error);
    
    const container = document.getElementById('handsontableRoot');
    if (container) {
      container.innerHTML = `
        <div style="display: flex; align-items: center; justify-content: center; height: 100%; flex-direction: column; gap: 20px; padding: 40px;">
          <h2 style="color: #f48771;">Error</h2>
          <p style="color: var(--text-muted);">${error.message}</p>
          <p style="color: var(--text-muted); font-size: 11px;">Check the console for details</p>
        </div>
      `;
    }
    
    // Send error to extension host
    if (window.vscode) {
      window.vscode.postMessage({
        command: 'error',
        text: `Failed to initialize: ${error.message}`
      });
    }
  }
})();

// Helper functions
function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = () => reject(new Error(`Failed to load script: ${url}`));
    document.head.appendChild(script);
  });
}

function loadStylesheet(url) {
  return new Promise((resolve, reject) => {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = url;
    link.onload = resolve;
    link.onerror = () => reject(new Error(`Failed to load stylesheet: ${url}`));
    document.head.appendChild(link);
  });
}

