/**
 * Harbor Formulas Webview Main Script
 * Initializes HyperFormula and Handsontable in the VS Code webview
 * 
 * This file is bundled with esbuild to include npm dependencies
 */

import { HyperFormula } from 'hyperformula';
import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';

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

    // Initialize HyperFormula
    const hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
      undoLimit: 500,
      useArrayArithmetic: true,
    });

    // Add Sheet1
    hf.addSheet('Sheet1');
    const sheetId = 0;

    // Create empty data - start with smaller size to avoid array length issues
    const rows = 50;
    const columns = 20;
    console.log(`Creating data array: ${rows} rows x ${columns} columns`);
    
    const data = [];
    try {
      for (let r = 0; r < rows; r++) {
        const row = [];
        for (let c = 0; c < columns; c++) {
          row.push('');
        }
        data.push(row);
      }
      console.log(`Data array created successfully: ${data.length} rows`);
    } catch (error) {
      console.error('Error creating data array:', error);
      throw error;
    }

    // Wait for container to have dimensions before initializing Handsontable
    const ensureContainerSize = () => {
      const rect = container.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) {
        console.log('Container has no size yet, waiting...', rect);
        setTimeout(ensureContainerSize, 100);
        return;
      }
      
      console.log('Container size:', rect.width, 'x', rect.height);
      
      // Ensure we have valid dimensions - use numbers, not strings
      const width = Math.max(Math.floor(rect.width), 800);
      const height = Math.max(Math.floor(rect.height), 600);
      
      console.log('Using dimensions:', width, 'x', height);
      console.log('Data array length:', data.length, 'rows');
      console.log('First row length:', data[0]?.length, 'columns');
      
      // Validate data before passing to Handsontable
      if (!data || data.length === 0) {
        throw new Error('Data array is empty');
      }
      if (!Array.isArray(data[0])) {
        throw new Error('Data array is not 2D');
      }
      
      // Initialize Handsontable with explicit numeric dimensions
      try {
        console.log('Initializing Handsontable...');
        console.log('Container element:', container);
        console.log('Container tagName:', container.tagName);
        console.log('Container dimensions:', { width, height });
        console.log('Data validation:', {
          isArray: Array.isArray(data),
          length: data.length,
          firstRowLength: data[0]?.length,
          sampleRow: data[0]
        });
        
        // Try minimal configuration first
        const hotConfig = {
          data: data,
          width: width,
          height: height,
          rowHeaders: true,
          colHeaders: true,
          licenseKey: 'non-commercial-and-evaluation',
        };
        
        console.log('Handsontable config:', JSON.stringify(hotConfig, null, 2));
        console.log('Creating Handsontable instance...');
        
        const hot = new Handsontable(container, hotConfig);
        console.log('Handsontable instance created:', hot);
        console.log('Handsontable initialized successfully');
        
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
        console.error('Error initializing Handsontable:', error);
        throw error;
      }
    };
    
    // Start checking for container size
    ensureContainerSize();

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

