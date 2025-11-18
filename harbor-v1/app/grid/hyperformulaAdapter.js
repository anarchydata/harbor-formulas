/**
 * HyperFormula ↔ Handsontable Data Adapter
 * 
 * This module bridges HyperFormula (calculation engine) with Handsontable (UI grid).
 * It handles:
 * - Syncing display values from HyperFormula to Handsontable
 * - Updating cells when HyperFormula recalculates
 * - Maintaining cell formatting (numeric vs text alignment)
 */

import { addressToCellRef, cellRefToAddress } from '../../utils/helpers.js';

/**
 * Creates a data adapter that syncs HyperFormula with Handsontable
 * @param {Object} params
 * @param {Object} params.hfInstance - HyperFormula instance (window.hf)
 * @param {Object} params.handsontableController - Handsontable controller from createHandsontableGrid
 * @param {number} params.sheetId - Current sheet ID (default: 0)
 * @returns {Object} Adapter API
 */
export function createHyperFormulaAdapter({
  hfInstance,
  handsontableController,
  sheetId = 0
}) {
  if (!hfInstance) {
    throw new Error('HyperFormula instance is required');
  }
  if (!handsontableController || !handsontableController.hot) {
    throw new Error('Handsontable controller is required');
  }

  const hot = handsontableController.hot;
  let isUpdating = false; // Flag to prevent circular updates

  /**
   * Updates a single cell's display value in Handsontable from HyperFormula
   * @param {string} cellRef - Cell reference (e.g., "A1")
   */
  function updateCellDisplay(cellRef) {
    if (!hfInstance || isUpdating) return;

    const address = cellRefToAddress(cellRef);
    if (!address) return;

    const [row, col] = address;

    try {
      // Check if cell has a formula
      let hasFormula = false;
      let storedFormula = null;
      try {
        storedFormula = hfInstance.getCellFormula({ col, row, sheet: sheetId });
        if (storedFormula) {
          hasFormula = true;
        }
      } catch (e) {
        // Cell doesn't have a formula
        hasFormula = false;
      }

      // Get the calculated value from HyperFormula
      const cellValue = hfInstance.getCellValue({ col, row, sheet: sheetId });

      // Format the display value
      const displayValue = cellValue !== null && cellValue !== undefined 
        ? cellValue.toString() 
        : '';

      // Update Handsontable cell
      handsontableController.setCellDisplayValue(row, col, displayValue);

      // Get the cell DOM element from Handsontable
      // Use the body element to find the cell by data attributes
      const bodyElement = handsontableController.getBodyElement();
      if (bodyElement) {
        const cellElement = bodyElement.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
        if (cellElement) {
          // Store formula in cell metadata for quick retrieval
          if (hasFormula && storedFormula) {
            cellElement.setAttribute('data-formula', storedFormula);
          } else if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            cellElement.setAttribute('data-formula', cellValue.toString());
          } else {
            cellElement.removeAttribute('data-formula');
          }

          // Apply numeric/text alignment
          const displaySpan = cellElement.querySelector('.grid-cell-display');
          if (displaySpan) {
            const isNumber = determineIfNumber(cellValue, displayValue);
            if (isNumber) {
              displaySpan.style.textAlign = 'right';
              cellElement.classList.add('cell-numeric');
              cellElement.classList.remove('cell-text');
            } else {
              displaySpan.style.textAlign = 'left';
              cellElement.classList.add('cell-text');
              cellElement.classList.remove('cell-numeric');
            }
          }
        }
      }
    } catch (error) {
      console.error(`Error updating cell display for ${cellRef}:`, error);
    }
  }

  /**
   * Determines if a value should be treated as a number (right-aligned)
   */
  function determineIfNumber(cellValue, displayValue) {
    if (cellValue === null || cellValue === undefined || displayValue === '') {
      return false;
    }

    // Check if it's a JavaScript number type
    if (typeof cellValue === 'number') {
      return true;
    }

    // Check if string represents a number
    const numValue = parseFloat(displayValue);
    if (!isNaN(numValue) && isFinite(numValue) && displayValue.trim() !== '') {
      const trimmed = displayValue.trim();
      // Match numbers: integers, decimals, scientific notation
      if (/^-?\d+\.?\d*([eE][+-]?\d+)?$/.test(trimmed)) {
        return true;
      }
    }

    return false;
  }

  /**
   * Updates multiple cells at once (batch update)
   * @param {string[]} cellRefs - Array of cell references
   */
  function updateCells(cellRefs) {
    if (!hfInstance || isUpdating) return;

    cellRefs.forEach(cellRef => {
      updateCellDisplay(cellRef);
    });
  }

  /**
   * Gets all cells that depend on a given cell
   * @param {number} row - Row index
   * @param {number} col - Column index
   * @returns {Array<{row: number, col: number, ref: string}>} Array of dependent cell info
   */
  function getDependentCells(row, col) {
    if (!hfInstance) return [];

    const dependentCells = [];
    const changedCellRef = addressToCellRef(row, col);

    try {
      // Get all cells in the visible grid
      const data = hot.getData();
      const maxRows = data.length;
      const maxCols = data[0]?.length || 0;

      // Check each cell for formulas that reference the changed cell
      for (let r = 0; r < maxRows; r++) {
        for (let c = 0; c < maxCols; c++) {
          try {
            const formula = hfInstance.getCellFormula({ col: c, row: r, sheet: sheetId });
            if (formula && formula.includes(changedCellRef)) {
              dependentCells.push({
                row: r,
                col: c,
                ref: addressToCellRef(r, c)
              });
            }
          } catch (e) {
            // Cell doesn't have a formula, skip
          }
        }
      }
    } catch (error) {
      console.error('Error getting dependent cells:', error);
    }

    return dependentCells;
  }

  /**
   * Syncs all visible cells from HyperFormula to Handsontable
   * Useful for initial load or full refresh
   */
  function syncAllCells() {
    if (!hfInstance || isUpdating) return;

    try {
      const data = hot.getData();
      const maxRows = data.length;
      const maxCols = data[0]?.length || 0;

      for (let row = 0; row < maxRows; row++) {
        for (let col = 0; col < maxCols; col++) {
          const cellRef = addressToCellRef(row, col);
          updateCellDisplay(cellRef);
        }
      }
    } catch (error) {
      console.error('Error syncing all cells:', error);
    }
  }

  /**
   * Sets the updating flag to prevent circular updates
   */
  function setUpdating(value) {
    isUpdating = value;
  }

  return {
    updateCellDisplay,
    updateCells,
    getDependentCells,
    syncAllCells,
    setUpdating
  };
}

