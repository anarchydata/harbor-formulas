/**
 * Grid Component
 * Handles all grid-related functionality including cell selection, editing, and display updates
 */

// Grid state
export let selectedCells = new Set();
export let isSelecting = false;
export let selectionStart = null;
export let isEditMode = false;
window.selectedCell = null;

// Helper functions for cell references
export function colToLetter(col) {
  let letter = "";
  let num = col;
  do {
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26) - 1;
  } while (num >= 0);
  return letter;
}

export function letterToCol(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function cellRefToAddress(cellRef) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  const col = letterToCol(match[1].toUpperCase());
  const row = parseInt(match[2]) - 1; // HyperFormula uses 0-based indexing
  return [row, col];
}

export function addressToCellRef(row, col) {
  return colToLetter(col) + (row + 1);
}

export function getCellReference(row, col) {
  return colToLetter(col) + (row + 1);
}

// Cell display and value functions
export function updateCellDisplay(cellRef, gridBody) {
  if (!window.hf) return;
  
  const address = cellRefToAddress(cellRef);
  if (!address) return;
  
  const [row, col] = address;
  const sheetId = 0;
  
  try {
    const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
    const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
    if (cell) {
      const display = cell.querySelector(".grid-cell-display");
      if (display) {
        display.textContent = cellValue !== null && cellValue !== undefined ? cellValue.toString() : "";
      }
    }
  } catch (error) {
    console.error(`Error updating cell display for ${cellRef}:`, error);
  }
}

export function getDependentCells(row, col, sheetId, gridBody) {
  if (!window.hf) return [];
  
  const dependentCells = [];
  
  try {
    const changedCellRef = addressToCellRef(row, col);
    const allCells = gridBody.querySelectorAll('td[data-ref]');
    
    allCells.forEach(cellElement => {
      const cellRef = cellElement.getAttribute("data-ref");
      if (!cellRef) return;
      
      const storedFormula = cellElement.getAttribute("data-formula");
      if (storedFormula && storedFormula.startsWith("=")) {
        const baseRef = changedCellRef;
        const colLetter = baseRef.match(/^([A-Z]+)/i)?.[1] || "";
        const rowNum = baseRef.match(/(\d+)$/)?.[1] || "";
        const escapedColLetter = colLetter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const pattern = new RegExp(`\\$?${escapedColLetter}\\$?${rowNum}\\b`, 'i');
        const referencesChangedCell = pattern.test(storedFormula);
        
        if (referencesChangedCell) {
          const depAddress = cellRefToAddress(cellRef);
          if (depAddress) {
            const [depRow, depCol] = depAddress;
            dependentCells.push({ row: depRow, col: depCol, ref: cellRef });
          }
        }
      } else {
        const depAddress = cellRefToAddress(cellRef);
        if (depAddress) {
          const [depRow, depCol] = depAddress;
          try {
            const cellFormula = window.hf.getCellFormula({ col: depCol, row: depRow, sheet: sheetId });
            if (cellFormula && cellFormula.startsWith("=")) {
              const baseRef = changedCellRef;
              const colLetter = baseRef.match(/^([A-Z]+)/i)?.[1] || "";
              const rowNum = baseRef.match(/(\d+)$/)?.[1] || "";
              const pattern = new RegExp(`\\b\\$?${colLetter}\\$?${rowNum}\\b`, 'i');
              
              if (pattern.test(cellFormula)) {
                dependentCells.push({ row: depRow, col: depCol, ref: cellRef });
              }
            }
          } catch (e) {
            // Cell might not exist, skip it
          }
        }
      }
    });
  } catch (error) {
    console.error("Error getting dependent cells:", error);
  }
  
  return dependentCells;
}

export function setCellValue(cellRef, value, gridBody) {
  if (!window.hf) {
    console.warn("HyperFormula not initialized");
    return;
  }
  
  const address = cellRefToAddress(cellRef);
  if (!address) {
    console.warn("Invalid cell reference:", cellRef);
    return;
  }
  
  const [row, col] = address;
  
  try {
    const sheetId = 0;
    let formulaToStore = value;
    
    window.hf.setCellContents({ col, row, sheet: sheetId }, [[formulaToStore]]);
    window.hf.rebuildAndRecalculate();
    
    updateCellDisplay(cellRef, gridBody);
    
    const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
    if (cell) {
      cell.setAttribute("data-formula", formulaToStore);
    }
    
    const dependentCells = getDependentCells(row, col, sheetId, gridBody);
    dependentCells.forEach(depCell => {
      updateCellDisplay(depCell.ref, gridBody);
      const depCellElement = gridBody.querySelector(`td[data-ref="${depCell.ref}"]`);
      if (depCellElement) {
        try {
          const depCellFormula = window.hf.getCellFormula({ col: depCell.col, row: depCell.row, sheet: sheetId });
          if (depCellFormula) {
            depCellElement.setAttribute("data-formula", depCellFormula);
          }
        } catch (e) {
          // Formula might not exist, skip
        }
      }
    });
  } catch (error) {
    console.error("Error setting cell value:", error);
  }
}

export function getCellValue(cellRef) {
  if (!window.hf) {
    return null;
  }
  
  const address = cellRefToAddress(cellRef);
  if (!address) {
    return null;
  }
  
  const [row, col] = address;
  try {
    const sheetId = 0;
    const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
    return cellValue;
  } catch (error) {
    console.error("Error getting cell value:", error);
    return null;
  }
}

// Selection functions
export function clearSelection() {
  selectedCells.forEach(cell => {
    cell.classList.remove("selected");
    cell.classList.remove("edit-mode");
    cell.removeAttribute("data-selection-edge");
  });
  selectedCells.clear();
}

export function updateCellRangePill() {
  const pill = document.getElementById("cellRangePill");
  if (!pill) return;

  if (selectedCells.size === 0) {
    pill.textContent = "";
    return;
  }

  const cellsArray = Array.from(selectedCells);
  const rows = cellsArray.map(cell => parseInt(cell.getAttribute("data-row"))).filter(r => !isNaN(r));
  const cols = cellsArray.map(cell => parseInt(cell.getAttribute("data-col"))).filter(c => !isNaN(c));

  if (rows.length === 0 || cols.length === 0) {
    pill.textContent = "";
    return;
  }

  const minRow = Math.min(...rows);
  const maxRow = Math.max(...rows);
  const minCol = Math.min(...cols);
  const maxCol = Math.max(...cols);

  if (minRow === maxRow && minCol === maxCol) {
    pill.textContent = getCellReference(minRow, minCol);
  } else {
    pill.textContent = `${getCellReference(minRow, minCol)}:${getCellReference(maxRow, maxCol)}`;
  }
}

export function selectRange(startCell, endCell, editor, gridBody, ensureSevenLines) {
  clearSelection();
  
  if (!startCell || !endCell) return;
  
  const startRow = parseInt(startCell.getAttribute("data-row"));
  const startCol = parseInt(startCell.getAttribute("data-col"));
  const endRow = parseInt(endCell.getAttribute("data-row"));
  const endCol = parseInt(endCell.getAttribute("data-col"));
  
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cell = gridBody.querySelector(`tr[data-row="${r}"] td[data-col="${c}"]`);
      if (cell) {
        const isRowHeader = cell === cell.parentElement.querySelector("td:first-child");
        if (!isRowHeader) {
          cell.classList.add("selected");
          selectedCells.add(cell);
          
          const edges = [];
          if (r === minRow) edges.push("top");
          if (r === maxRow) edges.push("bottom");
          if (c === minCol) edges.push("left");
          if (c === maxCol) edges.push("right");
          
          if (edges.length > 0) {
            cell.setAttribute("data-selection-edge", edges.join(" "));
          } else {
            cell.removeAttribute("data-selection-edge");
          }
        }
      }
    }
  }
  
  const firstCell = gridBody.querySelector(`tr[data-row="${minRow}"] td[data-col="${minCol}"]`);
  if (firstCell) {
    const isHeader = firstCell.tagName === "TH" || firstCell === firstCell.parentElement.querySelector("td:first-child");
    if (!isHeader) {
      const storedFormula = firstCell.getAttribute("data-formula") || "";
      let editorValue = "";
      if (storedFormula && storedFormula.startsWith("=")) {
        editorValue = storedFormula.substring(1);
      } else if (storedFormula) {
        editorValue = storedFormula;
      }
      
      const currentEditor = window.monacoEditor || editor;
      if (currentEditor) {
        const fullValue = "\n" + (editorValue || "");
        const line2Content = (editorValue || "").split('\n')[0] || "";
        const targetColumn = Math.max(1, line2Content.length + 1);
        
        window.isProgrammaticCursorChange = true;
        currentEditor.setValue(ensureSevenLines(fullValue));
        currentEditor.setPosition({ lineNumber: 2, column: targetColumn });
        
        Promise.resolve().then(() => {
          window.isProgrammaticCursorChange = false;
        });
        
        currentEditor.updateOptions({ readOnly: false });
      }
    } else {
      const currentEditor = window.monacoEditor || editor;
      if (currentEditor) {
        currentEditor.updateOptions({ readOnly: true });
      }
    }
    window.selectedCell = firstCell;
  }
  updateCellRangePill();
}

export function selectCell(cell, editor, gridBody, ensureSevenLines) {
  if (!cell) return;
  clearSelection();
  selectedCells.add(cell);
  cell.classList.add("selected");
  if (isEditMode) {
    cell.classList.add("edit-mode");
  }
  cell.setAttribute("data-selection-edge", "top bottom left right");
  window.selectedCell = cell;
  
  const cellRef = cell.getAttribute("data-ref");
  const isHeader = cell.tagName === "TH" || cell === cell.parentElement.querySelector("td:first-child");
  if (!isHeader) {
    let storedFormula = cell.getAttribute("data-formula");
    let editorValue = "";
    
    if (storedFormula) {
      if (storedFormula.startsWith("=")) {
        editorValue = storedFormula.substring(1);
      } else {
        editorValue = storedFormula;
      }
    } else {
      if (cellRef && window.hf) {
        const address = cellRefToAddress(cellRef);
        if (address) {
          const [row, col] = address;
          const sheetId = 0;
          let hasFormula = false;
          try {
            const cellFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
            if (cellFormula) {
              hasFormula = true;
              if (cellFormula.startsWith("=")) {
                editorValue = cellFormula.substring(1);
              } else {
                editorValue = cellFormula;
              }
              const formulaToStore = cellFormula.startsWith("=") ? cellFormula : "=" + cellFormula;
              cell.setAttribute("data-formula", formulaToStore);
            }
          } catch (formulaError) {
            hasFormula = false;
          }
          
          if (!hasFormula) {
            try {
              const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
              if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
                editorValue = cellValue.toString();
                cell.setAttribute("data-formula", cellValue.toString());
              } else {
                editorValue = "";
                cell.removeAttribute("data-formula");
              }
            } catch (valueError) {
              editorValue = "";
              cell.removeAttribute("data-formula");
            }
          }
        }
      }
    }
    
    const currentEditor = window.monacoEditor || editor;
    if (currentEditor && typeof currentEditor.setValue === 'function') {
      const fullValue = ensureSevenLines("\n" + (editorValue || ""));
      const line2Content = (editorValue || "").split('\n')[0] || "";
      const targetColumn = Math.max(1, line2Content.length + 1);
      
      window.isProgrammaticCursorChange = true;
      const model = currentEditor.getModel();
      
      if (model) {
        currentEditor.executeEdits('setCellValue', [
          {
            range: model.getFullModelRange(),
            text: fullValue
          }
        ], [
          {
            range: new window.monaco.Range(2, targetColumn, 2, targetColumn),
            selection: new window.monaco.Range(2, targetColumn, 2, targetColumn)
          }
        ]);
      } else {
        currentEditor.setValue(fullValue);
        currentEditor.setPosition({ lineNumber: 2, column: targetColumn });
      }
      
      Promise.resolve().then(() => {
        window.isProgrammaticCursorChange = false;
      });
      
      currentEditor.updateOptions({ readOnly: false });
      
      if (isEditMode && typeof currentEditor.focus === 'function') {
        currentEditor.focus();
      }
    }
  } else {
    const currentEditor = window.monacoEditor || editor;
    if (currentEditor) {
      currentEditor.updateOptions({ readOnly: true });
    }
  }
  updateCellRangePill();
}

export function enterEditMode(cell, editor, gridBody, ensureSevenLines) {
  if (!cell) return;
  
  const isHeader = cell.tagName === "TH" || cell === cell.parentElement.querySelector("td:first-child");
  if (isHeader) return;
  
  isEditMode = true;
  
  clearSelection();
  selectedCells.add(cell);
  cell.classList.add("selected");
  cell.classList.add("edit-mode");
  cell.setAttribute("data-selection-edge", "top bottom left right");
  window.selectedCell = cell;
  updateCellRangePill();
  
  const cellRef = cell.getAttribute("data-ref");
  const storedFormula = cell.getAttribute("data-formula") || "";
  
  let editorValue = "";
  if (storedFormula && storedFormula.startsWith("=")) {
    editorValue = storedFormula.substring(1);
  } else if (storedFormula) {
    editorValue = storedFormula;
  } else {
    if (cellRef && window.hf) {
      const address = cellRefToAddress(cellRef);
      if (address) {
        const [row, col] = address;
        const sheetId = 0;
        try {
          const cellFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
          if (cellFormula && cellFormula.startsWith("=")) {
            editorValue = cellFormula.substring(1);
          } else if (cellFormula) {
            editorValue = cellFormula;
          }
        } catch (e) {
          editorValue = "";
        }
      }
    }
  }
  
  const currentEditor = window.monacoEditor || editor;
  if (currentEditor && typeof currentEditor.setValue === 'function') {
    const fullValue = ensureSevenLines("\n" + (editorValue || ""));
    const line2Content = (editorValue || "").split('\n')[0] || "";
    const targetColumn = Math.max(1, line2Content.length + 1);
    
    window.isProgrammaticCursorChange = true;
    currentEditor.setValue(fullValue);
    currentEditor.updateOptions({ readOnly: false });
    currentEditor.setPosition({ lineNumber: 2, column: targetColumn });
    currentEditor.focus();
    
    Promise.resolve().then(() => {
      window.isProgrammaticCursorChange = false;
    });
    
    setTimeout(() => {
      if (currentEditor && typeof currentEditor.focus === 'function') {
        currentEditor.focus();
        const model = currentEditor.getModel();
        if (model) {
          const line2Content = model.getLineContent(2);
          const col = Math.max(1, line2Content.length + 1);
          currentEditor.setPosition({ lineNumber: 2, column: col });
        }
      }
    }, 50);
  }
}

// Export state setters/getters for external access
export function setIsEditMode(value) {
  isEditMode = value;
}

export function getIsEditMode() {
  return isEditMode;
}

export function setIsSelecting(value) {
  isSelecting = value;
}

export function getIsSelecting() {
  return isSelecting;
}

export function setSelectionStart(cell) {
  selectionStart = cell;
}

export function getSelectionStart() {
  return selectionStart;
}

