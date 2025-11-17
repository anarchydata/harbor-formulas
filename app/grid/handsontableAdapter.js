import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';

import { addressToCellRef } from '../../utils/helpers.js';

const DEFAULT_ROWS = 150;
const DEFAULT_COLUMNS = 50;

function columnLabelFromIndex(index = 0) {
  let label = '';
  let num = index;
  do {
    label = String.fromCharCode(65 + (num % 26)) + label;
    num = Math.floor(num / 26) - 1;
  } while (num >= 0);
  return label;
}

export function createHandsontableGrid({
  container,
  rows = DEFAULT_ROWS,
  columns = DEFAULT_COLUMNS,
  onCellRender,
  onSelection,
  onDoubleClick,
  onBeforeKeyDown
}) {
  if (!container) {
    throw new Error('Handsontable container element is required');
  }

  const data = Handsontable.helper.createEmptySpreadsheetData(rows, columns);
  
  // Calculate height based on container's parent (grid-wrapper-inner) or use a default
  // Wait for next frame to ensure container is laid out
  let containerHeight = 600;
  const parent = container.parentElement;
  if (parent) {
    // Try to get height from parent, fallback to viewport calculation
    containerHeight = parent.offsetHeight || window.innerHeight * 0.6 || 600;
  }
  
  const hot = new Handsontable(container, {
    data,
    rowHeaders: true,
    colHeaders: (index) => columnLabelFromIndex(index),
    readOnly: true,
    height: containerHeight,
    width: '100%',
    licenseKey: 'non-commercial-and-evaluation',
    theme: 'modern', // Use modern theme instead of deprecated classic
    selectionMode: 'multiple',
    outsideClickDeselects: false,
    disableVisualSelection: false, // Enable Handsontable's native visual selection
    autoRowSize: false,
    autoColumnSize: false,
    manualColumnResize: true,
    manualRowResize: true,
    className: 'hf-grid',
    afterGetColHeader(col, TH) {
      if (TH) {
        TH.dataset.colIndex = col;
      }
    },
    afterGetRowHeader(row, TH) {
      if (TH) {
        TH.dataset.rowIndex = row;
      }
    },
    cells() {
      const renderCell = (instance, TD, row, column, prop, value) => {
        TD.innerHTML = '';
        TD.dataset.row = row;
        TD.dataset.col = column;
        TD.dataset.ref = addressToCellRef(row, column);

        const span = document.createElement('span');
        span.className = 'grid-cell-display';
        span.textContent = value == null ? '' : value;
        TD.appendChild(span);

        if (typeof onCellRender === 'function') {
          onCellRender(TD, row, column);
        }

        return TD;
      };

      return {
        renderer: renderCell
      };
    },
    afterSelection(r, c, r2, c2) {
      // Handsontable selection event
      // r, c = start row/col, r2, c2 = end row/col
      if (typeof onSelection === 'function') {
        onSelection(r, c, r2, c2);
      }
    },
    afterDblClick(event, coords, TD) {
      // Handsontable double-click event
      if (typeof onDoubleClick === 'function') {
        onDoubleClick(event, coords, TD);
      }
    },
    beforeKeyDown(event) {
      // Handsontable keydown event (before default handling)
      if (typeof onBeforeKeyDown === 'function') {
        const result = onBeforeKeyDown(event);
        // If handler returns false, prevent default Handsontable behavior
        if (result === false) {
          event.stopImmediatePropagation();
          return false;
        }
      }
    }
  });

  function getBodyElement() {
    return hot.rootElement.querySelector('.ht_master .htCore tbody');
  }

  function getHeaderElement() {
    return hot.rootElement.querySelector('.ht_clone_top .htCore thead');
  }

  function getRowHeaderElement() {
    return hot.rootElement.querySelector('.ht_clone_left .htCore tbody');
  }

  function setCellDisplayValue(row, column, displayValue) {
    // Use silent update to avoid triggering change hooks
    hot.setDataAtCell(row, column, displayValue, 'hyperformula-update');
  }

  // Update height when container is resized
  function updateHeight() {
    const parent = container.parentElement;
    if (parent) {
      const newHeight = parent.offsetHeight || window.innerHeight * 0.6 || 600;
      hot.updateSettings({ height: newHeight });
    }
  }

  // Listen for resize events
  if (container.parentElement) {
    const resizeObserver = new ResizeObserver(() => {
      updateHeight();
    });
    resizeObserver.observe(container.parentElement);
  }

  // Also update on window resize
  window.addEventListener('resize', updateHeight);

  return {
    hot,
    getBodyElement,
    getHeaderElement,
    getRowHeaderElement,
    setCellDisplayValue,
    updateHeight
  };
}

