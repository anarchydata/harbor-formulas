import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';
import 'handsontable/styles/ht-theme-main.css';

import { addressToCellRef } from '../utils/helpers.js';

const DEFAULT_ROWS = 150;
const DEFAULT_COLUMNS = 50;

// Use built-in dark theme

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

  // Add dark theme class to container (Handsontable applies theme via class)
  container.classList.add('ht-theme-main-dark');

  console.log('Creating Handsontable grid with container:', container);
  
  // Function to get actual container dimensions
  const getContainerDimensions = () => {
    const rect = container.getBoundingClientRect();
    const parent = container.parentElement;
    const parentRect = parent ? parent.getBoundingClientRect() : null;
    
    return {
      width: rect.width || container.offsetWidth,
      height: rect.height || container.offsetHeight,
      parentWidth: parentRect?.width || parent?.offsetWidth || 0,
      parentHeight: parentRect?.height || parent?.offsetHeight || 0,
      computed: {
        width: window.getComputedStyle(container).width,
        height: window.getComputedStyle(container).height
      }
    };
  };
  
  const dims = getContainerDimensions();
  console.log('Container dimensions:', dims);
  
  // Calculate height - use actual computed dimensions or fallback
  let containerHeight = dims.height || dims.parentHeight || window.innerHeight * 0.6 || 600;
  let containerWidth = dims.width || dims.parentWidth || '100%';
  
  // If container has no height, wait for layout and use parent or viewport
  if (containerHeight === 0 || containerHeight < 100) {
    console.warn('Container has zero/small height, using calculated height');
    // Try to get from parent chain
    let current = container.parentElement;
    while (current && containerHeight < 100) {
      const h = current.offsetHeight || current.getBoundingClientRect().height;
      if (h > 0) {
        containerHeight = h;
        break;
      }
      current = current.parentElement;
    }
    // Fallback to viewport
    if (containerHeight < 100) {
      containerHeight = window.innerHeight * 0.6;
    }
    console.log('Using calculated height:', containerHeight);
  }

  const data = Handsontable.helper.createEmptySpreadsheetData(rows, columns);
  
  const hot = new Handsontable(container, {
    data,
    rowHeaders: true,
    colHeaders: (index) => columnLabelFromIndex(index),
    readOnly: true,
    height: containerHeight,
    width: containerWidth,
    licenseKey: 'non-commercial-and-evaluation',
    themeName: 'ht-theme-main-dark', // Theme name with obligatory 'ht-theme-*' prefix
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

  // Force render after container is laid out
  // Use multiple attempts to ensure it renders
  const ensureRender = () => {
    if (hot && typeof hot.render === 'function') {
      const dims = getContainerDimensions();
      console.log('Rendering Handsontable with dimensions:', dims);
      
      // Update height if container now has dimensions
      if (dims.height > 0 && dims.height !== containerHeight) {
        hot.updateSettings({ height: dims.height });
      }
      
      hot.render();
      console.log('Handsontable rendered, root element:', hot.rootElement);
      
      // Verify it actually rendered
      const body = hot.rootElement?.querySelector('.ht_master .htCore tbody');
      if (body) {
        console.log('Handsontable body element found:', body);
      } else {
        console.warn('Handsontable body element not found after render');
      }
    }
  };
  
  // Try immediate render
  requestAnimationFrame(ensureRender);
  
  // Also try after a short delay in case layout isn't ready
  setTimeout(ensureRender, 100);

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

