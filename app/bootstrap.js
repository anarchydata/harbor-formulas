import {
  createFunctionSnippet,
  ensureSevenLines,
  stripComments,
  escapeHtml,
  cellRefToAddress,
  addressToCellRef,
  isInsideQuotes,
  detectCellReferences,
  extendWithCustomFunctions,
  BASE_CHIP_COLORS
} from '../utils/helpers.js';
import { createDiagnosticsProvider as createExternalDiagnosticsProvider } from '../utils/diagnostics.js';
import {
  generateValueFill,
  AutofillDirection,
  adjustFormulaReferences
} from '../utils/autofill.js';
import { initCustomScrollbars } from './ui/customScrollbars.js';
import { initResizablePanes } from './ui/resizablePanes.js';
import { initClappyChat } from './ui/clappyChat.js';

export async function initializeApp() {
        console.log("DOMContentLoaded fired");
        try {
          // Initialize HyperFormula
          console.log("Loading HyperFormula...");
          const { HyperFormula } = await import('hyperformula');
          console.log("HyperFormula loaded successfully");

          // Initialize HyperFormula instance (following official documentation)
          window.hf = HyperFormula.buildEmpty({
            licenseKey: 'gpl-v3', // Using GPL license for open source projects
            undoLimit: 500,
            useArrayArithmetic: true // Enable array arithmetic globally so range math works like Excel
          });

          // Initialize named ranges storage (Set for fast lookup)
          window.namedRanges = new Set();

          // Store raw formulas (including comments) keyed by sheet + cell reference
          const existingRawStore = window.rawFormulaStore;
          const rawFormulaStore = existingRawStore instanceof Map ? existingRawStore : new Map();
          window.rawFormulaStore = rawFormulaStore;

          // Add Sheet1 - addSheet() returns the sheet name, not the ID
          const sheetName = window.hf.addSheet('Sheet1');
          console.log("addSheet('Sheet1') returned:", sheetName);

          // The first sheet added is always at index 0
          window.hfSheetId = 0;

          // Verify the sheet exists
          const sheetNames = window.hf.getSheetNames();
          if (sheetNames.length === 0 || sheetNames[0] !== 'Sheet1') {
            throw new Error("Sheet1 was not created successfully. Available sheets: " + sheetNames.join(', '));
          }

          console.log("HyperFormula initialized", window.hf);
          console.log("Sheet 'Sheet1' at index:", window.hfSheetId);
          console.log("Available sheets:", sheetNames);

          // Folder collapse/expand functionality
          const projectFolder = document.getElementById("projectFolder");
          const projectFolderContent = document.getElementById("projectFolderContent");
          const folderChevron = projectFolder.querySelector(".folder-chevron");

          if (projectFolder && projectFolderContent) {
            projectFolder.addEventListener("click", () => {
              const isCollapsed = projectFolderContent.classList.contains("collapsed");
              if (isCollapsed) {
                projectFolderContent.classList.remove("collapsed");
                folderChevron.classList.remove("collapsed");
              } else {
                projectFolderContent.classList.add("collapsed");
                folderChevron.classList.add("collapsed");
              }
            });
          }

          const monacoEditorContainer = document.getElementById("monacoEditor");
          const gridBody = document.getElementById("gridBody");
          const gridHeaderElement = document.getElementById("gridHeader");
          const gridWrapperInner = document.getElementById("gridWrapperInner");
          const fillHandleElement = document.getElementById("fillHandle");
        const clappyTranscript = document.getElementById("clappyTranscript");
        const clappyForm = document.getElementById("clappyForm");
        const clappyInput = document.getElementById("clappyInput");
        const chatCollapseToggle = document.getElementById("chatCollapseToggle");
        const mainChatContainer = document.getElementById("chatContainer");
        const clappyConsole = document.getElementById("clappyConsole");
        const gridContainer = document.querySelector(".grid-container");
        const middleSection = document.querySelector(".middle-section");
        const sidebarElement = document.querySelector(".sidebar");
        const sidebarToggleButton = document.getElementById("sidebarCollapseToggle");
        const sheetTabsBarElement = document.getElementById("sheetTabsBar");
          const editorStatusBarElement = document.getElementById("editorStatusBar");
          const editorStatusLineElement = document.getElementById("editorStatusLine");
          const editorStatusColumnElement = document.getElementById("editorStatusColumn");
          const dependenciesContentElement = document.getElementById("dependenciesContent");
          const dependenciesCountElement = document.getElementById("dependenciesCount");
          const updateDependenciesCountBadge = (countValue) => {
            if (!dependenciesCountElement) {
              return;
            }
            const safeCount = Number.isFinite(countValue) && countValue > 0 ? Math.floor(countValue) : 0;
            dependenciesCountElement.textContent = String(safeCount);
            dependenciesCountElement.setAttribute(
              "aria-label",
              `${safeCount} ${safeCount === 1 ? "dependency" : "dependencies"}`
            );
            dependenciesCountElement.dataset.hasValue = safeCount > 0 ? "true" : "false";
          };

          updateDependenciesCountBadge(0);

          window.monacoEditor = null;
          let editor = null;

        function updateSidebarToggleState(isCollapsed) {
          if (!sidebarToggleButton) {
            return;
          }
          sidebarToggleButton.textContent = isCollapsed ? ">>" : "<<";
          sidebarToggleButton.setAttribute("title", isCollapsed ? "Expand File Explorer" : "Collapse File Explorer");
          sidebarToggleButton.setAttribute("aria-expanded", (!isCollapsed).toString());
        }

        if (sidebarElement && sidebarToggleButton) {
          sidebarToggleButton.addEventListener("click", () => {
            const isCollapsed = sidebarElement.classList.toggle("collapsed");
            updateSidebarToggleState(isCollapsed);
          });
          updateSidebarToggleState(sidebarElement.classList.contains("collapsed"));
        }

        function getActiveSheetId() {
          return typeof window.hfSheetId === "number" ? window.hfSheetId : 0;
        }

        function syncActiveSheetTabFromHyperFormula() {
          if (!sheetTabsBarElement || !window.hf) {
            return;
          }

          const activeTab = sheetTabsBarElement.querySelector(".tab.active:not(.tab-add)");
          if (!activeTab) {
            return;
          }

          const sheetId = getActiveSheetId();
          let sheetNameFromHF = activeTab.dataset.sheetName || activeTab.textContent?.trim() || `Sheet${sheetId + 1}`;

          try {
            const hfSheetName = window.hf.getSheetName(sheetId);
            if (typeof hfSheetName === "string" && hfSheetName.trim()) {
              sheetNameFromHF = hfSheetName.trim();
            }
          } catch (error) {
            console.warn("Could not read sheet name from HyperFormula", error);
          }

          activeTab.dataset.sheetId = String(sheetId);
          activeTab.dataset.sheetName = sheetNameFromHF;
          activeTab.textContent = sheetNameFromHF;
        }

        function startSheetRename(tabButton) {
          if (!tabButton || tabButton.dataset.renaming === "true") {
            return;
          }

          const parentElement = tabButton.parentElement;
          if (!parentElement) {
            return;
          }

          const sheetId = Number.parseInt(tabButton.dataset.sheetId ?? `${getActiveSheetId()}`, 10) || getActiveSheetId();
          const existingName = (tabButton.dataset.sheetName || tabButton.textContent || "").trim() || `Sheet${sheetId + 1}`;
          const renameInput = document.createElement("input");
          renameInput.type = "text";
          renameInput.value = existingName;
          renameInput.className = "tab-rename-input";
          renameInput.setAttribute("aria-label", "Rename sheet");
          renameInput.style.width = `${Math.max(tabButton.offsetWidth, 60)}px`;
          renameInput.style.height = `${tabButton.offsetHeight}px`;

          tabButton.dataset.renaming = "true";
          parentElement.replaceChild(renameInput, tabButton);
          renameInput.focus();
          renameInput.select();

          const finalizeRename = (shouldCommit) => {
            if (!renameInput.isConnected) {
              return;
            }

            renameInput.removeEventListener("keydown", handleKeydown);
            renameInput.removeEventListener("blur", handleBlur);

            parentElement.replaceChild(tabButton, renameInput);
            tabButton.dataset.renaming = "false";

            if (!shouldCommit) {
              tabButton.textContent = existingName;
              tabButton.dataset.sheetName = existingName;
              return;
            }

            const requestedName = renameInput.value.trim();
            const finalName = requestedName || existingName;

            tabButton.textContent = finalName;
            tabButton.dataset.sheetName = finalName;
            tabButton.dataset.sheetId = String(sheetId);

            if (finalName === existingName) {
              return;
            }

            try {
              window.hf.renameSheet(sheetId, finalName);
              console.log(`Renamed sheet ${sheetId} to ${finalName}`);
            } catch (error) {
              console.error("Failed to rename sheet:", error);
              tabButton.textContent = existingName;
              tabButton.dataset.sheetName = existingName;
            }
          };

          const handleKeydown = (event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              finalizeRename(true);
            } else if (event.key === "Escape") {
              event.preventDefault();
              finalizeRename(false);
            }
          };

          const handleBlur = () => finalizeRename(true);

          renameInput.addEventListener("keydown", handleKeydown);
          renameInput.addEventListener("blur", handleBlur);
        }

        function attachSheetRenameHandlers() {
          if (!sheetTabsBarElement) {
            return;
          }

          sheetTabsBarElement.addEventListener("dblclick", (event) => {
            const potentialTab = event.target instanceof HTMLElement ? event.target.closest(".tab") : null;
            if (!potentialTab || potentialTab.classList.contains("tab-add")) {
              return;
            }
            event.preventDefault();
            startSheetRename(potentialTab);
          });
        }

        attachSheetRenameHandlers();
        syncActiveSheetTabFromHyperFormula();
          const SAMPLE_FAKE_DATA = [
            ["Name", "Item", "Qty", "Price", "Date", "Code"],
            ["Alice", "Item1", 10, 5.99, "2025-01-01", "INV-001"],
            ["Bob", "Item2", 12, 7.49, "2025-01-08", "INV-002"],
            ["Carol", "Item3", 14, 9.29, "2025-01-15", "INV-003"],
            ["Dave", "Item4", 16, 12.5, "2025-01-22", "INV-004"],
            ["Eve", "Item5", 18, 14.99, "2025-01-29", "INV-005"]
          ];

          if (!gridBody) {
            console.error("gridBody not found - cannot build grid!");
            return;
          }
          if (!gridHeaderElement) {
            console.error("gridHeader not found - cannot build grid!");
            return;
          }
          let selectionDragPreview = null;
          if (gridWrapperInner) {
            selectionDragPreview = document.createElement("div");
            selectionDragPreview.id = "selectionDragPreview";
            selectionDragPreview.className = "selection-drag-preview";
            gridWrapperInner.appendChild(selectionDragPreview);
          }

          const chipColorPalette = (BASE_CHIP_COLORS || [])
            .filter((color) => color && typeof color.hex === 'string' && color.hex.trim() && color.hex.trim().toUpperCase() !== '#000000')
            .map((color, index) => ({
              ...color,
              className: `chip-color-${index}`
            }));

          const cellChipColorAssignments = new Map();
          const chipHighlightedCells = new Map();
          let nextChipColorIndex = 0;
          let isFormulaSelecting = false;
          let selectedCells = new Set();
          window.selectedCells = selectedCells;
          const highlightedRowHeaders = new Set();
          const highlightedColumnHeaders = new Set();
          let isSelecting = false;
          let selectionStart = null;
          let selectionAnchorCell = null;
          let selectionDragState = null;
          let isEditMode = false;
          let suppressNextDocumentEnter = false;
          const enterModePointerState = {
            active: false,
            pointerRow: null,
            pointerCol: null,
            anchorCellRef: null,
            activeChipRange: null,
            rangeAnchorRef: null
          };
          const editModeRangeHighlight = {
            active: false,
            bounds: null
          };
          let editModePointerArmed = false;
          const POINTER_NAV_SEPARATOR_REGEX = /[ ,()+\-*/^]/;
          const ENTER_MODE_POINTER_BREAK_REGEX = /[\s,+\-*/^()]/;

          function resetEnterModePointerState() {
            enterModePointerState.active = false;
            enterModePointerState.pointerRow = null;
            enterModePointerState.pointerCol = null;
            enterModePointerState.anchorCellRef = null;
            enterModePointerState.activeChipRange = null;
            enterModePointerState.rangeAnchorRef = null;
            clearEditModeRangeHighlight();
          }

          function initializeEnterModePointerState(cell) {
            if (!cell) {
              resetEnterModePointerState();
              return;
            }
            const row = parseInt(cell.getAttribute("data-row"), 10);
            const col = parseInt(cell.getAttribute("data-col"), 10);
            if (Number.isNaN(row) || Number.isNaN(col)) {
              resetEnterModePointerState();
              return;
            }
            enterModePointerState.active = true;
            enterModePointerState.pointerRow = row;
            enterModePointerState.pointerCol = col;
            enterModePointerState.anchorCellRef = cell.getAttribute("data-ref") || null;
            enterModePointerState.activeChipRange = null;
            enterModePointerState.rangeAnchorRef = null;
            clearEditModeRangeHighlight();
          }

          function alignPointerWithAnchorCell() {
            const anchorCell = window.selectedCell || null;
            if (anchorCell) {
              initializeEnterModePointerState(anchorCell);
            } else {
              resetEnterModePointerState();
            }
          }

          function resetPointerAfterSeparatorIfNeeded(text = "") {
            if (!text || !enterModePointerState.active) {
              return;
            }
            const shouldReset = Array.from(text).some(char => ENTER_MODE_POINTER_BREAK_REGEX.test(char));
            if (!shouldReset) {
              return;
            }
            enterModePointerState.activeChipRange = null;
            alignPointerWithAnchorCell();
          }

          function updateEnterModePointerChipRange(range) {
            if (!range) {
              enterModePointerState.activeChipRange = null;
              return;
            }
            enterModePointerState.activeChipRange = {
              startLineNumber: range.startLineNumber,
              startColumn: range.startColumn,
              endLineNumber: range.endLineNumber,
              endColumn: range.endColumn
            };
          }

          function isCursorWithinPointerChip(position) {
            const stored = enterModePointerState.activeChipRange;
            if (!stored || !position) {
              return false;
            }
            if (position.lineNumber < stored.startLineNumber || position.lineNumber > stored.endLineNumber) {
              return false;
            }
            if (position.lineNumber === stored.startLineNumber && position.column < stored.startColumn) {
              return false;
            }
            if (position.lineNumber === stored.endLineNumber && position.column > stored.endColumn) {
              return false;
            }
            return true;
          }

          function getPointerChipRangeForPosition(position) {
            if (!isCursorWithinPointerChip(position)) {
              return null;
            }
            const stored = enterModePointerState.activeChipRange;
            if (!stored || !window.monaco || !window.monaco.Range) {
              return null;
            }
            return new window.monaco.Range(
              stored.startLineNumber,
              stored.startColumn,
              stored.endLineNumber,
              stored.endColumn
            );
          }

          function normalizeSingleCellReference(referenceText = '') {
            if (!referenceText || typeof referenceText !== 'string') {
              return null;
            }
            const trimmed = referenceText.trim();
            if (!trimmed || trimmed.includes(':')) {
              return null;
            }
            const bangIndex = trimmed.lastIndexOf('!');
            const withoutSheet = bangIndex !== -1 ? trimmed.slice(bangIndex + 1) : trimmed;
            if (!withoutSheet) {
              return null;
            }
            return withoutSheet.replace(/\$/g, '');
          }

          function updateEnterModePointerCoordinatesFromReference(reference) {
            if (!enterModePointerState.active) {
              return;
            }
            const normalized = normalizeSingleCellReference(reference);
            if (!normalized) {
              return;
            }
            const address = cellRefToAddress(normalized);
            if (!address) {
              return;
            }
            const [row, col] = address;
            enterModePointerState.pointerRow = row;
            enterModePointerState.pointerCol = col;
          }

          function moveEnterModePointer(deltaRow = 0, deltaCol = 0) {
            if (!enterModePointerState.active) {
              initializeEnterModePointerState(window.selectedCell || null);
            }
            if (!enterModePointerState.active) {
              return null;
            }
            const maxRow = typeof window.totalGridRows === 'number'
              ? Math.max(0, window.totalGridRows - 1)
              : (typeof enterModePointerState.pointerRow === 'number' ? enterModePointerState.pointerRow : 0);
            const maxCol = typeof window.totalGridColumns === 'number'
              ? Math.max(0, window.totalGridColumns - 1)
              : (typeof enterModePointerState.pointerCol === 'number' ? enterModePointerState.pointerCol : 0);
            const currentRow = typeof enterModePointerState.pointerRow === 'number' ? enterModePointerState.pointerRow : 0;
            const currentCol = typeof enterModePointerState.pointerCol === 'number' ? enterModePointerState.pointerCol : 0;
            const targetRow = Math.max(0, Math.min(maxRow, currentRow + deltaRow));
            const targetCol = Math.max(0, Math.min(maxCol, currentCol + deltaCol));
            enterModePointerState.pointerRow = targetRow;
            enterModePointerState.pointerCol = targetCol;
            const reference = addressToCellRef(targetRow, targetCol);
            if (!reference) {
              return null;
            }
            return {
              reference,
              row: targetRow,
              col: targetCol
            };
          }

          function getChipReferenceAtCursor() {
              return null;
          }

          function isCursorOnCellChip() {
            return false;
          }

          function armEditModePointerNavigation() {}

          function disarmEditModePointerNavigation() {
            clearEditModeRangeHighlight();
          }

          function isEditModePointerNavigationAllowed() {
              return true;
          }

          function getCurrentPointerReference() {
              const anchorCell = window.selectedCell || null;
              return anchorCell ? anchorCell.getAttribute("data-ref") : null;
          }

          function buildRangeReference(startRef, endRef) {
            if (!startRef && !endRef) {
              return "";
            }
            if (!startRef) {
              return endRef || "";
            }
            if (!endRef) {
              return startRef || "";
            }
            if (startRef === endRef) {
              return startRef;
            }
            const startAddress = cellRefToAddress(startRef);
            const endAddress = cellRefToAddress(endRef);
            if (!startAddress || !endAddress) {
              return endRef;
            }
            const minRow = Math.min(startAddress[0], endAddress[0]);
            const maxRow = Math.max(startAddress[0], endAddress[0]);
            const minCol = Math.min(startAddress[1], endAddress[1]);
            const maxCol = Math.max(startAddress[1], endAddress[1]);
            const normalizedStart = addressToCellRef(minRow, minCol);
            const normalizedEnd = addressToCellRef(maxRow, maxCol);
            if (!normalizedStart || !normalizedEnd) {
              return endRef;
            }
            return `${normalizedStart}:${normalizedEnd}`;
          }

          function computeBoundsFromCellRefs(startRef, endRef) {
            if (!startRef || !endRef) {
              return null;
            }
            const startAddress = cellRefToAddress(startRef);
            const endAddress = cellRefToAddress(endRef);
            if (!startAddress || !endAddress) {
              return null;
            }
            const minRow = Math.min(startAddress[0], endAddress[0]);
            const maxRow = Math.max(startAddress[0], endAddress[0]);
            const minCol = Math.min(startAddress[1], endAddress[1]);
            const maxCol = Math.max(startAddress[1], endAddress[1]);
            return {
              minRow,
              maxRow,
              minCol,
              maxCol,
              rowCount: maxRow - minRow + 1,
              colCount: maxCol - minCol + 1
            };
          }

          function updateEditModeRangeHighlight(startRef, endRef) {
            if (!isEditMode) {
              return;
            }
            const bounds = computeBoundsFromCellRefs(startRef, endRef);
            if (!bounds) {
              clearEditModeRangeHighlight();
              return;
            }
            editModeRangeHighlight.active = true;
            editModeRangeHighlight.bounds = bounds;
            updateSelectionOverlay();
          }

          function clearEditModeRangeHighlight() {
            const wasActive = editModeRangeHighlight.active;
            editModeRangeHighlight.active = false;
            editModeRangeHighlight.bounds = null;
            if (wasActive || isEditMode) {
              updateSelectionOverlay();
            }
          }
          const MODES = Object.freeze({
            READY: 'ready',
            ENTER: 'enter',
            EDIT: 'edit'
          });
          const MODE_METADATA = {
            [MODES.READY]: { prefix: "🔎\uFE0E Selecting" },
            [MODES.ENTER]: { prefix: "📥\uFE0E Entering" },
            [MODES.EDIT]: { prefix: "✏\uFE0E Editing" }
          };
          let currentMode = MODES.READY;
          const EDITOR_CURSOR_TRACKING_MODES = new Set([MODES.ENTER, MODES.EDIT]);
          const EDITOR_STATUS_PLACEHOLDER = "--";

          function formatEditorStatusValue(value) {
            if (typeof value !== "number" || !Number.isFinite(value)) {
              return EDITOR_STATUS_PLACEHOLDER;
            }
            return Math.max(0, Math.floor(value)).toString().padStart(3, "0");
          }

          function updateEditorStatusDisplay(positionOverride = null) {
            if (!editorStatusLineElement || !editorStatusColumnElement) {
              return;
            }

            const shouldTrack = EDITOR_CURSOR_TRACKING_MODES.has(currentMode);
            if (!shouldTrack) {
              editorStatusBarElement?.classList.remove("is-active");
              editorStatusLineElement.textContent = EDITOR_STATUS_PLACEHOLDER;
              editorStatusColumnElement.textContent = EDITOR_STATUS_PLACEHOLDER;
              return;
            }

            const activePosition =
              positionOverride ||
              (editor && typeof editor.getPosition === "function" ? editor.getPosition() : null);

            editorStatusBarElement?.classList.add("is-active");

            if (!activePosition) {
              editorStatusLineElement.textContent = EDITOR_STATUS_PLACEHOLDER;
              editorStatusColumnElement.textContent = EDITOR_STATUS_PLACEHOLDER;
              return;
            }

            editorStatusLineElement.textContent = formatEditorStatusValue(activePosition.lineNumber);
            editorStatusColumnElement.textContent = formatEditorStatusValue(activePosition.column);
          }

          const modePrefixElement = document.getElementById("cellModePrefix");
          const rangeChipElement = document.getElementById("cellRangeChip");
          let currentSelectionLabel = "";
          window.selectedCell = null;
          const fillPreviewCells = new Set();
          let fillHandleDragState = null;
          const selectionOverlayElement = document.getElementById("selectionOverlay");
          const selectionDragHandles = selectionOverlayElement ? selectionOverlayElement.querySelectorAll(".selection-border-handle") : [];
          if (selectionDragHandles && selectionDragHandles.length > 0) {
            selectionDragHandles.forEach(handle => {
              handle.addEventListener("mousedown", (event) => {
                if (event.button !== 0) {
                  return;
                }
                if (isEditMode) {
                  return;
                }
                if (!selectedCells || selectedCells.size === 0) {
                  return;
                }
                startSelectionDrag(event);
              });
            });
          }

          function getRowHeaderCell(rowIndex) {
            if (rowIndex === null || rowIndex === undefined) return null;
            return gridBody.querySelector(`tr[data-row="${rowIndex}"] td:first-child`);
          }

          function getColumnHeaderCell(colIndex) {
            if (colIndex === null || colIndex === undefined) return null;
            if (!gridHeaderElement) return null;
            return gridHeaderElement.querySelector(`th[data-col-index="${colIndex}"]`);
          }

          function clearHeaderHighlights() {
            highlightedRowHeaders.forEach(cell => cell.classList.remove("header-highlight"));
            highlightedColumnHeaders.forEach(cell => cell.classList.remove("header-highlight"));
            highlightedRowHeaders.clear();
            highlightedColumnHeaders.clear();
          }

          function applyHeaderHighlights(minRow, maxRow, minCol, maxCol) {
            clearHeaderHighlights();

            if (typeof minRow === "number" && typeof maxRow === "number") {
              for (let row = minRow; row <= maxRow; row++) {
                const rowHeaderCell = getRowHeaderCell(row);
                if (rowHeaderCell) {
                  rowHeaderCell.classList.add("header-highlight");
                  highlightedRowHeaders.add(rowHeaderCell);
                }
              }
            }

            if (typeof minCol === "number" && typeof maxCol === "number") {
              for (let col = minCol; col <= maxCol; col++) {
                const colHeaderCell = getColumnHeaderCell(col);
                if (colHeaderCell) {
                  colHeaderCell.classList.add("header-highlight");
                  highlightedColumnHeaders.add(colHeaderCell);
                }
              }
            }
          }

          function updateHeaderHighlightsFromSelection() {
            if (!selectedCells || selectedCells.size === 0) {
              clearHeaderHighlights();
              return;
            }

            const rows = [];
            const cols = [];

            selectedCells.forEach(cell => {
              const rowAttr = cell.getAttribute("data-row");
              const fallbackRowAttr = cell.getAttribute("data-row-index");
              const colAttr = cell.getAttribute("data-col");
              const fallbackColAttr = cell.getAttribute("data-col-index");

              const rowValue = rowAttr ?? fallbackRowAttr;
              const colValue = colAttr ?? fallbackColAttr;

              if (rowValue !== null && rowValue !== undefined) {
                const parsedRow = parseInt(rowValue, 10);
                if (!Number.isNaN(parsedRow)) {
                  rows.push(parsedRow);
                }
              }

              if (colValue !== null && colValue !== undefined) {
                const parsedCol = parseInt(colValue, 10);
                if (!Number.isNaN(parsedCol)) {
                  cols.push(parsedCol);
                }
              }
            });

            const hasRowSelection = rows.length > 0;
            const hasColSelection = cols.length > 0;

            if (!hasRowSelection && !hasColSelection) {
              clearHeaderHighlights();
              return;
            }

            const minRow = hasRowSelection ? Math.min(...rows) : null;
            const maxRow = hasRowSelection ? Math.max(...rows) : null;
            const minCol = hasColSelection ? Math.min(...cols) : null;
            const maxCol = hasColSelection ? Math.max(...cols) : null;

            applyHeaderHighlights(minRow, maxRow, minCol, maxCol);
          }

          function updateModeIndicatorElements() {
            const metadata = MODE_METADATA[currentMode] || MODE_METADATA[MODES.READY];
            if (!metadata) return;
            if (modePrefixElement) {
              modePrefixElement.textContent = metadata.prefix || "";
            }
            if (rangeChipElement) {
              const label = currentSelectionLabel ?? "";
              rangeChipElement.textContent = label;
            }
          }

          function setMode(newMode) {
            if (!MODE_METADATA[newMode]) {
              console.warn("Unknown mode requested:", newMode);
              return;
            }
            const previousMode = currentMode;
            currentMode = newMode;
            isEditMode = newMode !== MODES.READY;
            const wasEditingMode = previousMode === MODES.ENTER || previousMode === MODES.EDIT;
            const isEditingMode = newMode === MODES.ENTER || newMode === MODES.EDIT;
            if (wasEditingMode && !isEditingMode) {
              resetEnterModePointerState();
              disarmEditModePointerNavigation();
            }
            if (newMode !== MODES.EDIT) {
              disarmEditModePointerNavigation();
            }
            updateModeIndicatorElements();
            if (newMode === MODES.READY && typeof syncResultPaneWithSelection === "function") {
              requestAnimationFrame(() => syncResultPaneWithSelection());
            }
            const editorWrapper = document.querySelector('.editor-wrapper');
            if (editorWrapper) {
              if (newMode === MODES.EDIT) {
                editorWrapper.classList.add('editor-edit-mode');
              } else {
                editorWrapper.classList.remove('editor-edit-mode');
              }
            }
            updateEditorStatusDisplay();
          }

          setMode(MODES.READY);

          const REFERENCE_STATE_SEQUENCE = [
            { col: false, row: false },
            { col: true, row: true },
            { col: false, row: true },
            { col: true, row: false }
          ];

          function cycleReferenceComponent(referenceText = '') {
            if (!referenceText) return referenceText;
            const match = referenceText.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/);
            if (!match) return referenceText;
            const [, colAnchor, colLetters, rowAnchor, rowDigits] = match;

            let stateIndex = 0;
            if (colAnchor === '$' && rowAnchor === '$') {
              stateIndex = 1;
            } else if (!colAnchor && rowAnchor === '$') {
              stateIndex = 2;
            } else if (colAnchor === '$' && !rowAnchor) {
              stateIndex = 3;
            }

            const nextState = REFERENCE_STATE_SEQUENCE[(stateIndex + 1) % REFERENCE_STATE_SEQUENCE.length];
            const nextColAnchor = nextState.col ? '$' : '';
            const nextRowAnchor = nextState.row ? '$' : '';
            return `${nextColAnchor}${colLetters}${nextRowAnchor}${rowDigits}`;
          }

          function buildToggledReferenceText(originalText = '') {
            if (!originalText) return null;

            let prefix = '';
            let referencePortion = originalText;
            const lastBangIndex = originalText.lastIndexOf('!');
            if (lastBangIndex !== -1) {
              prefix = originalText.slice(0, lastBangIndex + 1);
              referencePortion = originalText.slice(lastBangIndex + 1);
            }

            const leadingWhitespace = (referencePortion.match(/^\s*/) || [''])[0];
            const trailingWhitespace = (referencePortion.match(/\s*$/) || [''])[0];
            const trimmedPortion = referencePortion.trim();
            if (!trimmedPortion) return null;

            const rangeMatch = trimmedPortion.match(/^(\$?[A-Za-z]+\$?\d+)(\s*:\s*)(\$?[A-Za-z]+\$?\d+)$/i);
            if (rangeMatch) {
              const [, firstRef, delimiter, secondRef] = rangeMatch;
              const toggledFirst = cycleReferenceComponent(firstRef);
              const toggledSecond = cycleReferenceComponent(secondRef);
              if (toggledFirst === firstRef && toggledSecond === secondRef) {
                return null;
              }
              const rebuiltRange = `${toggledFirst}${delimiter}${toggledSecond}`;
              return prefix + leadingWhitespace + rebuiltRange + trailingWhitespace;
            }

            const toggledSingle = cycleReferenceComponent(trimmedPortion);
            if (toggledSingle === trimmedPortion) {
              return null;
            }
            return prefix + leadingWhitespace + toggledSingle + trailingWhitespace;
          }

          function findReferenceAtOffset(offset, referencesList, sourceText) {
            if (!Array.isArray(referencesList) || !sourceText) {
              return null;
            }

            for (const ref of referencesList) {
              if (typeof ref.start !== 'number' || typeof ref.end !== 'number') {
                continue;
              }

              let effectiveStart = ref.start;
              let effectiveEnd = ref.end;

              if (effectiveStart > 0) {
                const precedingChar = sourceText[effectiveStart - 1];
                if (precedingChar === '$') {
                  effectiveStart -= 1;
                }
              }

              if (offset >= effectiveStart && offset <= effectiveEnd) {
                return {
                  original: ref,
                  start: effectiveStart,
                  end: effectiveEnd
                };
              }
            }

            return null;
          }

          function toggleReferenceAtCursor() {
            const currentEditor = window.monacoEditor || editor;
            if (!currentEditor || typeof currentEditor.getModel !== 'function') {
              return false;
            }
            const model = currentEditor.getModel();
            if (!model) {
              return false;
            }
            const cursorPosition = currentEditor.getPosition();
            if (!cursorPosition) {
              return false;
            }
            const cursorOffset = model.getOffsetAt(cursorPosition);
            const fullValue = model.getValue();
            if (typeof fullValue !== 'string' || fullValue.trim() === '') {
              return false;
            }
            const references = detectCellReferences(fullValue);
            if (!Array.isArray(references) || references.length === 0) {
              return false;
            }
            const locatedRef = findReferenceAtOffset(cursorOffset, references, fullValue);
            if (!locatedRef) {
              return false;
            }
            const originalText = fullValue.slice(locatedRef.start, locatedRef.end);
            const toggledText = buildToggledReferenceText(originalText);
            if (!toggledText || toggledText === originalText) {
              return false;
            }

            const startPos = model.getPositionAt(locatedRef.start);
            const endPos = model.getPositionAt(locatedRef.end);
            const editRange = new monaco.Range(
              startPos.lineNumber,
              startPos.column,
              endPos.lineNumber,
              endPos.column
            );

            window.isProgrammaticCursorChange = true;
            currentEditor.executeEdits('toggle-reference-absolute', [
              {
                range: editRange,
                text: toggledText
              }
            ]);

            const newStartPos = model.getPositionAt(locatedRef.start);
            const newEndPos = model.getPositionAt(locatedRef.start + toggledText.length);
            currentEditor.setSelection(new monaco.Selection(
              newStartPos.lineNumber,
              newStartPos.column,
              newEndPos.lineNumber,
              newEndPos.column
            ));

            Promise.resolve().then(() => {
              window.isProgrammaticCursorChange = false;
              if (typeof syncEditingCellDisplayWithPane === 'function') {
                requestAnimationFrame(() => syncEditingCellDisplayWithPane(true, currentEditor));
              }
            });

            return true;
          }

          function cancelFormulaSelection() {}

          function normalizeCellReference(ref = '') {
            if (!ref) return '';
            const withoutSheet = ref.includes('!') ? ref.substring(ref.indexOf('!') + 1) : ref;
            return withoutSheet.replace(/\$/g, '').toUpperCase();
          }

          function expandReferenceToCells(reference = '') {
            const refs = [];
            if (!reference) return refs;
            const normalized = normalizeCellReference(reference);
            if (!normalized) return refs;

            if (!normalized.includes(':')) {
              refs.push(normalized);
              return refs;
            }

            const parts = normalized.split(':').filter(Boolean);
            if (parts.length !== 2) {
              return refs;
            }

            const startAddress = cellRefToAddress(parts[0]);
            const endAddress = cellRefToAddress(parts[1]);
            if (!startAddress || !endAddress) {
              return refs;
            }

            const [startRow, startCol] = startAddress;
            const [endRow, endCol] = endAddress;
            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);

            for (let row = minRow; row <= maxRow; row++) {
              for (let col = minCol; col <= maxCol; col++) {
                refs.push(addressToCellRef(row, col));
              }
            }

            return refs;
          }

          function updateGridHighlights(assignments = new Map()) {
            if (!gridBody) return;

            const normalizedAssignments = new Map();
            assignments.forEach((value, cellRef) => {
              if (!cellRef) return;
              if (value && typeof value === 'object') {
                const normalizedInfo = {
                  className: value.className || '',
                  groupKey: value.groupKey || ''
                };
                normalizedAssignments.set(cellRef, normalizedInfo);
              } else if (value) {
                normalizedAssignments.set(cellRef, { className: value, groupKey: '' });
              }
            });

            const groupMembership = new Map();
            normalizedAssignments.forEach((info, cellRef) => {
              const resolvedGroupKey = info.groupKey || cellRef;
              info.groupKey = resolvedGroupKey;
              if (!groupMembership.has(resolvedGroupKey)) {
                groupMembership.set(resolvedGroupKey, new Set());
              }
              groupMembership.get(resolvedGroupKey).add(cellRef);
            });

            const previousEntries = Array.from(chipHighlightedCells.entries());
            previousEntries.forEach(([cellRef, className]) => {
              if (!normalizedAssignments.has(cellRef)) {
                const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
                if (cell) {
                  cell.classList.remove('chip-selector');
                  cell.removeAttribute('data-chip-selector-edge');
                  if (className) {
                    cell.classList.remove(className);
                  }
                }
                chipHighlightedCells.delete(cellRef);
              }
            });

            if (currentMode === MODES.READY) {
              return;
            }

            const directions = [
              { edge: 'top', deltaRow: -1, deltaCol: 0 },
              { edge: 'bottom', deltaRow: 1, deltaCol: 0 },
              { edge: 'left', deltaRow: 0, deltaCol: -1 },
              { edge: 'right', deltaRow: 0, deltaCol: 1 }
            ];
            const allEdges = directions.map(d => d.edge);

            function getEdgesForCell(cellRef, info) {
              const address = cellRefToAddress(cellRef);
              if (!address) {
                return allEdges;
              }
              const [row, col] = address;
              const groupCells = groupMembership.get(info.groupKey);
              if (!groupCells || groupCells.size <= 1) {
                return allEdges;
              }

              const edges = [];
              directions.forEach(({ edge, deltaRow, deltaCol }) => {
                const neighborRow = row + deltaRow;
                const neighborCol = col + deltaCol;
                if (neighborRow < 0 || neighborCol < 0) {
                  edges.push(edge);
                  return;
                }
                const neighborRef = addressToCellRef(neighborRow, neighborCol);
                if (!neighborRef || !groupCells.has(neighborRef)) {
                  edges.push(edge);
                }
              });

              return edges;
            }

            normalizedAssignments.forEach((info, cellRef) => {
              const className = info.className;
              if (!className) return;
              const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
              if (!cell) return;
              const existingClass = chipHighlightedCells.get(cellRef);
              if (existingClass && existingClass !== className) {
                cell.classList.remove(existingClass);
              }
              if (!cell.classList.contains(className)) {
                cell.classList.add(className);
              }
              cell.classList.add('chip-selector');
              const edges = getEdgesForCell(cellRef, info);
              if (edges.length > 0) {
                cell.setAttribute('data-chip-selector-edge', edges.join(' '));
              } else {
                cell.removeAttribute('data-chip-selector-edge');
              }
              chipHighlightedCells.set(cellRef, className);
            });
          }

          function resetChipColorAssignments() {
            cellChipColorAssignments.clear();
            nextChipColorIndex = 0;
            updateGridHighlights(new Map());
          }

          function getChipColorKey(refKey = '') {
            if (!refKey) return '';
            const normalized = refKey.trim().toUpperCase();
            if (!normalized) return '';
            const colonIndex = normalized.indexOf(':');
            if (colonIndex !== -1) {
              const left = normalized.substring(0, colonIndex);
              return left || normalized;
            }
            return normalized;
          }

          function getChipColorClass(refKey = '') {
            if (!refKey || chipColorPalette.length === 0) {
              return '';
            }

            const normalizedKey = getChipColorKey(refKey);
            if (!normalizedKey) {
              return '';
            }
            if (!cellChipColorAssignments.has(normalizedKey)) {
              const colorInfo = chipColorPalette[nextChipColorIndex % chipColorPalette.length];
              cellChipColorAssignments.set(normalizedKey, colorInfo.className);
              nextChipColorIndex++;
            }
            return cellChipColorAssignments.get(normalizedKey) || '';
          }

          function preserveChipColorAssignment(previousReference = '', newReference = '') {
            if (!previousReference || !newReference) {
              return;
            }
            const previousKey = getChipColorKey(previousReference.trim());
            const newKey = getChipColorKey(newReference.trim());
            if (!previousKey || !newKey || previousKey === newKey) {
              return;
            }
            const existingClass = cellChipColorAssignments.get(previousKey);
            if (!existingClass) {
              return;
            }
            cellChipColorAssignments.set(newKey, existingClass);
          }

          if (!clappyTranscript || !clappyForm || !clappyInput) {
            console.warn("Clappy elements not found, continuing without chat UI.");
          }

          // Load Monaco in background - don't wait for it
          if (monacoEditorContainer) {
            setTimeout(() => {
              console.log("Loading Monaco Editor...");
              import('monaco-editor').then(monaco => {
                console.log("Monaco Editor loaded successfully");

                // Store monaco globally so it's accessible in selectCell and other functions
                window.monaco = monaco;

                // MonacoEnvironment should already be set above, before the import

                try {
                  // Define theme for excel-formula language
                  monaco.editor.defineTheme("excel-formula-dark", {
                    base: "vs-dark",
                    inherit: true,
                    rules: [
                      { token: "", foreground: "D4D4D4" }, // Default text color
                      { token: "comment", foreground: "858585", fontStyle: "italic" }, // Soft grey for comments
                      { token: "keyword", foreground: "569CD6" },
                      { token: "string", foreground: "D19A66" },
                      { token: "string.escape", foreground: "D19A66" },
                      { token: "string.quote", foreground: "D19A66" },
                      { token: "delimiter", foreground: "D19A66" },
                      { token: "delimiter.string", foreground: "D19A66" },
                      { token: "delimiter.bracket", foreground: "D4D4D4" },
                      { token: "delimiter.parenthesis", foreground: "D4D4D4" },
                      { token: "number", foreground: "B5CEA8" },
                      { token: "type", foreground: "4EC9B0" },
                      { token: "function", foreground: "569CD6" }, // Blue color like JSON delimiters
                      { token: "variable", foreground: "9CDCFE" },
                      { token: "identifier", foreground: "9CDCFE" },
                      { token: "cell-reference", foreground: "9CDCFE" },
                      { token: "operator", foreground: "D4D4D4" },
                    ],
                    colors: {
                      "editor.background": "#000000",
                      "editor.foreground": "#D4D4D4",
                      "editorLineNumber.foreground": "#858585",
                      "editor.selectionBackground": "#264F78",
                      "editorCursor.foreground": "#AEAFAD",
                      "editor.lineHighlightBackground": "#000000",
                      "editorIndentGuide.background": "#404040",
                      "editorIndentGuide.activeBackground": "#707070",
                      "editorError.foreground": "#FF5C57",
                      "editorWarning.foreground": "#FFB454",
                      "editorInfo.foreground": "#5CC5FF",
                      "editorError.border": "#000000",
                      "editorWarning.border": "#000000",
                      "editorInfo.border": "#000000"
                    }
                  });

                  // Register excel-formula language
                  monaco.languages.register({ id: "excel-formula" });

                  // Load function list and set up tokenizer
                  import('../hyperformula-functions-monaco.js').then(({ hyperFormulaFunctions }) => {
                    const extendedFunctions = extendWithCustomFunctions(hyperFormulaFunctions);
                    console.log(`Setting up tokenizer with ${extendedFunctions.length} functions (including custom extensions)`);

                    // Store function list globally for auto-capitalization and completion
                    window.hyperFormulaFunctions = extendedFunctions;

                    // Build regex pattern from all function names
                    const functionNames = extendedFunctions.map(f => f.name);
                    // Escape special regex characters and join with |
                    const functionPattern = functionNames
                      .map(name => name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
                      .join('|');

                  // Set up tokenization rules for excel-formula
                  monaco.languages.setMonarchTokensProvider("excel-formula", {
                    tokenizer: {
                      root: [
                        // Block comments (/* ... */)
                        [/\/\*/, "comment", "@comment"],
                        // Inline comments (// to end of line)
                        [/\/\/.*$/, "comment"],
                        // Cell references (e.g., A1, B2, $A$1, Sheet1!A1)
                        [/[A-Z]+\$?[0-9]+\$?|'[^']*'![A-Z]+\$?[0-9]+\$?/, "cell-reference"],
                          // Functions - use all HyperFormula function names
                          [new RegExp(`\\b(${functionPattern})\\b`, 'i'), "function"],
                        // Numbers (integers and decimals)
                        [/\d+\.?\d*/, "number"],
                        // Strings (text in quotes)
                        [/"([^"\\]|\\.)*"/, "string"],
                        // Operators
                        [/[+\-*/=<>]/, "operator"],
                        // Parentheses and brackets
                        [/[()]/, "delimiter.parenthesis"],
                        [/[[\]]/, "delimiter.bracket"],
                        // Commas and semicolons
                        [/[,;]/, "delimiter"],
                        // Whitespace
                        [/\s+/, ""],
                        // Everything else
                        [/./, ""]
                      ],
                      comment: [
                        // End of block comment
                        [/\*\//, "comment", "@pop"],
                        // Continue in comment (any character)
                        [/./, "comment"]
                      ]
                    }
                    });

                    console.log('Tokenizer set up with function highlighting');
                  }).catch(err => {
                    console.error('Failed to load HyperFormula functions for tokenizer:', err);
                    // Fallback to basic tokenizer if function list fails to load
                    monaco.languages.setMonarchTokensProvider("excel-formula", {
                      tokenizer: {
                        root: [
                          [/[A-Z]+\$?[0-9]+\$?|'[^']*'![A-Z]+\$?[0-9]+\$?/, "cell-reference"],
                          [/\d+\.?\d*/, "number"],
                          [/"([^"\\]|\\.)*"/, "string"],
                          [/[+\-*/=<>]/, "operator"],
                          [/[()]/, "delimiter.parenthesis"],
                          [/[[\]]/, "delimiter.bracket"],
                          [/[,;]/, "delimiter"],
                          [/\s+/, ""],
                          [/./, ""]
                        ]
                      }
                    });
                  });

                  // Set up language configuration
                  monaco.languages.setLanguageConfiguration("excel-formula", {
                    comments: {
                      lineComment: "//",
                      blockComment: ["/*", "*/"]
                    },
                    brackets: [
                      ["(", ")"],
                      ["[", "]"]
                    ],
                    autoClosingPairs: [
                      { open: "(", close: ")" },
                      { open: "[", close: "]" },
                      { open: '"', close: '"' }
                    ],
                    surroundingPairs: [
                      { open: "(", close: ")" },
                      { open: "[", close: "]" },
                      { open: '"', close: '"' }
                    ]
                  });

                  // Allow editor to stretch to available height (Monaco adds its own scrollbars)
                  monacoEditorContainer.style.height = "100%";
                  monacoEditorContainer.style.minHeight = "0px";

                  // Create initial value with default padded lines
                  const initialValue = ensureSevenLines("");

                    editor = monaco.editor.create(monacoEditorContainer, {
                    value: initialValue,
                    language: "excel-formula",
                    theme: "excel-formula-dark",
                    cursorBlinking: 'hidden', // Hide cursor initially (will show in edit mode)
                    automaticLayout: false,
                    fontSize: 12,
                    fontFamily: "Consolas, 'Courier New', monospace",
                    lineNumbers: 'on',
                    minimap: { enabled: false },
                      scrollBeyondLastLine: false,
                      readOnly: false,
                    contextmenu: false,
                    wordWrap: "off", // Disable word wrap - use horizontal scrolling instead
                    lineHeight: 1.5,
                    tabSize: 4,
                    insertSpaces: true, // Use spaces instead of tab characters
                    renderLineHighlight: "none",
                    overviewRulerBorder: false,
                    hideCursorInOverviewRuler: true,
                    overviewRulerLanes: 0,
                    renderIndentGuides: false,
                    selectionHighlight: true,
                    occurrencesHighlight: true,
                    guides: {
                      indentation: false,
                      highlightActiveIndentation: false,
                      bracketPairsHorizontal: false,
                      highlightActiveBracketPair: false
                    },
                    wordBasedSuggestions: false,
                    folding: true,
                    foldingStrategy: "indentation",
                    showFoldingControls: "always",
                    unfoldOnClickAfterEndOfLine: true,
                    foldingHighlight: true,
                    scrollbar: {
                      vertical: "auto",
                      horizontal: "auto",
                      useShadows: false,
                      verticalHasArrows: false,
                      horizontalHasArrows: false,
                    },
                    // Configure suggest widget to be constrained to editor container
                    suggest: {
                      showKeywords: false,
                      showSnippets: false,
                      showClasses: false,
                      showFunctions: true,
                      showVariables: false,
                      showModules: false,
                      showProperties: false,
                      showValues: false,
                      showEnums: false,
                      showStructs: false,
                      showInterfaces: false,
                      showOperators: false,
                      showUnits: false,
                      showColors: false,
                      showFiles: false,
                      showReferences: false,
                      showFolders: false,
                      showTypeParameters: false,
                      showIssues: false,
                      showUsers: false,
                      showText: false,
                      maxVisibleSuggestions: 12,
                      filterGraceful: true,
                      shareSuggestSelections: false,
                      showIcons: true,
                      showStatusBar: true,
                      preview: true,
                      previewMode: 'prefix',
                      quickSuggestions: {
                        other: true, // Enable suggestions as you type
                        comments: false,
                        strings: false
                      },
                      suggestOnTriggerCharacters: false, // Don't trigger on special characters
                      acceptSuggestionOnCommitCharacter: true,
                      acceptSuggestionOnEnter: 'off', // Enter always goes to new line, never accepts suggestions
                      acceptSuggestionOnTab: 'on',
                      snippetSuggestions: 'top',
                      hideStatusBar: false
                    },
                    // Configure parameter hints
                    parameterHints: {
                      enabled: true,
                      cycle: false
                    }
                  });

                  // Ensure theme is applied
                  monaco.editor.setTheme("excel-formula-dark");

                  // No longer need first line decoration since "=" is in the header

                  // Track if we're programmatically setting cursor (to avoid event loop)
                  window.isProgrammaticCursorChange = false;
                  window.cursorPositionListener = null;

                  // Line 1 is now a valid line - no restrictions needed
                  // Removed old cursor position listener that was preventing line 1 interaction

                  // Set initial cursor position to line 1
                  editor.setPosition({ lineNumber: 1, column: 1 });
                  updateEditorStatusDisplay();

                  // Throttled layout update for performance
                  let layoutTimeout = null;
                  const updateLayout = () => {
                    if (layoutTimeout) return;
                    layoutTimeout = requestAnimationFrame(() => {
                      editor.layout();
                      layoutTimeout = null;
                    });
                  };

                  // Only update layout when the editor container actually resizes
                  const resizeObserver = new ResizeObserver(() => {
                    updateLayout();
                  });
                  resizeObserver.observe(monacoEditorContainer);

                  // Initial layout
                  editor.layout();

                  // Constrain suggest widget dimensions to formula editor container
                  // This ensures autocomplete popup dimensions are constrained to formula pane and doesn't overflow into grid
                  const formulaEditor = document.querySelector('.formula-editor');
                  const editorWrapper = document.querySelector('.editor-wrapper');

                  if (formulaEditor && editorWrapper) {
                    // Function to constrain suggest widget - called every time it appears
                    const constrainSuggestWidget = () => {
                      const suggestWidget = document.querySelector('.formula-editor .monaco-editor .suggest-widget');
                      if (!suggestWidget || suggestWidget.style.display === 'none' || suggestWidget.style.visibility === 'hidden') {
                        return;
                      }

                      // Use requestAnimationFrame to ensure DOM has updated
                        requestAnimationFrame(() => {
                        // Get bounding rectangles
                          const formulaRect = formulaEditor.getBoundingClientRect();
                          const wrapperRect = editorWrapper.getBoundingClientRect();
                          const widgetRect = suggestWidget.getBoundingClientRect();
                          const editorRect = monacoEditorContainer.getBoundingClientRect();

                        if (!formulaRect || !widgetRect || !editorRect) return;

                        // Calculate available space - must stay within formula editor container
                          const margin = 10;

                        // Maximum width: formula editor width minus margins
                        // This prevents overflow into the grid area - use the actual formula editor width
                        const maxWidth = Math.min(
                          formulaRect.width - (margin * 2),
                          wrapperRect.width - (margin * 2),
                          400 // Absolute maximum width
                        );

                        // Maximum height: available space in wrapper, but not more than 300px
                        const maxHeight = Math.min(
                          wrapperRect.height - 20, // Leave 20px margin
                          formulaRect.height - 60, // Leave space for header and editor
                          300 // Absolute max
                        );

                          // Get current widget dimensions
                          const currentWidth = widgetRect.width || 300;
                          const currentHeight = widgetRect.height || 200;

                        // Calculate widget position relative to formula editor
                        const widgetLeftRelative = widgetRect.left - formulaRect.left;
                        const widgetRightRelative = widgetRect.right - formulaRect.left;

                        // Constrain width to fit within formula editor (don't exceed right edge)
                          let constrainedWidth = Math.min(currentWidth, maxWidth);

                        // Ensure right edge doesn't exceed formula editor bounds
                        const maxRightInFormula = formulaRect.width - margin;
                        const widgetRightInFormula = widgetLeftRelative + constrainedWidth;

                        // Calculate position relative to Monaco editor for setting left/top
                        const editorLeftRelative = editorRect.left - formulaRect.left;
                        const editorTopRelative = editorRect.top - formulaRect.top;

                        if (widgetRightInFormula > maxRightInFormula) {
                          // Widget extends too far right - adjust left position
                          let adjustedLeftInFormula = maxRightInFormula - constrainedWidth;
                          // But don't go too far left
                          adjustedLeftInFormula = Math.max(margin, adjustedLeftInFormula);

                          // If we still can't fit, reduce width instead
                          if ((adjustedLeftInFormula + constrainedWidth) > maxRightInFormula) {
                            constrainedWidth = maxRightInFormula - adjustedLeftInFormula;
                          }

                          // Apply adjusted left position relative to Monaco editor
                          suggestWidget.style.setProperty('left', `${adjustedLeftInFormula - editorLeftRelative}px`, 'important');
                        } else {
                          // Ensure left edge doesn't go too far left
                          const minLeftInFormula = margin;
                          if (widgetLeftRelative < minLeftInFormula) {
                            suggestWidget.style.setProperty('left', `${minLeftInFormula - editorLeftRelative}px`, 'important');
                          }
                        }

                        // Constrain height
                        let constrainedHeight = Math.min(currentHeight, maxHeight);

                        // Check vertical position - ensure it doesn't overflow bottom of formula editor
                        const widgetTopInFormula = widgetRect.top - formulaRect.top;
                        const widgetBottomInFormula = widgetTopInFormula + constrainedHeight;
                        const maxBottomInFormula = formulaRect.height - margin;

                        if (widgetBottomInFormula > maxBottomInFormula) {
                          // Try to move widget up
                          const adjustedTopInFormula = maxBottomInFormula - constrainedHeight;
                          if (adjustedTopInFormula >= margin) {
                            // Can move up, adjust position
                            suggestWidget.style.setProperty('top', `${adjustedTopInFormula - editorTopRelative}px`, 'important');
                          } else {
                            // Can't move up enough, reduce height instead
                            constrainedHeight = Math.max(50, maxBottomInFormula - widgetTopInFormula);
                            suggestWidget.style.setProperty('top', `${margin - editorTopRelative}px`, 'important');
                          }
                        }

                        // Apply all constraints with !important to override Monaco's styles
                        suggestWidget.style.setProperty('max-width', `${constrainedWidth}px`, 'important');
                        suggestWidget.style.setProperty('width', `${constrainedWidth}px`, 'important');
                        suggestWidget.style.setProperty('max-height', `${constrainedHeight}px`, 'important');
                        suggestWidget.style.setProperty('height', `${constrainedHeight}px`, 'important');
                        suggestWidget.style.setProperty('overflow-y', 'auto', 'important');
                        suggestWidget.style.setProperty('overflow-x', 'hidden', 'important');
                        suggestWidget.style.setProperty('z-index', '10000', 'important');
                        suggestWidget.style.setProperty('position', 'absolute', 'important');

                        // Force clipping by ensuring right edge doesn't exceed bounds
                        // Check one more time after setting width
                        setTimeout(() => {
                          const updatedWidgetRect = suggestWidget.getBoundingClientRect();
                          const updatedRightInFormula = updatedWidgetRect.right - formulaRect.left;
                          if (updatedRightInFormula > maxRightInFormula) {
                            const overflowAmount = updatedRightInFormula - maxRightInFormula;
                            const newWidth = constrainedWidth - overflowAmount - margin;
                            if (newWidth > 100) {
                              suggestWidget.style.setProperty('width', `${newWidth}px`, 'important');
                              suggestWidget.style.setProperty('max-width', `${newWidth}px`, 'important');
                            }
                          }
                        }, 0);
                      });
                    };

                    // Use MutationObserver to catch when suggest widget appears/changes
                    const observer = new MutationObserver(() => {
                      constrainSuggestWidget();
                    });

                    // Observe the editor container for suggest widget changes
                    observer.observe(monacoEditorContainer, {
                      childList: true,
                      subtree: true,
                      attributes: true,
                      attributeFilter: ['style', 'class']
                    });

                    // Also observe the formula editor for size changes
                    observer.observe(formulaEditor, {
                      childList: true,
                      subtree: true,
                      attributes: true,
                      attributeFilter: ['style', 'class']
                    });

                    // Also listen for resize events to re-constrain
                    const resizeObserver = new ResizeObserver(() => {
                      constrainSuggestWidget();
                    });
                    resizeObserver.observe(formulaEditor);
                    resizeObserver.observe(monacoEditorContainer);

                    // Periodically check and constrain (fallback for cases MutationObserver might miss)
                    // Use more frequent checking to catch widget immediately when it appears
                    const constraintInterval = setInterval(() => {
                      constrainSuggestWidget();
                    }, 50); // Check every 50ms for faster constraint application

                    // Store interval to clear later if needed
                    window.suggestWidgetConstraintInterval = constraintInterval;

                    // Also listen for Monaco's suggest events
                    editor.onDidChangeCursorPosition(() => {
                      setTimeout(constrainSuggestWidget, 10);
                    });

                    // Watch for when suggestions are shown/hidden
                    const suggestController = editor.getContribution('suggestController');
                    if (suggestController && suggestController.model) {
                      suggestController.model.onDidChange(() => {
                        setTimeout(constrainSuggestWidget, 10);
                      });
                    }
                  }

                  // Import and register HyperFormula functions for IntelliSense
                  import('../hyperformula-functions-monaco.js').then(({ hyperFormulaFunctions }) => {
                    const extendedFunctions = extendWithCustomFunctions(hyperFormulaFunctions);
                    console.log(`Registering ${extendedFunctions.length} functions for IntelliSense (including custom extensions)`);
                    window.hyperFormulaFunctions = extendedFunctions;

                    // Register completion provider for all HyperFormula functions and punctuation
                    // createFunctionSnippet is imported from helpers
                    monaco.languages.registerCompletionItemProvider('excel-formula', {
                      provideCompletionItems: function(model, position, context) {
                        const suggestions = [];
                        const word = model.getWordUntilPosition(position);

                        // Get all lines from editor
                        const allLines = model.getValue().split('\n');
                        if (allLines.length < 1) return { suggestions };

                        const currentLineIndex = position.lineNumber - 1; // Convert to 0-based
                        // Line 1 is now valid, so no need to skip it

                        const currentLine = allLines[currentLineIndex];
                        const textBeforeCursor = currentLine.substring(0, position.column - 1);

                        // Get the text at the cursor position to better detect what's being typed
                        const textUntilPosition = model.getValueInRange({
                          startLineNumber: position.lineNumber,
                          startColumn: 1,
                          endLineNumber: position.lineNumber,
                          endColumn: position.column
                        });

                        // Determine whether the cursor is currently inside an open string literal
                        const textFromDocStart = model.getValueInRange({
                          startLineNumber: 1,
                          startColumn: 1,
                          endLineNumber: position.lineNumber,
                          endColumn: position.column
                        });

                        const isInsideStringLiteral = (() => {
                          let inString = false;
                          for (let i = 0; i < textFromDocStart.length; i++) {
                            const char = textFromDocStart[i];
                            if (char === '"') {
                              const nextChar = textFromDocStart[i + 1];
                              if (inString && nextChar === '"') {
                                // Escaped double quote within a string (""), skip next char
                                i++;
                                continue;
                              }
                              inString = !inString;
                            }
                          }
                          return inString;
                        })();

                        // Don't show IntelliSense suggestions while inside a string literal
                        if (isInsideStringLiteral) {
                          return { suggestions };
                        }

                        // Check if we're in a function call context for comma/parenthesis completion
                        const functionCallPattern = /\b([A-Za-z]+\w*)\s*\([^)]*$/i;
                        const functionMatch = textBeforeCursor.match(functionCallPattern);
                        const isInFunctionCall = functionMatch !== null;

                        // Always provide comma and closing parenthesis suggestions (for Tab completion)
                        // These will be available via Tab key but won't show popup automatically
                        if (isInFunctionCall) {
                          const { openParens, closeParens } = analyzeParentheses(textBeforeCursor);
                          const hasUnclosedParens = openParens > closeParens;

                          if (hasUnclosedParens) {
                            // Patterns: after string literal, number, cell reference, expression, or closing parenthesis
                            // Note: Single quotes are reserved by Excel for sheet names, so only check double quotes for strings
                            const afterStringPattern = /"[^"]*"\s*$/;
                            const afterNumberPattern = /(\d+(\.\d+)?)\s*$/;
                            const afterCellRefPattern = /([A-Z]+\$?\d+\$?)\s*$/i;
                            // Expression pattern: value operator value (like A1 = 2, A1 + 1, etc.)
                            // Note: Single quotes are reserved by Excel for sheet names
                            const afterExpressionPattern = /([A-Z]+\$?\d+\$?|\d+|"[^"]*")\s*[+\-*/=<>!]+\s*([A-Z]+\$?\d+\$?|\d+|"[^"]*")\s*$/i;
                            const afterClosingParenPattern = /\)\s*$/;

                            // Check if we need a comma (after a complete argument)
                            const needsComma = afterStringPattern.test(textBeforeCursor) || 
                                              afterNumberPattern.test(textBeforeCursor) || 
                                              afterCellRefPattern.test(textBeforeCursor) ||
                                              afterExpressionPattern.test(textBeforeCursor) ||
                                              afterClosingParenPattern.test(textBeforeCursor);

                            if (needsComma) {
                              // Add comma as tab-completable suggestion (highest priority)
                              suggestions.push({
                                label: ',',
                                kind: monaco.languages.CompletionItemKind.Snippet,
                                insertText: ', ',
                                detail: 'Comma separator',
                                documentation: 'Press Tab to add comma',
                                range: {
                                  startLineNumber: position.lineNumber,
                                  endLineNumber: position.lineNumber,
                                  startColumn: position.column,
                                  endColumn: position.column
                                },
                                sortText: '0000', // Highest priority
                                preselect: true // Auto-select for Tab
                              });
                            }

                            // Check if we need a closing parenthesis
                            // Get the function call text
                            const functionCallText = textBeforeCursor.substring(functionMatch.index);
                            const argsMatch = functionCallText.match(/\((.+)$/);

                            if (argsMatch) {
                              const args = argsMatch[1];
                              // Count commas to see how many arguments we have
                              const argCount = (args.match(/,/g) || []).length + 1;

                              // Check if we're at the end of what looks like a complete argument
                              const trimmedArgs = args.trim();
                              // Match: ends with value, cell ref, string, number, or expression
                              // Note: Single quotes are reserved by Excel for sheet names
                              const endsWithCompleteArg = /([A-Z]+\$?\d+\$?|\d+|"[^"]*"|\)|[A-Z]+\$?\d+\$?|\d+|"[^"]*")\s*[+\-*/=<>!]+\s*([A-Z]+\$?\d+\$?|\d+|"[^"]*")\s*$/i.test(trimmedArgs) ||
                                                          /([A-Z]+\$?\d+\$?|\d+|"[^"]*"|\))\s*$/i.test(trimmedArgs);

                              // Check if there's already a closing paren after cursor
                              const textAfterCursor = currentLine.substring(position.column - 1);
                              const hasClosingParen = textAfterCursor.trim().startsWith(')');

                              // If we have at least one complete argument and no closing paren, suggest it
                              if (endsWithCompleteArg && argCount >= 1 && !hasClosingParen) {
                                suggestions.push({
                                  label: ')',
                                  kind: monaco.languages.CompletionItemKind.Snippet,
                                  insertText: ')',
                                  detail: 'Close parenthesis',
                                  documentation: 'Press Tab to close function',
                                  range: {
                                    startLineNumber: position.lineNumber,
                                    endLineNumber: position.lineNumber,
                                    startColumn: position.column,
                                    endColumn: position.column
                                  },
                                  sortText: '0001', // Second highest priority
                                  preselect: !needsComma // Auto-select if comma not needed
                                });
                              }
                            }
                          }
                        }

                        // Now add function suggestions with snippet templates
                        // Get the actual word being typed - case insensitive
                        const wordMatch = textBeforeCursor.match(/([A-Za-z_][A-Za-z0-9_]*)$/i);
                        const actualWord = wordMatch ? wordMatch[1] : word.word;

                        // Check if the word looks like a cell reference (like "a1", "B2", "A$1", "$A$1")
                        // Pattern: 1-3 letters followed by optional $, then digits, optionally followed by $
                        const isCellReference = /^[A-Z]{1,3}\$?\d+\$?$/i.test(actualWord);

                        // Show function suggestions if typing a word
                        // Case-insensitive - allow completion even when not in function call
                        if (actualWord && actualWord.length >= 1) {
                          // Determine if we're typing a cell reference
                          // Check if what we've typed so far looks like a cell reference
                          const beforeWord = textBeforeCursor.substring(0, textBeforeCursor.length - actualWord.length);
                          const fullWordSoFar = beforeWord + actualWord;
                          // Pattern: letters (1-3) optionally followed by $ and digits
                          // This is more lenient - if it's clearly a cell ref pattern, don't show functions
                          const looksLikeCellRef = /^[A-Z]{1,3}\$?\d*$/i.test(fullWordSoFar);

                          // Don't show function suggestions if:
                          // 1. It's a complete cell reference (letters + digits), AND
                          // 2. It's not being explicitly invoked (Ctrl+Space), AND
                          // 3. It's short (2-4 chars) - longer words are more likely to be function names
                          const shouldSuppressForCellRef = isCellReference && 
                                                          context.triggerKind !== monaco.languages.CompletionTriggerKind.Invoke &&
                                                          actualWord.length <= 4;

                          // Don't show function suggestions if we're inside a function call (already typed function)
                          // But allow if we're just typing the function name
                          const isTypingFunctionName = !isInFunctionCall || textBeforeCursor.trim() === actualWord;

                          // Show suggestions if:
                          // 1. Not a cell reference suppression case, OR
                          // 2. Explicitly invoked (Ctrl+Space), OR
                          // 3. Word is long enough to be a function (3+ chars), OR
                          // 4. Doesn't look like a cell reference
                          const shouldShowSuggestions = (!shouldSuppressForCellRef || 
                                                         context.triggerKind === monaco.languages.CompletionTriggerKind.Invoke ||
                                                         actualWord.length >= 3 ||
                                                         !looksLikeCellRef) &&
                                                         isTypingFunctionName;

                          if (shouldShowSuggestions) {
                            // Use global function list if available
                            if (!window.hyperFormulaFunctions || window.hyperFormulaFunctions.length === 0) {
                              return { suggestions }; // Function list not loaded yet
                            }
                        const wordUpper = actualWord.toUpperCase();
                            const functionSuggestions = window.hyperFormulaFunctions
                          .filter(func => {
                                // Case-insensitive matching - exact match or starts with typed text
                            return func.name.toUpperCase().startsWith(wordUpper);
                          })
                          .map(func => {
                            // Calculate the correct range based on actual word
                              const wordStart = textUntilPosition.length - actualWord.length;
                            const correctRange = {
                              startLineNumber: position.lineNumber,
                              endLineNumber: position.lineNumber,
                              startColumn: wordStart + 1,
                              endColumn: position.column
                            };

                                // Create snippet template from function signature
                                // For IF: IF(logical, value, value) becomes IF(${1:logical}, ${2:value}, ${3:value})
                                const snippetTemplate = createFunctionSnippet(func);

                            return {
                              label: func.name,
                              kind: monaco.languages.CompletionItemKind.Function,
                                  insertText: snippetTemplate,
                              insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
                              documentation: {
                                value: func.description 
                                  ? `**${func.name}**\n\n${func.description}\n\n**Syntax:** \`${func.signature}\``
                                  : `**${func.name}**\n\n**Syntax:** \`${func.signature}\``
                              },
                              range: correctRange,
                                  detail: func.signature,
                                  sortText: '1000', // Lower priority than punctuation
                                  preselect: func.name.toUpperCase() === wordUpper // Auto-select on exact match
                            };
                          });

                            suggestions.push(...functionSuggestions);
                          }
                        }

                        return { suggestions };
                      },
                      triggerCharacters: [] // No automatic triggers - Tab completes suggestions
                    });

                    // Register hover provider for function documentation
                    monaco.languages.registerHoverProvider('excel-formula', {
                      provideHover: function(model, position) {
                        const word = model.getWordAtPosition(position);
                        if (!word) return null;

                        const funcName = word.word.toUpperCase();
                        const func = extendedFunctions.find(f => f.name.toUpperCase() === funcName);

                        if (func) {
                          return {
                            range: new monaco.Range(
                              position.lineNumber,
                              word.startColumn,
                              position.lineNumber,
                              word.endColumn
                            ),
                            contents: [
                              { value: `**${func.name}**` },
                              func.description ? { value: func.description } : null,
                              { value: `\`${func.signature}\`` }
                            ].filter(Boolean)
                          };
                        }
                        return null;
                      }
                    });

                    // Register diagnostics provider for formula validation
                    monaco.languages.registerDocumentFormattingEditProvider('excel-formula', {
                      provideDocumentFormattingEdits: function(model, options, token) {
                        return [];
                      }
                    });

                    // Register diagnostics provider to validate formulas
                    monaco.languages.registerDocumentRangeFormattingEditProvider('excel-formula', {
                      provideDocumentRangeFormattingEdits: function(model, range, options, token) {
                        return [];
                      }
                    });

                    // Helper caches for function metadata used by diagnostics
                    const functionSignatureInfoCache = new Map();

                    function ensureFunctionMap() {
                      if (!window.hyperFormulaFunctions || window.hyperFormulaFunctions.length === 0) {
                        return null;
                      }

                      if (!window.__hyperFormulaFunctionMap) {
                        window.__hyperFormulaFunctionMap = new Map();
                        window.hyperFormulaFunctions.forEach((func) => {
                          if (func?.name) {
                            window.__hyperFormulaFunctionMap.set(func.name.toUpperCase(), func);
                          }
                        });
                      }

                      return window.__hyperFormulaFunctionMap;
                    }

                    function parseFunctionSignature(signature = '') {
                      const metadata = {
                        minArgs: 0,
                        maxArgs: 0,
                        variadic: false
                      };

                      const match = signature.match(/^[^(]+\((.*)\)$/);
                      if (!match) {
                        return metadata;
                      }

                      const paramsString = match[1].trim();
                      if (!paramsString) {
                        return metadata;
                      }

                      const params = paramsString.split(',').map((param) => param.trim()).filter(Boolean);
                      metadata.maxArgs = params.length;

                      for (let i = 0; i < params.length; i++) {
                        const param = params[i];
                        const isVariadicParam = param.endsWith('...');
                        const normalized = param.replace(/\.\.\.$/, '').trim();
                        const isOptional = normalized.startsWith('[') && normalized.endsWith(']');

                        if (isVariadicParam) {
                          metadata.variadic = true;
                          metadata.maxArgs = Infinity;
                          // Variadic arguments are optional repeats of the previous definition
                          if (i === 0) {
                            // Most HyperFormula variadic functions require at least one argument
                            metadata.minArgs = Math.max(metadata.minArgs, 1);
                          }
                          break;
                        }

                        if (!isOptional) {
                          metadata.minArgs++;
                        }
                      }

                      return metadata;
                    }

                    function getFunctionSignatureInfo(funcName) {
                      const functionMap = ensureFunctionMap();
                      if (!functionMap || !funcName) {
                        return null;
                      }

                      const upperName = funcName.toUpperCase();
                      if (functionSignatureInfoCache.has(upperName)) {
                        return functionSignatureInfoCache.get(upperName);
                      }

                      const funcDef = functionMap.get(upperName);
                      if (!funcDef) {
                        return null;
                      }

                      const parsed = parseFunctionSignature(funcDef.signature);
                      const info = {
                        ...parsed,
                        signature: funcDef.signature,
                        displayName: funcDef.name || funcDef.signature || upperName
                      };

                      functionSignatureInfoCache.set(upperName, info);
                      return info;
                    }

                    function findClosingParenthesis(text, openIndex) {
                      let depth = 0;
                      let inString = false;
                      let escaped = false;

                      for (let i = openIndex; i < text.length; i++) {
                        const char = text[i];

                        if (escaped) {
                          escaped = false;
                          continue;
                        }

                        if (char === '\\') {
                          escaped = true;
                          continue;
                        }

                        if (char === '"') {
                          inString = !inString;
                          continue;
                        }

                        if (inString) {
                          continue;
                        }

                        if (char === '(') {
                          depth++;
                        } else if (char === ')') {
                          depth--;
                          if (depth === 0) {
                            return i;
                          }
                        }
                      }

                      return -1;
                    }

                    function parseParameterSegments(paramString) {
                      const segments = [];
                      const emptySegments = [];
                      let depth = 0;
                      let inString = false;
                      let escaped = false;
                      let segmentStart = 0;

                      for (let i = 0; i < paramString.length; i++) {
                        const char = paramString[i];

                        if (escaped) {
                          escaped = false;
                          continue;
                        }

                        if (char === '\\') {
                          escaped = true;
                          continue;
                        }

                        if (char === '"') {
                          inString = !inString;
                          continue;
                        }

                        if (!inString) {
                          if (char === '(') {
                            depth++;
                          } else if (char === ')') {
                            if (depth > 0) {
                              depth--;
                            }
                          } else if (char === ',' && depth === 0) {
                            segments.push({
                              start: segmentStart,
                              end: i,
                              text: paramString.substring(segmentStart, i)
                            });
                            segmentStart = i + 1;
                            continue;
                          }
                        }
                      }

                      segments.push({
                        start: segmentStart,
                        end: paramString.length,
                        text: paramString.substring(segmentStart)
                      });

                      const hasComma = paramString.includes(',');
                      const hasContent = paramString.trim().length > 0;

                      if (!hasComma && !hasContent) {
                        return { segments: [], emptySegments: [], actualCount: 0 };
                      }

                      segments.forEach((segment) => {
                        if (segment.text.trim().length === 0) {
                          emptySegments.push(segment);
                        }
                      });

                      return {
                        segments,
                        emptySegments,
                        actualCount: segments.length
                      };
                    }

                    function indexToEditorPosition(text, absoluteIndex) {
                      const clampedIndex = Math.max(0, Math.min(absoluteIndex, text.length));
                      const before = text.substring(0, clampedIndex);
                      const linesBefore = before.split('\n');
                      const lineNumber = linesBefore.length;
                      const lastLine = linesBefore[linesBefore.length - 1] || '';

                      return {
                        lineNumber,
                        column: lastLine.length + 1
                      };
                    }

                    const LINE_CONTINUATION_CHARS = new Set(['+', '-', '*', '/', '^', '&', '=', '<', '>', ',', '(', '{', ':', ';', '!']);

                    function analyzeParentheses(text) {
                      let openParens = 0;
                      let closeParens = 0;
                      let depth = 0;
                      let firstExtraClosingIndex = -1;
                      let inDoubleQuotes = false;
                      let inSingleQuotes = false;
                      let braceDepth = 0;

                      for (let i = 0; i < text.length; i++) {
                        const char = text[i];
                        const nextChar = text[i + 1];

                        if (!inSingleQuotes && char === '"') {
                          if (nextChar === '"') {
                            i++;
                            continue;
                          }
                          inDoubleQuotes = !inDoubleQuotes;
                          continue;
                        }

                        if (!inDoubleQuotes && char === "'") {
                          if (nextChar === "'") {
                            i++;
                            continue;
                          }
                          inSingleQuotes = !inSingleQuotes;
                          continue;
                        }

                        if (inDoubleQuotes || inSingleQuotes) {
                          continue;
                        }

                        if (char === '{') {
                          braceDepth++;
                          continue;
                        }

                        if (char === '}' && braceDepth > 0) {
                          braceDepth--;
                          continue;
                        }

                        if (braceDepth > 0) {
                          continue;
                        }

                        if (char === '(') {
                          openParens++;
                          depth++;
                        } else if (char === ')') {
                          closeParens++;
                          depth--;
                          if (depth < 0 && firstExtraClosingIndex === -1) {
                            firstExtraClosingIndex = i;
                            depth = 0;
                          }
                        }
                      }

                      return { openParens, closeParens, firstExtraClosingIndex };
                    }

                    function detectStrayTopLevelText(fullText = '') {
                      if (!fullText) {
                        return [];
                      }

                      const straySegments = [];
                      const rawLines = fullText.split('\n');
                      let parenDepth = 0;
                      let braceDepth = 0;
                      let inDoubleQuotes = false;
                      let inSingleQuotes = false;
                      let inBlockComment = false;
                      let formulaStarted = false;
                      let previousLineAllowsContinuation = true;

                      for (let lineIndex = 0; lineIndex < rawLines.length; lineIndex++) {
                        const originalLine = rawLines[lineIndex];
                        const line = originalLine.endsWith('\r') ? originalLine.slice(0, -1) : originalLine;
                        const stateBeforeLine = {
                          parenDepth,
                          braceDepth,
                          inDoubleQuotes,
                          inSingleQuotes,
                          previousLineAllowsContinuation
                        };

                        let sanitizedLine = '';
                        let i = 0;
                        let lastSignificantChar = null;

                        while (i < line.length) {
                          const char = line[i];
                          const nextChar = line[i + 1];

                          if (inBlockComment) {
                            if (char === '*' && nextChar === '/') {
                              inBlockComment = false;
                              i += 2;
                            } else {
                              i++;
                            }
                            continue;
                          }

                          if (!inSingleQuotes && !inDoubleQuotes && char === '/' && nextChar === '*') {
                            inBlockComment = true;
                            i += 2;
                            continue;
                          }

                          if (!inSingleQuotes && !inDoubleQuotes && char === '/' && nextChar === '/') {
                            break;
                          }

                          sanitizedLine += char;

                          if (!inSingleQuotes && char === '"') {
                            if (nextChar === '"') {
                              sanitizedLine += nextChar;
                              i += 2;
                              continue;
                            }
                            inDoubleQuotes = !inDoubleQuotes;
                            i++;
                            continue;
                          }

                          if (!inDoubleQuotes && char === "'") {
                            if (nextChar === "'") {
                              sanitizedLine += nextChar;
                              i += 2;
                              continue;
                            }
                            inSingleQuotes = !inSingleQuotes;
                            i++;
                            continue;
                          }

                          if (!inDoubleQuotes && !inSingleQuotes) {
                            if (char === '{') {
                              braceDepth++;
                            } else if (char === '}' && braceDepth > 0) {
                              braceDepth--;
                            } else if (braceDepth === 0) {
                              if (char === '(') {
                                parenDepth++;
                              } else if (char === ')' && parenDepth > 0) {
                                parenDepth--;
                              }
                            }

                            if (!/\s/.test(char)) {
                              lastSignificantChar = char;
                            }
                          }

                          i++;
                        }

                        const trimmed = sanitizedLine.trim();
                        const hasContent = trimmed.length > 0;

                        if (!formulaStarted && hasContent && trimmed.startsWith('=')) {
                          formulaStarted = true;
                        }

                        if (
                          formulaStarted &&
                          hasContent &&
                          !stateBeforeLine.inDoubleQuotes &&
                          !stateBeforeLine.inSingleQuotes
                        ) {
                          const startsWithOperator = /^[+\-*/^&=)]/.test(trimmed);
                          const startsWithComma = trimmed.startsWith(',');
                          const allowsOperatorContinuation =
                            startsWithOperator ||
                            (startsWithComma && stateBeforeLine.parenDepth > 0) ||
                            (trimmed.startsWith(':') && stateBeforeLine.braceDepth > 0);

                          if (!stateBeforeLine.previousLineAllowsContinuation && !allowsOperatorContinuation) {
                            const leadingWhitespaceMatch = line.match(/^\s*/);
                            const leadingWhitespaceLength = leadingWhitespaceMatch ? leadingWhitespaceMatch[0].length : 0;
                            straySegments.push({
                              lineNumber: lineIndex + 1,
                              startColumn: leadingWhitespaceLength + 1,
                              endColumn: line.length + 1
                            });
                          }
                        }

                        if (hasContent) {
                          const endsWithContinuationChar =
                            lastSignificantChar !== null && LINE_CONTINUATION_CHARS.has(lastSignificantChar);
                          previousLineAllowsContinuation =
                            parenDepth > 0 ||
                            braceDepth > 0 ||
                            inDoubleQuotes ||
                            inSingleQuotes ||
                            endsWithContinuationChar;
                        }
                      }

                      return straySegments;
                    }

                    const externalDiagnosticsProvider =
                      typeof createExternalDiagnosticsProvider === 'function'
                        ? createExternalDiagnosticsProvider(monaco)
                        : null;

                    // Custom diagnostics provider for formula validation
                    const diagnosticsProvider = {
                      provideDiagnostics: function(model, lastResult) {
                        const diagnostics = [];
                        const text = model.getValue();

                        const mergeWithExternalMarkers = (markers = []) => {
                          if (
                            !externalDiagnosticsProvider ||
                            typeof externalDiagnosticsProvider.provideDiagnostics !== 'function'
                          ) {
                            return { markers };
                          }

                          const externalResult = externalDiagnosticsProvider.provideDiagnostics(model, lastResult);
                          const externalMarkers = Array.isArray(externalResult?.markers) ? externalResult.markers : [];

                          if (!externalMarkers.length) {
                            return { markers };
                          }

                          if (!Array.isArray(markers) || !markers.length) {
                            return { markers: externalMarkers };
                          }

                          return { markers: [...externalMarkers, ...markers] };
                        };

                        const lines = text.split('\n');

                        // Get formula from all lines and strip comments (both inline // and block /* */)
                        // Use the stripComments helper function for consistency
                        const fullText = lines.join('\n');
                        const formulaWithoutComments = stripComments(fullText) || '';

                        // Use formula without comments for validation
                        let formulaText = formulaWithoutComments;
                        if (formulaText) {
                          const trimmed = formulaText.trim();
                          if (!trimmed.startsWith('=')) {
                            formulaText = `=${trimmed}`;
                          } else {
                            formulaText = trimmed;
                          }
                        } else {
                          // Pure comments or whitespace – no diagnostics
                          return mergeWithExternalMarkers([]);
                        }

                        // Track recognized cell/range references so diagnostics can skip valid chips
                        const cellReferenceRanges = detectCellReferences(formulaText)
                          .filter((ref) => typeof ref.start === 'number' && typeof ref.end === 'number' && ref.end > ref.start)
                          .map((ref) => ({
                            start: ref.start,
                            end: ref.end
                          }));

                        // Check for common syntax errors
                        // 1. Check for unclosed string quotes
                        let inString = false;
                        let escaped = false;
                        let unclosedQuoteIndex = -1;
                        let unclosedQuoteLine = -1;
                        let unclosedQuoteColumn = -1;

                        for (let i = 0; i < formulaText.length; i++) {
                          const char = formulaText[i];

                          if (escaped) {
                            escaped = false;
                            continue;
                          }

                          if (char === '\\') {
                            escaped = true;
                            continue;
                          }

                          if (char === '"') {
                            if (!inString) {
                              // Opening quote - track its position (this will be overwritten if we close and open again)
                              unclosedQuoteIndex = i;
                              const beforeQuote = formulaText.substring(0, i);
                              const quoteLines = beforeQuote.split('\n');
                              unclosedQuoteLine = quoteLines.length - 1;
                              unclosedQuoteColumn = quoteLines[unclosedQuoteLine].length;
                            }
                            inString = !inString;
                          }
                        }

                        // If we're still in a string at the end, there's an unclosed quote
                        if (inString && unclosedQuoteIndex !== -1) {
                          // Find the end position (end of formula)
                          const endLines = formulaText.split('\n');
                          const endLineIndex = endLines.length - 1;
                          const endEditorLine = endLineIndex + 1;
                          const endColumn = endLines[endLineIndex].length + 1;

                          diagnostics.push({
                            severity: monaco.MarkerSeverity.Error,
                            startLineNumber: unclosedQuoteLine + 1, // +1 for 1-based line numbers
                            startColumn: unclosedQuoteColumn + 1, // +1 for 1-based columns
                            endLineNumber: endEditorLine,
                            endColumn: endColumn,
                            message: 'Unclosed string - missing closing quote (")',
                            source: 'Formula Validator',
                            code: 'UNCLOSED_STRING'
                          });
                        }

                        // 2. Check for disconnected expressions / stray text
                        const straySegments = detectStrayTopLevelText(fullText);
                        straySegments.forEach((segment) => {
                          diagnostics.push({
                            severity: monaco.MarkerSeverity.Error,
                            startLineNumber: segment.lineNumber,
                            startColumn: segment.startColumn,
                            endLineNumber: segment.lineNumber,
                            endColumn: segment.endColumn,
                            message: 'Stray text detected outside the main formula. Remove or convert it into a comment (//).',
                            source: 'Formula Validator',
                            code: 'STRAY_TEXT'
                          });
                        });

                        // 3. Check for periods in function arguments (should be commas)
                        // Need to check if we're in a function context and if parentheses are closed
                        const { openParens, closeParens } = analyzeParentheses(formulaText);
                        const hasUnclosedParens = openParens > closeParens;

                        // Pattern: value followed by period (in function context)
                        // Match periods that appear after values in function calls
                        // Note: Single quotes are reserved by Excel for sheet names
                        const periodPattern = /(\w+|"[^"]*"|[A-Z]+\$?\d+\$?)\s*\./g;
                        let match;
                        while ((match = periodPattern.exec(formulaText)) !== null) {
                          const tokenStartIndex = match.index;
                          const dotIndex = match.index + match[0].length - 1;

                          // Ignore dots that live inside completed quoted strings (e.g., "0.00")
                          if (isInsideQuotes(formulaText, tokenStartIndex) || isInsideQuotes(formulaText, dotIndex)) {
                            continue;
                          }

                          // Calculate line and column in the editor
                          const beforeMatch = formulaText.substring(0, match.index);
                          const formulaLines = beforeMatch.split('\n');
                          const formulaLineIndex = formulaLines.length - 1;
                          const editorLineNumber = formulaLineIndex + 1; // +1 for 1-based line numbers
                          const periodColumn = formulaLines[formulaLineIndex].length + match[1].length + 1; // Position of period

                          // Check if we're in a function call context
                          const textBeforePeriod = formulaText.substring(0, match.index + match[0].length);
                          const functionCallPattern = /\b([A-Z]+\w*)\s*\(/;
                          const isInFunction = functionCallPattern.test(textBeforePeriod);

                          if (isInFunction) {
                            if (hasUnclosedParens) {
                              // Function parentheses not closed - squiggle just the period
                              diagnostics.push({
                                severity: monaco.MarkerSeverity.Error,
                                startLineNumber: editorLineNumber,
                                startColumn: periodColumn,
                                endLineNumber: editorLineNumber,
                                endColumn: periodColumn + 1,
                                message: 'Expected comma (,) instead of period (.)',
                                source: 'Formula Validator',
                                code: 'PUNCTUATION_ERROR'
                              });
                            } else {
                              // Function parentheses are closed - squiggle the entire function
                              // Find the start of the function
                              const functionMatch = textBeforePeriod.match(/\b([A-Z]+\w*)\s*\(/);
                              if (functionMatch) {
                                const functionStart = functionMatch.index;
                                const functionStartLines = formulaText.substring(0, functionStart).split('\n');
                                const functionStartLineIndex = functionStartLines.length - 1;
                                const functionStartEditorLine = functionStartLineIndex + 2;
                                const functionStartColumn = functionStartLines[functionStartLineIndex].length + 1;

                                // Find the end of the function (last closing paren)
                                const lastCloseParen = formulaText.lastIndexOf(')');
                                const functionEndLines = formulaText.substring(0, lastCloseParen + 1).split('\n');
                                const functionEndLineIndex = functionEndLines.length - 1;
                                const functionEndEditorLine = functionEndLineIndex + 2;
                                const functionEndColumn = functionEndLines[functionEndLineIndex].length;

                                diagnostics.push({
                                  severity: monaco.MarkerSeverity.Error,
                                  startLineNumber: functionStartEditorLine,
                                  startColumn: functionStartColumn,
                                  endLineNumber: functionEndEditorLine,
                                  endColumn: functionEndColumn,
                                  message: 'Expected comma (,) instead of period (.)',
                                  source: 'Formula Validator',
                                  code: 'PUNCTUATION_ERROR'
                                });
                              }
                            }
                          }
                        }

                        // 4. Validate function calls for empty parameters and argument counts
                        if (window.hyperFormulaFunctions && window.hyperFormulaFunctions.length > 0) {
                          const functionCallPattern = /\b([A-Za-z_][A-Za-z0-9_]*)\s*\(/g;
                          let functionMatch;

                          while ((functionMatch = functionCallPattern.exec(formulaText)) !== null) {
                            const funcName = functionMatch[1];
                            const metadata = getFunctionSignatureInfo(funcName);
                            if (!metadata) {
                              continue;
                            }

                            const relativeParenIndex = functionMatch[0].lastIndexOf('(');
                            if (relativeParenIndex === -1) {
                              continue;
                            }

                            const openParenIndex = functionMatch.index + relativeParenIndex;
                            const closeParenIndex = findClosingParenthesis(formulaText, openParenIndex);
                            if (closeParenIndex === -1) {
                              continue;
                            }

                            const paramString = formulaText.substring(openParenIndex + 1, closeParenIndex);
                            const { emptySegments, actualCount } = parseParameterSegments(paramString);

                            emptySegments.forEach((segment) => {
                              const absoluteStart = openParenIndex + 1 + segment.start;
                              const absoluteEnd = openParenIndex + 1 + segment.end;

                              let highlightStart = absoluteStart;
                              let highlightEnd = absoluteEnd;

                              if (highlightStart === highlightEnd) {
                                highlightEnd = Math.min(highlightStart + 1, closeParenIndex + 1);
                              }

                              const startPos = indexToEditorPosition(formulaText, highlightStart);
                              const endPos = indexToEditorPosition(formulaText, highlightEnd);

                              diagnostics.push({
                                severity: monaco.MarkerSeverity.Warning,
                                startLineNumber: startPos.lineNumber,
                                startColumn: startPos.column,
                                endLineNumber: endPos.lineNumber,
                                endColumn: endPos.column,
                                message: 'Empty parameter - may cause unexpected behavior',
                                source: 'Formula Validator',
                                code: 'EMPTY_PARAMETER'
                              });
                            });

                            let argErrorMessage = null;
                            if (actualCount < metadata.minArgs) {
                              argErrorMessage = `Expected at least ${metadata.minArgs} argument(s), got ${actualCount}`;
                            } else if (metadata.maxArgs !== Infinity && actualCount > metadata.maxArgs) {
                              argErrorMessage = `Expected at most ${metadata.maxArgs} argument(s), got ${actualCount}`;
                            }

                            if (argErrorMessage) {
                              const callStartPos = indexToEditorPosition(formulaText, openParenIndex + 1);
                              const callEndPos = indexToEditorPosition(formulaText, Math.max(openParenIndex + 1, closeParenIndex));

                              diagnostics.push({
                                severity: monaco.MarkerSeverity.Error,
                                startLineNumber: callStartPos.lineNumber,
                                startColumn: callStartPos.column,
                                endLineNumber: callEndPos.lineNumber,
                                endColumn: callEndPos.column,
                                message: `${metadata.displayName}: ${argErrorMessage}. Signature: ${metadata.signature}`,
                                source: 'Formula Validator',
                                code: 'FUNCTION_SIGNATURE_MISMATCH'
                              });
                            }
                          }
                        }

                        // 5. Check for bare identifiers (strings without quotes) that are not valid
                        // Find all bare identifiers (words not in quotes, not numbers, not operators, not delimiters)
                        const identifierPattern = /\b[A-Za-z_][A-Za-z0-9_]*\b/g;
                        const processedPositions = new Set(); // Track processed positions to avoid duplicates

                        while ((match = identifierPattern.exec(formulaText)) !== null) {
                          const identifier = match[0];
                          const startPos = match.index;
                          const endPos = startPos + identifier.length;

                          // Skip identifiers that fall entirely inside a recognized cell/range reference (chip)
                          const insideCellReference = cellReferenceRanges.some((range) => {
                            return startPos >= range.start && endPos <= range.end;
                          });
                          if (insideCellReference) {
                            continue;
                          }

                          // Skip if we've already processed this position
                          if (processedPositions.has(startPos)) continue;
                          processedPositions.add(startPos);

                          // Ignore identifiers that live inside completed quoted strings
                          if (isInsideQuotes(formulaText, startPos)) {
                            continue;
                          }

                          // Check if it's followed by a parenthesis - treat as function call even if list not loaded yet
                          const trailingText = formulaText.substring(endPos);
                          const leadingWhitespaceMatch = trailingText.match(/^\s*/);
                          const nextNonWhitespaceChar = leadingWhitespaceMatch ? trailingText[leadingWhitespaceMatch[0].length] : trailingText[0];
                          if (nextNonWhitespaceChar === '(') continue;

                          // Check if it's a function (case-insensitive)
                          const isFunction = window.hyperFormulaFunctions && 
                                             window.hyperFormulaFunctions.some(f => 
                                               f.name.toUpperCase() === identifier.toUpperCase()
                                             );
                          if (isFunction) continue;

                          // Check if it's a logical value
                          const isLogical = ['TRUE', 'FALSE'].includes(identifier.toUpperCase());
                          if (isLogical) continue;

                          // Check if it's a cell reference (pattern: 1-3 letters, optional $, digits, optional $)
                          const isCellReference = /^\$?[A-Z]{1,3}\$?\d+\$?$/i.test(identifier);
                          if (isCellReference) continue;

                          // Check if it's a named range
                          const isNamedRange = window.namedRanges && window.namedRanges.has(identifier.toUpperCase());
                          if (isNamedRange) continue;

                          // If we get here, it's an invalid bare identifier - mark it with red squiggle
                          const beforeMatch = formulaText.substring(0, startPos);
                          const formulaLines = beforeMatch.split('\n');
                          const formulaLineIndex = formulaLines.length - 1;
                          const editorLineNumber = formulaLineIndex + 1; // +1 for 1-based line numbers
                          const startColumn = formulaLines[formulaLineIndex].length + 1; // +1 for 1-based
                          const endColumn = startColumn + identifier.length;

                          diagnostics.push({
                            severity: monaco.MarkerSeverity.Error,
                            startLineNumber: editorLineNumber,
                            startColumn: startColumn,
                            endLineNumber: editorLineNumber,
                            endColumn: endColumn,
                            message: `Invalid identifier: "${identifier}". Use double quotes for strings or define as a named range.`,
                            source: 'Formula Validator',
                            code: 'INVALID_IDENTIFIER'
                          });
                        }

                        // 6. Validate formula syntax with HyperFormula if available
                        // Only validate if the formula looks complete (has closing parens, etc.)
                        if (window.hf && formulaText) {
                          try {
                            // Ensure formula starts with '='
                            const testFormula = formulaText.startsWith('=') ? formulaText : '=' + formulaText;

                            // Check for common syntax errors first (parentheses, periods, etc.)
                            const { openParens, closeParens, firstExtraClosingIndex } = analyzeParentheses(formulaText);

                            // Only validate with HyperFormula if the formula looks syntactically complete
                            // Don't validate incomplete formulas (missing closing parens, etc.)
                            // This prevents false positives for formulas that are still being typed
                            const isLikelyComplete = openParens === closeParens && formulaText.trim().length > 0;

                            if (isLikelyComplete) {
                              // Try to validate the formula by attempting to set it in a test cell
                              // Use a very high cell address that won't interfere with actual data
                              const testRow = 99999;
                              const testCol = 99999;
                              const testSheetId = 0;

                              try {
                                // Try to set the formula in a test location
                                // This will throw an error if the formula syntax is invalid
                                window.hf.setCellContents({ col: testCol, row: testRow, sheet: testSheetId }, [[testFormula]]);

                                // If successful, clear the test cell immediately
                                window.hf.setCellContents({ col: testCol, row: testRow, sheet: testSheetId }, [['']]);

                                // Formula is valid - no error markers needed

                              } catch (formulaError) {
                                // Formula might be invalid, but check error type
                                // Some errors are runtime errors (like #REF!, #NAME?, etc.) not syntax errors
                                const errorMessage = (formulaError.message || '').toLowerCase();
                                const errorString = String(formulaError).toLowerCase();
                                const fullError = errorMessage + ' ' + errorString;

                                // Only mark as syntax error if it's actually a syntax/parse error
                                // Runtime errors like #REF!, #NAME?, #VALUE!, etc. are not syntax errors
                                // Also ignore errors about unknown addresses (cells that don't exist yet)
                                // HyperFormula often throws errors for valid formulas if referenced cells don't exist
                                const isSyntaxError = (fullError.includes('parse') || 
                                                     fullError.includes('syntax') ||
                                                     fullError.includes('unexpected token') ||
                                                     fullError.includes('invalid character') ||
                                                     fullError.includes('cannot parse') ||
                                                     fullError.includes('parsing error')) &&
                                                     !fullError.includes('unknown') &&
                                                     !fullError.includes('address') &&
                                                     !fullError.includes('cell') &&
                                                     !fullError.includes('reference') &&
                                                     !fullError.includes('ref') &&
                                                     !fullError.includes('name') &&
                                                     !fullError.includes('value') &&
                                                     !fullError.includes('div') &&
                                                     !fullError.includes('num') &&
                                                     !fullError.includes('na') &&
                                                     !fullError.includes('error') &&
                                                     !fullError.includes('#');

                                // Don't mark errors for valid formulas - most HyperFormula errors are runtime, not syntax
                                // Only mark if it's clearly a parse/syntax error
                                if (isSyntaxError) {
                                  // Mark as syntax error
                                  const formulaStart = formulaText.startsWith('=') ? 1 : 0;
                                  diagnostics.push({
                                    severity: monaco.MarkerSeverity.Error,
                                    startLineNumber: 2, // Line 2 is where formula starts
                                    startColumn: formulaStart + 1,
                                    endLineNumber: 2,
                                    endColumn: Math.min(formulaText.length + 1, 200),
                                    message: 'Invalid formula syntax',
                                    source: 'HyperFormula',
                                    code: 'SYNTAX_ERROR'
                                  });
                                }
                                // If it's a runtime error (like #REF!, #NAME?), don't mark it - those are valid formulas
                                // If it's an unknown address error, don't mark it - the cell might not exist yet
                              }
                            } else {
                              // Formula is incomplete - only check for obvious syntax errors
                              if (openParens > closeParens) {
                                // Missing closing parenthesis - but only mark if formula looks complete otherwise
                                // Don't mark if user is still typing
                                const lastOpenParen = formulaText.lastIndexOf('(');
                                if (lastOpenParen > 0) {
                                  // Check if there's content after the last open paren
                                  const afterLastParen = formulaText.substring(lastOpenParen + 1);
                                  // Only mark as error if there's significant content after the last paren
                                  // This prevents marking incomplete formulas as errors
                                  if (afterLastParen.trim().length > 0 && !afterLastParen.match(/^\s*$/)) {
                                    const formulaLines = formulaText.substring(0, lastOpenParen + 1).split('\n');
                                    const lineNum = formulaLines.length;
                                    const colNum = formulaLines[formulaLines.length - 1].length;

                                    diagnostics.push({
                                      severity: monaco.MarkerSeverity.Warning, // Use warning instead of error for incomplete formulas
                                      startLineNumber: lineNum + 1, // Line numbers are 1-based
                                      startColumn: colNum + 1,
                                      endLineNumber: lineNum + 1,
                                      endColumn: colNum + 2,
                                      message: 'Missing closing parenthesis',
                                      source: 'Formula Validator',
                                      code: 'MISSING_PARENTHESIS'
                                    });
                                  }
                                }
                              } else if (closeParens > openParens) {
                                // Extra closing parenthesis - always mark this as an error
                                const extraIndex = typeof firstExtraClosingIndex === 'number' ? firstExtraClosingIndex : -1;

                                if (extraIndex !== -1) {
                                  const formulaLines = formulaText.substring(0, extraIndex + 1).split('\n');
                                  const lineNum = formulaLines.length;
                                  const colNum = formulaLines[formulaLines.length - 1].length;

                                  diagnostics.push({
                                    severity: monaco.MarkerSeverity.Error,
                                    startLineNumber: lineNum + 1,
                                    startColumn: colNum,
                                    endLineNumber: lineNum + 1,
                                    endColumn: colNum + 1,
                                    message: 'Extra closing parenthesis',
                                    source: 'Formula Validator',
                                    code: 'EXTRA_PARENTHESIS'
                                  });
                                }
                              }
                            }
                          } catch (e) {
                            // If validation completely fails, don't mark anything
                            // This prevents false positives
                          }
                        }

                        return mergeWithExternalMarkers(diagnostics);
                      }
                    };

                    // Function to update messages pane with error/warning counts and details
                    function updateMessagesPane(markers) {
                      if (!markers || markers.length === 0) {
                        markers = [];
                      }

                      // Count errors and warnings
                      const errorCount = markers.filter(m => m.severity === monaco.MarkerSeverity.Error).length;
                      const warningCount = markers.filter(m => m.severity === monaco.MarkerSeverity.Warning).length;

                      // Get indicator elements
                      const errorIcon = document.querySelector('.error-icon');
                      const errorCountSpan = document.querySelector('.error-count');
                      const warningIcon = document.querySelector('.warning-icon');
                      const warningCountSpan = document.querySelector('.warning-count');
                      const messagesContent = document.getElementById('messagesContent');

                      // Update error indicator
                      if (errorIcon && errorCountSpan) {
                        errorCountSpan.textContent = errorCount;
                        if (errorCount > 0) {
                          errorIcon.classList.add('has-errors');
                          errorCountSpan.classList.add('has-errors');
                        } else {
                          errorIcon.classList.remove('has-errors');
                          errorCountSpan.classList.remove('has-errors');
                        }
                      }

                      // Update warning indicator
                      if (warningIcon && warningCountSpan) {
                        warningCountSpan.textContent = warningCount;
                        if (warningCount > 0) {
                          warningIcon.classList.add('has-warnings');
                          warningCountSpan.classList.add('has-warnings');
                        } else {
                          warningIcon.classList.remove('has-warnings');
                          warningCountSpan.classList.remove('has-warnings');
                        }
                      }

                      // Populate messages content
                      if (messagesContent) {
                        if (markers.length === 0) {
                          messagesContent.innerHTML = '<div class="message-item message-success"><span class="message-icon" style="color: #4ec9b0;">✓</span><span class="message-content">No errors or warnings</span></div>';
                        } else {
                          // Sort by severity (errors first) and then by line number
                          const sortedMarkers = [...markers].sort((a, b) => {
                            if (a.severity !== b.severity) {
                              return a.severity - b.severity; // Error (8) comes before Warning (4)
                            }
                            return a.startLineNumber - b.startLineNumber;
                          });

                          const messagesHTML = sortedMarkers.map((marker, index) => {
                            const severity = marker.severity === monaco.MarkerSeverity.Error ? 'error' : 'warning';
                            const icon = marker.severity === monaco.MarkerSeverity.Error ? '✕\uFE0E' : '⚠\uFE0E';
                            const source = marker.source || 'Unknown';
                            const code = marker.code || '';
                            const messageText = marker.message || 'No message';

                            // Store marker data in data attributes for navigation
                            return `
                              <div class="message-item message-${severity}" 
                                   data-line="${marker.startLineNumber}" 
                                   data-column="${marker.startColumn}"
                                   data-end-line="${marker.endLineNumber || marker.startLineNumber}"
                                   data-end-column="${marker.endColumn || marker.startColumn}"
                                   data-marker-index="${index}">
                                <span class="message-icon">${icon}</span>
                                <span class="message-text">
                                  <span class="message-location">Line ${marker.startLineNumber}:${marker.startColumn}</span>
                                  <span class="message-source">[${source}]</span>
                                  <span class="message-content">${escapeHtml(messageText)}</span>
                                  ${code ? `<span class="message-code">(${code})</span>` : ''}
                                </span>
                              </div>
                            `;
                          }).join('');

                          messagesContent.innerHTML = messagesHTML;

                          // Add click handlers to navigate to error/warning locations
                          const messageItems = messagesContent.querySelectorAll('.message-item[data-line]');
                          messageItems.forEach(item => {
                            item.addEventListener('click', () => {
                              const line = parseInt(item.getAttribute('data-line'));
                              const column = parseInt(item.getAttribute('data-column'));
                              const endLine = parseInt(item.getAttribute('data-end-line'));
                              const endColumn = parseInt(item.getAttribute('data-end-column'));

                              // Get the editor instance (use window.monacoEditor or the local editor variable)
                              const currentEditor = window.monacoEditor || editor;
                              const currentMonaco = window.monaco || monaco;

                              // Navigate to the error/warning location
                              if (currentEditor && currentMonaco && typeof currentEditor.setPosition === 'function') {
                                // Set cursor position
                                currentEditor.setPosition({ lineNumber: line, column: column });

                                // Select the error/warning range
                                const range = new currentMonaco.Range(line, column, endLine, endColumn);
                                currentEditor.setSelection(range);

                                // Reveal the line in the center of the viewport
                                currentEditor.revealLineInCenter(line);

                                // Focus the editor
                                currentEditor.focus();
                              }
                            });
                          });
                        }
                      }
                    }

                    const RESULT_PREVIEW_DELAY_MS = 900;
                    const RESULT_PREVIEW_CELL = { row: 800, col: 800 };
                    const CHIP_INSTANTIATION_SETTLE_MS = 150;
                    const resultPaneElement = document.querySelector('.result-pane');
                    const resultHeaderValueElement = document.getElementById('resultHeaderValue');
                    let resultPreviewTimeout = null;
                    let lastResultRenderedText = '';
                    let lastResultRenderStatus = 'idle';
                    let lastResultEvaluatedFormula = '';
                    let formulaResultDirty = false;
                    let chipInstantiationPending = false;
                    let chipInstantiationSettleTimeout = null;
                    let pendingResultPreviewMarkers = null;

                    function getResultPaneElement() {
                      return document.getElementById('resultContent');
                    }

                    function isHashErrorText(text) {
                      if (!text) return false;
                      return text.trim().startsWith('#');
                    }

                    function updateResultPaneDisplay(displayValue, status = 'idle') {
                      const resultContainer = getResultPaneElement();
                      if (!resultContainer) {
                        return;
                      }

                      const normalizedValue = displayValue === null || displayValue === undefined
                        ? ''
                        : String(displayValue);

                      if (normalizedValue === lastResultRenderedText && status === lastResultRenderStatus) {
                        return;
                      }

                      const hashIndicatesError = isHashErrorText(normalizedValue);
                      lastResultRenderedText = normalizedValue;
                      lastResultRenderStatus = hashIndicatesError ? 'error' : status;

                      if (!normalizedValue) {
                        resultContainer.innerHTML = '<div class="result-placeholder">Result preview updates after changes without errors.</div>';
                        updateResultHeaderValue();
                        return;
                      }

                      const modifierClass = lastResultRenderStatus === 'error' ? 'result-output-error' : 'result-output-success';
                      resultContainer.innerHTML = `
                        <div class="result-output ${modifierClass}">
                          <span class="result-label">Result</span>
                          <span class="result-value">${escapeHtml(normalizedValue)}</span>
                        </div>
                      `;
                      updateResultHeaderValue();
                    }

                    function updateResultHeaderValue() {
                      if (!resultHeaderValueElement) {
                        return;
                      }
                      if (lastResultRenderedText) {
                        resultHeaderValueElement.textContent = lastResultRenderedText;
                        resultHeaderValueElement.setAttribute('data-status', lastResultRenderStatus || 'idle');
                      } else {
                        resultHeaderValueElement.textContent = '';
                        resultHeaderValueElement.removeAttribute('data-status');
                      }
                    }
                    window.updateResultHeaderValue = updateResultHeaderValue;

                    function getDisplayTextForCell(cell) {
                      if (!cell) {
                        return '';
                      }
                      const explicitDisplay = cell.getAttribute('data-display-text');
                      if (explicitDisplay !== null && explicitDisplay !== undefined) {
                        return explicitDisplay;
                      }
                      const textContent = cell.textContent;
                      return textContent === null || textContent === undefined ? '' : textContent.trim();
                    }

                    function syncResultPaneWithSelection() {
                      if (currentMode !== MODES.READY) {
                        return;
                      }
                      const activeCell = window.selectedCell || null;
                      if (!activeCell) {
                        updateResultPaneDisplay('', 'idle');
                        return;
                      }
                      const displayText = getDisplayTextForCell(activeCell);
                      if (!displayText) {
                        updateResultPaneDisplay('', 'idle');
                        return;
                      }
                      updateResultPaneDisplay(displayText, 'success');
                    }
                    window.syncResultPaneWithSelection = syncResultPaneWithSelection;

                    function cancelScheduledResultPreview() {
                      if (resultPreviewTimeout) {
                        clearTimeout(resultPreviewTimeout);
                        resultPreviewTimeout = null;
                      }
                    }

                    function markChipInstantiationStarted() {
                      chipInstantiationPending = true;
                      pendingResultPreviewMarkers = null;
                      if (chipInstantiationSettleTimeout) {
                        clearTimeout(chipInstantiationSettleTimeout);
                        chipInstantiationSettleTimeout = null;
                      }
                      cancelScheduledResultPreview();
                    }

                    function scheduleChipInstantiationSettled() {
                      if (!chipInstantiationPending) {
                        pendingResultPreviewMarkers = null;
                        return;
                      }
                      if (chipInstantiationSettleTimeout) {
                        clearTimeout(chipInstantiationSettleTimeout);
                      }
                      chipInstantiationSettleTimeout = setTimeout(() => {
                        chipInstantiationPending = false;
                        chipInstantiationSettleTimeout = null;
                        const queuedMarkers = pendingResultPreviewMarkers;
                        pendingResultPreviewMarkers = null;
                        if (queuedMarkers) {
                          scheduleResultPreview(queuedMarkers);
                        }
                      }, CHIP_INSTANTIATION_SETTLE_MS);
                    }

                    function getFormulaTextForResultPreview() {
                      const currentEditor = window.monacoEditor || editor;
                      if (!currentEditor || typeof currentEditor.getValue !== 'function') {
                        return '';
                      }
                      const normalized = stripComments(currentEditor.getValue() || '').trim();
                      if (!normalized) {
                        return '';
                      }
                      return normalized.startsWith('=') ? normalized : `=${normalized}`;
                    }

                    function interpretPreviewValue(rawValue) {
                      if (rawValue === null || typeof rawValue === 'undefined') {
                        return { status: 'success', value: '' };
                      }

                      if (Array.isArray(rawValue)) {
                        const flattened = rawValue.map(item => {
                          if (Array.isArray(item)) {
                            return item.map(sub => (sub === null || sub === undefined) ? '' : String(sub)).join(', ');
                          }
                          return item === null || item === undefined ? '' : String(item);
                        }).join(' ; ');
                        return { status: 'success', value: flattened };
                      }

                      if (typeof rawValue === 'object') {
                        const typeString = rawValue.type ? String(rawValue.type).toLowerCase() : '';
                        const valuePortion = rawValue.value !== undefined
                          ? rawValue.value
                          : (rawValue.message !== undefined ? rawValue.message : '');
                        const valueText = valuePortion === null || valuePortion === undefined
                          ? ''
                          : String(valuePortion);
                        const isError = typeString.includes('error');
                        if (isError) {
                          return { status: 'error', value: valueText || '#ERROR!' };
                        }
                        if (valueText) {
                          return { status: 'success', value: valueText };
                        }
                        return { status: 'success', value: JSON.stringify(rawValue) };
                      }

                      return { status: 'success', value: String(rawValue) };
                    }

                    function evaluateFormulaForPreview(formulaText) {
                      if (!window.hf || !formulaText) {
                        return { status: 'error', value: '#ERROR!' };
                      }
                      const sheetId = typeof window.hfSheetId === 'number' ? window.hfSheetId : 0;
                      const previewAddress = {
                        sheet: sheetId,
                        row: RESULT_PREVIEW_CELL.row,
                        col: RESULT_PREVIEW_CELL.col
                      };
                      try {
                        window.hf.setCellContents(previewAddress, [[formulaText]]);
                        const rawValue = window.hf.getCellValue(previewAddress);
                        return interpretPreviewValue(rawValue);
                      } catch (error) {
                        console.error('Result preview evaluation failed:', error);
                        return { status: 'error', value: '#ERROR!' };
                      }
                    }

                    function runResultPreview(formulaText) {
                      resultPreviewTimeout = null;
                      const evaluation = evaluateFormulaForPreview(formulaText);
                      lastResultEvaluatedFormula = formulaText;
                      formulaResultDirty = false;
                      updateResultPaneDisplay(evaluation.value, evaluation.status);
                    }

                    function scheduleResultPreview(markers) {
                      if (currentMode === MODES.READY) {
                        cancelScheduledResultPreview();
                        lastResultEvaluatedFormula = '';
                        formulaResultDirty = false;
                        if (typeof syncResultPaneWithSelection === 'function') {
                          syncResultPaneWithSelection();
                        } else {
                          updateResultPaneDisplay('', 'idle');
                        }
                        return;
                      }
                      if (chipInstantiationPending) {
                        pendingResultPreviewMarkers = Array.isArray(markers) ? markers : (markers || []);
                        return;
                      }
                      const hasErrors = Array.isArray(markers) && markers.some(marker => marker.severity === monaco.MarkerSeverity.Error);
                      if (hasErrors) {
                        cancelScheduledResultPreview();
                        updateResultPaneDisplay('#ERROR!', 'error');
                        return;
                      }

                      const formulaText = getFormulaTextForResultPreview();
                      if (!formulaText) {
                        cancelScheduledResultPreview();
                        lastResultEvaluatedFormula = '';
                        updateResultPaneDisplay('', 'idle');
                        formulaResultDirty = false;
                        return;
                      }

                      if (!formulaResultDirty && formulaText === lastResultEvaluatedFormula) {
                        return;
                      }

                      cancelScheduledResultPreview();
                      resultPreviewTimeout = setTimeout(() => {
                        runResultPreview(formulaText);
                      }, RESULT_PREVIEW_DELAY_MS);
                    }

                    // Use Monaco's markers API to set diagnostics
                    function validateFormula() {
                      const model = editor.getModel();
                      if (!model) return;

                      const diagnostics = diagnosticsProvider.provideDiagnostics(model, null);
                      const markers = Array.isArray(diagnostics?.markers) ? diagnostics.markers : [];

                      // Apply Monaco markers so inline highlights/squiggles appear
                      monaco.editor.setModelMarkers(model, 'formula-validator', markers);

                      // Update messages pane with diagnostics
                      updateMessagesPane(markers);
                      scheduleResultPreview(markers);
                    }

                    // Validate on content change (with debounce for performance)
                    let validationTimeout;
                    editor.onDidChangeModelContent(() => {
                      formulaResultDirty = true;
                      clearTimeout(validationTimeout);
                      validationTimeout = setTimeout(() => {
                        validateFormula();
                      }, 300); // Debounce validation to avoid performance issues
                    });

                    // Initial validation to set up messages pane
                    validateFormula();

                    function getTopLeftCellFromNormalizedReference(referenceText = '') {
                      if (!referenceText) {
                        return '';
                      }
                      if (!referenceText.includes(':')) {
                        return referenceText;
                      }
                      const parts = referenceText.split(':').filter(Boolean);
                      if (parts.length === 0) {
                        return '';
                      }
                      if (parts.length === 1) {
                        return parts[0];
                      }
                      const startAddress = cellRefToAddress(parts[0]);
                      const endAddress = cellRefToAddress(parts[1]);
                      if (!startAddress && !endAddress) {
                        return '';
                      }
                      if (!startAddress) {
                        return parts[1];
                      }
                      if (!endAddress) {
                        return parts[0];
                      }
                      const minRow = Math.min(startAddress[0], endAddress[0]);
                      const minCol = Math.min(startAddress[1], endAddress[1]);
                      return addressToCellRef(minRow, minCol);
                    }

                    function getCellPreviewForDependencies(cellRef) {
                      if (!cellRef || !window.hf) {
                        return { text: '', type: 'empty' };
                      }
                      const address = cellRefToAddress(cellRef);
                      if (!address) {
                        return { text: '', type: 'empty' };
                      }
                      const [row, col] = address;
                      const sheetId = getActiveSheetId();
                      let formulaText = '';
                      try {
                        formulaText = window.hf.getCellFormula({ col, row, sheet: sheetId });
                      } catch (error) {
                        formulaText = '';
                      }
                      if (typeof formulaText === 'string' && formulaText.trim()) {
                        const trimmedFormula = formulaText.trim();
                        return {
                          text: trimmedFormula.startsWith('=') ? trimmedFormula : `=${trimmedFormula}`,
                          type: 'formula'
                        };
                      }
                      let rawValue = null;
                      try {
                        rawValue = window.hf.getCellValue({ col, row, sheet: sheetId });
                      } catch (error) {
                        rawValue = null;
                      }
                      if (rawValue === null || typeof rawValue === 'undefined') {
                        return { text: '', type: 'empty' };
                      }
                      const interpreted = typeof interpretPreviewValue === 'function'
                        ? interpretPreviewValue(rawValue)
                        : null;
                      const interpretedValue = interpreted && Object.prototype.hasOwnProperty.call(interpreted, 'value')
                        ? interpreted.value
                        : rawValue;
                      const finalText = interpretedValue === null || typeof interpretedValue === 'undefined'
                        ? ''
                        : String(interpretedValue);
                      return {
                        text: finalText,
                        type: finalText ? 'value' : 'empty'
                      };
                    }

                    function renderDependenciesPaneFromReferences(referenceMatches = [], context = {}) {
                      if (!dependenciesContentElement) {
                        return;
                      }

                      const formulaText = typeof context.formulaText === 'string' ? context.formulaText : '';
                      const normalizedFormula = stripComments(formulaText).trim();
                      if (!normalizedFormula) {
                        updateDependenciesCountBadge(0);
                        dependenciesContentElement.innerHTML = '<div class="dependencies-empty">Enter a formula to view dependencies.</div>';
                        return;
                      }

                      const uniqueRefs = [];
                      const seenRefs = new Set();

                      referenceMatches.forEach((match) => {
                        const displayText = (match?.fullMatch || match?.range || match?.text || '').trim();
                        if (!displayText) {
                          return;
                        }
                        const normalized = normalizeCellReference(match?.text || match?.range || match?.fullMatch || displayText);
                        if (!normalized || seenRefs.has(normalized)) {
                          return;
                        }
                        seenRefs.add(normalized);
                        uniqueRefs.push({
                          displayText,
                          normalized,
                          colorKey: match?.fullMatch || match?.range || match?.text || displayText
                        });
                      });

                      if (uniqueRefs.length === 0) {
                        updateDependenciesCountBadge(0);
                        dependenciesContentElement.innerHTML = '<div class="dependencies-empty">No cell references detected.</div>';
                        return;
                      }

                      const listItemsHtml = uniqueRefs.map((refInfo) => {
                        const firstCellRef = getTopLeftCellFromNormalizedReference(refInfo.normalized);
                        const chipClassName = getChipColorClass(refInfo.colorKey);
                        const valueInfo = getCellPreviewForDependencies(firstCellRef);
                        const hasValueText = typeof valueInfo.text === 'string' && valueInfo.text.trim().length > 0;
                        const valueMarkup = hasValueText
                          ? escapeHtml(valueInfo.text)
                          : '<span class="dependency-empty-value">Empty</span>';
                        const valueClass = valueInfo.type === 'formula' ? 'is-formula' : '';
                        const safeReferenceLabel = escapeHtml(refInfo.displayText);
                        const cellMarkup = firstCellRef
                          ? escapeHtml(firstCellRef)
                          : '<span class="dependency-empty-value">Unknown</span>';
                        const chipClassAttribute = chipClassName ? ` ${chipClassName}` : '';
                        return `
                          <div class="dependency-item${chipClassAttribute}">
                            <span class="dependency-chip" aria-hidden="true"></span>
                            <div class="dependency-body">
                              <div class="dependency-reference">${safeReferenceLabel}</div>
                              <div class="dependency-first-cell">
                                <span class="dependency-first-address">${cellMarkup}</span>
                                <span class="dependency-value ${valueClass}">${valueMarkup}</span>
                              </div>
                            </div>
                          </div>
                        `;
                      }).join('');

                      updateDependenciesCountBadge(uniqueRefs.length);
                      dependenciesContentElement.innerHTML = `<div class="dependencies-list">${listItemsHtml}</div>`;
                    }

                    // Cell chip tracking for inline cell references
                    let cellChipDecorations = [];
                    let cellChipRanges = [];
                    let chipUpdateTimeout;
                    let isCursorInChip = false;
                    let editorHasFocus = true;
                    const defaultQuickSuggestions = editor.getOption(monaco.editor.EditorOption.quickSuggestions);
                    let chipSuppressedQuickSuggestions = false;

                    function suppressSuggestionsInChip(shouldSuppress) {
                      if (shouldSuppress) {
                        if (!chipSuppressedQuickSuggestions) {
                          chipSuppressedQuickSuggestions = true;
                          editor.updateOptions({ quickSuggestions: false });
                        }
                        const suggestController = editor.getContribution('editor.contrib.suggestController');
                        if (suggestController && typeof suggestController.cancel === 'function') {
                          suggestController.cancel();
                        }
                      } else if (chipSuppressedQuickSuggestions) {
                        chipSuppressedQuickSuggestions = false;
                        editor.updateOptions({ quickSuggestions: defaultQuickSuggestions });
                      }
                    }

                    function isPositionInsideChip(position) {
                      if (!position || cellChipRanges.length === 0) return false;
                      return cellChipRanges.some(range => range.containsPosition(position));
                    }

                    function hasChipPaddingAtLineStart(lineNumber) {
                      if (!lineNumber) return false;
                      return cellChipRanges.some(range => 
                        range.startLineNumber === lineNumber &&
                        range.startColumn === 1
                      );
                    }

                    // Update cell chips based on editor content using decorations
                    // Track last processed text to avoid re-processing after space insertion
                    let lastProcessedText = '';

                    function isIncompleteRangeReference(refText = '') {
                      if (!refText) return false;
                      return refText.trim().endsWith(':');
                    }

                    function updateCellChips() {
                      const model = editor.getModel();
                      if (!model) return;

                      // Get all text from editor
                      const fullText = model.getValue();

                      // Detect cell references
                      const detectedRefs = detectCellReferences(fullText);
                      const processedRefs = editorHasFocus ? detectedRefs : detectedRefs.filter(ref => !isIncompleteRangeReference(ref.text));
                      if (detectedRefs.length === 0 && cellChipColorAssignments.size > 0) {
                        resetChipColorAssignments();
                      }
                      const cellRefs = processedRefs;
                      renderDependenciesPaneFromReferences(cellRefs, { formulaText: fullText });
                      const newHighlightAssignments = new Map();

                      // Ensure there is always at least one literal space between consecutive chips
                      const sortedRefs = [...cellRefs].sort((a, b) => a.start - b.start);
                      let insertedGap = false;
                      if (sortedRefs.length > 1) {
                        for (let i = 0; i < sortedRefs.length - 1; i++) {
                          const currentRef = sortedRefs[i];
                          const nextRef = sortedRefs[i + 1];
                          if (nextRef.start <= currentRef.end) {
                            continue; // Overlapping or nested references - handled elsewhere
                          }
                          if (nextRef.start === currentRef.end) {
                            const insertPos = model.getPositionAt(nextRef.start);
                            const gapEdits = [{
                              range: new monaco.Range(insertPos.lineNumber, insertPos.column, insertPos.lineNumber, insertPos.column),
                              text: ' '
                            }];
                            window.isProgrammaticCursorChange = true;
                            editor.executeEdits('insert-chip-gap', gapEdits);

                            const cursorPos = editor.getPosition();
                            if (cursorPos) {
                              const isAfterGap = cursorPos.lineNumber > insertPos.lineNumber ||
                                (cursorPos.lineNumber === insertPos.lineNumber && cursorPos.column >= insertPos.column);
                              if (isAfterGap) {
                                editor.setPosition(new monaco.Position(insertPos.lineNumber, insertPos.column + 1));
                              }
                            }

                            Promise.resolve().then(() => {
                              window.isProgrammaticCursorChange = false;
                              setTimeout(() => {
                                updateCellChips();
                              }, 10);
                            });
                            insertedGap = true;
                            break;
                          }
                        }
                        if (insertedGap) {
                          return;
                        }
                      }

                      // Convert character indices to line/column positions and create decorations
                      const decorations = [];
                      const newChipRanges = [];
                      const lines = fullText.split('\n');
                      const edits = [];
                      let shouldMoveCursor = false;
                      let cursorMovePosition = null;

                      // Build line start positions map
                      const lineStarts = [];
                      let charCount = 0;
                      for (let i = 0; i < lines.length; i++) {
                        lineStarts.push(charCount);
                        charCount += lines[i].length + 1; // +1 for newline
                      }

                      cellRefs.forEach(ref => {
                        // Find which line this cell reference is on
                        for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
                          const lineStart = lineStarts[lineIndex];
                          const lineEnd = lineStart + lines[lineIndex].length;

                          // Check if this cell ref is on this line
                          if (ref.start >= lineStart && ref.start < lineEnd) {
                            const lineNumber = lineIndex + 1;
                            const line = lines[lineIndex];

                            // Calculate column positions relative to the line
                            const textStartColumn = ref.start - lineStart + 1; // Monaco is 1-indexed
                            const textEndColumn = ref.end - lineStart + 1; // End position (exclusive in Monaco)
                            const zeroBasedStart = ref.start - lineStart;
                            const zeroBasedEnd = ref.end - lineStart;
                            const coreText = line.slice(zeroBasedStart, zeroBasedEnd);

                            // Check if spaces exist before and after, but allow tight binding with colons for ranges
                            const charBefore = zeroBasedStart > 0 ? line[zeroBasedStart - 1] : '';
                            const charAfter = zeroBasedEnd < line.length ? line[zeroBasedEnd] : '';
                            const beforeIsSpace = /\s/.test(charBefore);
                            const afterIsSpace = /\s/.test(charAfter);
                            const beforeIsColon = charBefore === ':';
                            const afterIsColon = charAfter === ':';
                            const shouldInsertSpaceBefore = !beforeIsSpace && !beforeIsColon;
                            const shouldInsertSpaceAfter = !afterIsSpace && !afterIsColon;
                            const hasSpaceBefore = beforeIsSpace;
                            const hasSpaceAfter = afterIsSpace;

                            // Normalize whitespace around range colons to keep them tight (e.g., "A1:B2")
                            const colonIndex = coreText.indexOf(':');
                            if (colonIndex !== -1) {
                              let whitespaceBefore = 0;
                              for (let i = colonIndex - 1; i >= 0 && /\s/.test(coreText[i]); i--) {
                                whitespaceBefore++;
                              }
                              if (whitespaceBefore > 0) {
                                edits.push({
                                  range: new monaco.Range(
                                    lineNumber,
                                    textStartColumn + colonIndex - whitespaceBefore,
                                    lineNumber,
                                    textStartColumn + colonIndex
                                  ),
                                  text: ''
                                });
                              }

                              let whitespaceAfter = 0;
                              for (let i = colonIndex + 1; i < coreText.length && /\s/.test(coreText[i]); i++) {
                                whitespaceAfter++;
                              }
                              if (whitespaceAfter > 0) {
                                edits.push({
                                  range: new monaco.Range(
                                    lineNumber,
                                    textStartColumn + colonIndex + 1,
                                    lineNumber,
                                    textStartColumn + colonIndex + 1 + whitespaceAfter
                                  ),
                                  text: ''
                                });
                              }
                            }

                            // Get current cursor position to check if we need to move it
                            const currentPosition = editor.getPosition();
                            const isAtEndOfCell = currentPosition && 
                                                 currentPosition.lineNumber === lineNumber &&
                                                 currentPosition.column === textEndColumn;

                            // Insert a literal space before if needed (ALWAYS, even at start of line)
                            // Spaces are part of the actual text so cursor alignment stays correct
                            if (shouldInsertSpaceBefore) {
                              edits.push({
                                range: new monaco.Range(lineNumber, textStartColumn, lineNumber, textStartColumn),
                                text: ' '
                              });
                            }

                            // Insert a literal space after if needed (unless we're building a range with :)
                            if (shouldInsertSpaceAfter) {
                              edits.push({
                                range: new monaco.Range(lineNumber, textEndColumn, lineNumber, textEndColumn),
                                text: ' '
                              });
                            }

                            // If cursor was at the end of the cell reference, move it 2 positions ahead (after the closing space)
                            if (isAtEndOfCell && (shouldInsertSpaceBefore || shouldInsertSpaceAfter)) {
                              shouldMoveCursor = true;
                              // After insertion, account for padding we added
                              const spacesToAdd = 
                                (shouldInsertSpaceBefore ? 1 : 0) + 
                                (shouldInsertSpaceAfter ? 1 : 0);
                              cursorMovePosition = new monaco.Position(lineNumber, textEndColumn + spacesToAdd);
                            }

                            const chipColorClass = getChipColorClass(ref.fullMatch || ref.range || ref.text);
                            const normalizedReference = normalizeCellReference(ref.text || ref.range || ref.fullMatch);
                            if (chipColorClass && normalizedReference) {
                              const targetRefs = expandReferenceToCells(normalizedReference);
                              targetRefs.forEach((cellRef) => {
                                if (cellRef) {
                                  newHighlightAssignments.set(cellRef, {
                                    className: chipColorClass,
                                    groupKey: normalizedReference
                                  });
                                }
                              });
                            }

                            // Decorate the FULL range " A1 " (including spaces) so cursor aligns correctly
                            // Monaco may split the decoration on spaces, but CSS will merge them visually
                            // This ensures cursor positioning matches the visual chip boundaries
                            if (hasSpaceBefore && hasSpaceAfter) {
                              // Decorate spaces and core separately to control borders precisely
                              const leftSpaceStart = textStartColumn - 1;
                              const leftSpaceEnd = textStartColumn;
                              const rightSpaceStart = textEndColumn;
                              const rightSpaceEnd = textEndColumn + 1;

                              // Track combined chip range from left space start to right space end
                              const chipRange = new monaco.Range(lineNumber, leftSpaceStart, lineNumber, rightSpaceEnd);
                              newChipRanges.push(chipRange);

                              // Left space (rounded left edge)
                              const leftClassNames = ['cell-chip-inline', 'cell-chip-inline-left'];
                              if (chipColorClass) {
                                leftClassNames.push(chipColorClass);
                              }
                              decorations.push({
                                range: new monaco.Range(lineNumber, leftSpaceStart, lineNumber, leftSpaceEnd),
                                options: {
                                  inlineClassName: leftClassNames.join(' '),
                                  stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges
                                }
                              });

                              // Core reference (no rounded corners, full background)
                              const coreClassNames = ['cell-chip-inline', 'cell-chip-inline-core'];
                              if (chipColorClass) {
                                coreClassNames.push(chipColorClass);
                              }
                              decorations.push({
                                range: new monaco.Range(lineNumber, textStartColumn, lineNumber, textEndColumn),
                                options: {
                                  inlineClassName: coreClassNames.join(' '),
                                  hoverMessage: { value: `Cell reference: ${ref.range}` },
                                  stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges
                                }
                              });

                              // Right space (rounded right edge)
                              const rightClassNames = ['cell-chip-inline', 'cell-chip-inline-right'];
                              if (chipColorClass) {
                                rightClassNames.push(chipColorClass);
                              }
                              decorations.push({
                                range: new monaco.Range(lineNumber, rightSpaceStart, lineNumber, rightSpaceEnd),
                                options: {
                                  inlineClassName: rightClassNames.join(' '),
                                  stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges
                                }
                              });
                            }
                            // If spaces don't exist, skip decoration - will be created after spaces are inserted and updateCellChips re-runs
                            break;
                          }
                        }
                      });

                      // Apply edits if any (insert spaces)
                      if (edits.length > 0) {
                        window.isProgrammaticCursorChange = true;
                        editor.executeEdits('insert-cell-chip-spaces', edits);

                        // Move cursor 2 positions ahead if it was at the end of a cell reference
                        if (shouldMoveCursor && cursorMovePosition) {
                          editor.setPosition(cursorMovePosition);
                        }

                        Promise.resolve().then(() => {
                          window.isProgrammaticCursorChange = false;
                          // Re-run updateCellChips after edits are applied to get correct positions
                          // Use a small delay to ensure the model has updated
                          setTimeout(() => {
                            updateCellChips();
                          }, 10);
                        });
                        return; // Exit early - will re-run after edits are applied
                      }

                      updateGridHighlights(newHighlightAssignments);

                      // Update decorations (deltaDecorations replaces old with new)
                      // Only run if no edits were needed

                      cellChipDecorations = editor.deltaDecorations(cellChipDecorations, decorations);
                      cellChipRanges = newChipRanges;

                      // Re-evaluate cursor suppression state
                      const currentPosition = editor.getPosition();
                      const inChip = isPositionInsideChip(currentPosition);
                      if (inChip !== isCursorInChip) {
                        isCursorInChip = inChip;
                        suppressSuggestionsInChip(inChip);
                      }

                      // Store the processed text to help prevent duplicates
                      lastProcessedText = fullText;
                      scheduleChipInstantiationSettled();
                    }

                    function getChipRangeAtPosition(position) {
                      if (!position || cellChipRanges.length === 0) return null;
                      return cellChipRanges.find((range) => range.containsPosition(position)) || null;
                    }

                    function buildChipInsertionText(reference, lineContent = '', startColumn, endColumn) {
                      const beforeIndex = startColumn - 2;
                      const afterIndex = endColumn - 1;
                      const beforeChar = beforeIndex >= 0 ? lineContent[beforeIndex] : '';
                      const afterChar = afterIndex < lineContent.length ? lineContent[afterIndex] : '';
                      const needsLeadingSpace = beforeChar && !/\s/.test(beforeChar);
                      const needsTrailingSpace = !afterChar || !/\s/.test(afterChar);
                      const leading = needsLeadingSpace ? ' ' : '';
                      const trailing = needsTrailingSpace ? ' ' : '';
                      return `${leading}${reference}${trailing}`;
                    }

                    function replaceRangeWithChip(range, reference, options = {}) {
                      const currentEditor = window.monacoEditor || editor;
                      if (!currentEditor || !reference) return null;
                      const model = currentEditor.getModel();
                      if (!model) return null;

            const previousValueInRange = model.getValueInRange(range) || '';
            preserveChipColorAssignment(previousValueInRange, reference);

                      markChipInstantiationStarted();

                      const lineContent = model.getLineContent(range.startLineNumber) || '';
                      const chipText = options.preservePadding
                        ? ` ${reference} `
                        : buildChipInsertionText(reference, lineContent, range.startColumn, range.endColumn);

                      let newRange = null;
                      window.isProgrammaticCursorChange = true;
                      try {
                      currentEditor.executeEdits('replace-chip', [{
                        range,
                        text: chipText,
                        forceMoveMarkers: true
                      }]);

                      const targetColumn = range.startColumn + chipText.length;
                      currentEditor.setPosition(new monaco.Position(range.startLineNumber, Math.max(1, targetColumn)));
                        newRange = new monaco.Range(
                          range.startLineNumber,
                          range.startColumn,
                          range.startLineNumber,
                          range.startColumn + Math.max(0, chipText.length)
                        );
                      } finally {
                        window.isProgrammaticCursorChange = false;
                      }

                      Promise.resolve().then(() => {
                        updateCellChips();
                      });
                      return newRange;
                    }

                    function insertChipAtPosition(position, reference, options = {}) {
                      const insertRange = new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column);
                      return replaceRangeWithChip(insertRange, reference, options);
                    }

                    function getTokenRangeAtPosition(model, position) {
                      if (!model || !position) return null;
                      const word = model.getWordAtPosition(position);
                      if (!word) return null;
                      return new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn);
                    }

                    function applyReferenceToFormula(reference) {
                      if (!reference) return;
                      const currentEditor = window.monacoEditor || editor;
                      if (!currentEditor) return;
                      const model = currentEditor.getModel();
                      if (!model) return;

                      const position = currentEditor.getPosition();
                      if (!position) return;

                      let targetRange = getChipRangeAtPosition(position);
                      if (!targetRange && currentMode === MODES.ENTER) {
                        targetRange = getPointerChipRangeForPosition(position);
                      }
                      let appliedRange = null;

                      if (targetRange) {
                        appliedRange = replaceRangeWithChip(targetRange, reference, { preservePadding: true });
                      } else {
                        const tokenRange = getTokenRangeAtPosition(model, position);
                        if (tokenRange) {
                          const tokenOptions = currentMode === MODES.ENTER ? { preservePadding: true } : undefined;
                          appliedRange = replaceRangeWithChip(tokenRange, reference, tokenOptions);
                        } else {
                          const insertionOptions = currentMode === MODES.ENTER ? { preservePadding: true } : undefined;
                          appliedRange = insertChipAtPosition(position, reference, insertionOptions);
                        }
                      }

                      if (currentMode === MODES.ENTER && appliedRange) {
                        updateEnterModePointerChipRange(appliedRange);
                        updateEnterModePointerCoordinatesFromReference(reference);
                      } else if (currentMode === MODES.EDIT) {
                        if (!enterModePointerState.active) {
                          initializeEnterModePointerState(window.selectedCell || null);
                        }
                        updateEnterModePointerCoordinatesFromReference(reference);
                      }

                      if (typeof currentEditor.focus === 'function') {
                        currentEditor.focus();
                      }
                    }
                    window.applyReferenceToFormula = applyReferenceToFormula;

                    // Auto-capitalize function names (unless inside quotes)
                    let capitalizeTimeout;
                    editor.onDidChangeModelContent((e) => {
                      // Skip if this is a programmatic change
                      if (window.isProgrammaticCursorChange) return;
                      if (currentMode === MODES.EDIT && editModePointerArmed && e && Array.isArray(e.changes)) {
                        const insertedDelimiter = e.changes.some(change => change?.text && POINTER_NAV_SEPARATOR_REGEX.test(change.text));
                        if (insertedDelimiter) {
                          disarmEditModePointerNavigation();
                        }
                      }

                      clearTimeout(capitalizeTimeout);
                      capitalizeTimeout = setTimeout(() => {
                        const model = editor.getModel();
                        if (!model) return;

                        const position = editor.getPosition();
                        if (!position) return;

                        // Get current line
                        const allLines = model.getValue().split('\n');
                        if (allLines.length < 1) return;

                        const currentLineIndex = position.lineNumber - 1;
                        // Line 1 is now valid

                        const currentLine = allLines[currentLineIndex];
                        const textBeforeCursor = currentLine.substring(0, position.column - 1);
                        const textAfterCursor = currentLine.substring(position.column - 1);
                        const fullLine = currentLine;

                        // Don't capitalize if we're inside quotes
                        if (isInsideQuotes(fullLine, position.column - 1)) {
                          return;
                        }

                        // Find function names in the line and capitalize them
                        // Match word boundaries followed by function names
                        // Use global function list if available
                        if (!window.hyperFormulaFunctions || window.hyperFormulaFunctions.length === 0) {
                          return; // Function list not loaded yet
                        }
                        const functionNames = window.hyperFormulaFunctions.map(f => f.name);
                        const edits = [];

                        // Build regex to match function names (case-insensitive, word boundaries)
                        functionNames.forEach(funcName => {
                          // Escape special regex characters
                          const escapedName = funcName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                          // Match function name at word boundary, case-insensitive
                          const regex = new RegExp(`\\b(${escapedName})\\b`, 'gi');

                          let match;
                          while ((match = regex.exec(fullLine)) !== null) {
                            const matchedText = match[0];
                            const matchedName = match[1];

                            // Check if this match is inside quotes
                            if (isInsideQuotes(fullLine, match.index + matchedText.length)) {
                              continue; // Skip if inside quotes
                            }

                            // Only capitalize if it's not already the correct case
                            if (matchedText !== funcName) {
                              // Check what's before the function name
                              const beforeChar = match.index > 0 ? fullLine[match.index - 1] : '';
                              const isValidBefore = match.index === 0 || /[^A-Za-z0-9_$]/.test(beforeChar);

                              // Check what's after the function name
                              const afterText = fullLine.substring(match.index + matchedText.length);
                              // After the function name should be:
                              // - Opening parenthesis (function call)
                              // - Space or end of line (standalone function name)
                              // - Non-alphanumeric character (part of expression)
                              // - Empty (cursor is right after the function name)
                              const isValidAfter = afterText.length === 0 || 
                                                   /^\s*\(/.test(afterText) || 
                                                   /^\s*$/.test(afterText) || 
                                                   /^\s*[^A-Za-z0-9_]/.test(afterText);

                              // Capitalize if:
                              // 1. It's preceded by valid character (start of line, operator, etc.) OR at start of line
                              // 2. It's followed by valid character (opening paren, space, end, etc.)
                              // This allows capitalization as you type, even before the opening paren
                              if (isValidBefore && isValidAfter) {
                                edits.push({
                                  range: new monaco.Range(
                                    currentLineIndex + 1,
                                    match.index + 1,
                                    currentLineIndex + 1,
                                    match.index + matchedText.length + 1
                                  ),
                                  text: funcName
                                });
                              }
                            }
                          }
                        });

                        // Auto-capitalize cell references (like "a1" -> "A1")
                        // Match cell references: 1-3 lowercase letters followed by digits, optionally with $ anchors
                        // Use a simpler pattern that matches cell references more reliably
                        const cellRefPattern = /([a-z]{1,3})(\$?)(\d+)(\$?)/gi;
                        let cellMatch;
                        const processedRanges = []; // Track processed ranges to avoid duplicates

                        while ((cellMatch = cellRefPattern.exec(fullLine)) !== null) {
                          const colPart = cellMatch[1]; // lowercase column letters
                          const dollarBeforeCol = cellMatch[2] || ''; // $ before column (if any)
                          const rowPart = cellMatch[3]; // row number
                          const dollarBeforeRow = cellMatch[4] || ''; // $ before row (if any)

                          // Reconstruct the full matched cell reference
                          const matchedText = colPart + dollarBeforeCol + rowPart + dollarBeforeRow;
                          const matchStartIndex = cellMatch.index;
                          const matchEndIndex = matchStartIndex + matchedText.length;

                          // Check if this range was already processed (avoid duplicates from overlapping matches)
                          const alreadyProcessed = processedRanges.some(range => 
                            matchStartIndex >= range.start && matchEndIndex <= range.end
                          );
                          if (alreadyProcessed) continue;
                          processedRanges.push({ start: matchStartIndex, end: matchEndIndex });

                          // Skip if inside quotes
                          if (isInsideQuotes(fullLine, matchEndIndex)) {
                            continue;
                          }

                          // Check what's before and after the cell reference
                          const beforeChar = matchStartIndex > 0 ? fullLine[matchStartIndex - 1] : '';
                          const afterChar = matchEndIndex < fullLine.length ? fullLine[matchEndIndex] : '';

                          // Valid if preceded by operator, comma, parenthesis, space, or start of line
                          // But not by a letter or digit (would be part of a larger word)
                          const isValidBefore = matchStartIndex === 0 || /[^A-Za-z0-9_]/.test(beforeChar);

                          // Valid if followed by operator, comma, parenthesis, space, or end of line
                          // But not by a letter or digit (would be part of a larger word)
                          const isValidAfter = matchEndIndex >= fullLine.length || 
                                               /[^A-Za-z0-9_]/.test(afterChar);

                          // Only capitalize if it looks like a valid cell reference context and column is lowercase
                          if (isValidBefore && isValidAfter && colPart === colPart.toLowerCase()) {
                            // Build new cell reference with uppercase column, preserving $ anchors
                            let newCellRef = colPart.toUpperCase();
                            if (dollarBeforeCol) {
                              newCellRef = '$' + newCellRef;
                            }
                            newCellRef += rowPart;
                            if (dollarBeforeRow) {
                              newCellRef += '$';
                            }

                            // Only edit if different
                            if (matchedText !== newCellRef) {
                              edits.push({
                                range: new monaco.Range(
                                  currentLineIndex + 1,
                                  matchStartIndex + 1,
                                  currentLineIndex + 1,
                                  matchEndIndex + 1
                                ),
                                text: newCellRef
                              });
                            }
                          }
                        }

                        // Apply edits if any
                        if (edits.length > 0) {
                          // Set flag to prevent recursive changes
                          window.isProgrammaticCursorChange = true;
                          editor.executeEdits('auto-capitalize-functions', edits);
                          // Reset flag after a microtask
                          Promise.resolve().then(() => {
                            window.isProgrammaticCursorChange = false;
                          });
                        }
                      }, 100); // Small delay to allow typing to complete
                    });

                    // Update cell chips on content change
                    editor.onDidChangeModelContent(() => {
                      clearTimeout(chipUpdateTimeout);
                      chipUpdateTimeout = setTimeout(() => {
                        updateCellChips();
                      }, 100);
                    });

                    editor.onDidType((text) => {
                      if (window.isProgrammaticCursorChange) return;
                      const model = editor.getModel();
                      if (!model) return;

                      if (currentMode === MODES.ENTER || currentMode === MODES.EDIT) {
                        resetPointerAfterSeparatorIfNeeded(text);
                      }

                      if (text === ':') {
                        const position = editor.getPosition();
                        if (!position) return;
                        const colonLine = position.lineNumber;
                        const colonColumn = Math.max(1, position.column - 1);
                        const beforePosition = new monaco.Position(colonLine, Math.max(1, colonColumn - 1));
                        const chipRange = getChipRangeAtPosition(beforePosition);
                        if (!chipRange) return;
                        if (beforePosition.column < chipRange.endColumn - 1) {
                          return;
                        }
                        const insertColumn = Math.max(chipRange.startColumn, chipRange.endColumn - 1);
                        const colonRange = new monaco.Range(colonLine, colonColumn, colonLine, colonColumn + 1);

                        window.isProgrammaticCursorChange = true;
                        editor.executeEdits('chip-colon-adjust', [
                          {
                            range: colonRange,
                            text: '',
                            forceMoveMarkers: true
                          },
                          {
                            range: new monaco.Range(colonLine, insertColumn, colonLine, insertColumn),
                            text: ':',
                            forceMoveMarkers: true
                          }
                        ]);
                        const lockPosition = new monaco.Position(colonLine, insertColumn + 1);
                        editor.setPosition(lockPosition);
                        Promise.resolve().then(() => {
                          window.isProgrammaticCursorChange = false;
                        });
                        return;
                      }

                    });

                    // Initial chip update
                    updateCellChips();

                    // Auto-expand function snippets when opening parenthesis is typed
                    // Example: typing "IF(" should expand to "IF(${1:logical}, ${2:value}, ${3:value})"
                    editor.onDidChangeModelContent((e) => {
                      // Skip if this is a programmatic change
                      if (window.isProgrammaticCursorChange) return;

                      // Check each change to see if it's just an opening parenthesis
                      e.changes.forEach(change => {
                        // Only process changes that add text (not deletions)
                        if (!change.text || !change.text.includes('(')) {
                          return;
                        }

                        // When pasting full formulas (which include commas, quotes, etc.), don't auto-expand snippets.
                        // Only auto-expand when the inserted text consists solely of parentheses/whitespace (i.e. typing "(" or Monaco's "()")
                        if (!/^[()\s]*$/.test(change.text)) {
                          return;
                        }

                        // Find the position of the first "(" within the inserted text
                        const parenIndexInText = change.text.indexOf('(');
                        if (parenIndexInText === -1) {
                          return;
                        }

                        const model = editor.getModel();
                        if (!model) return;

                        // Determine the actual column (1-based) where "(" ended up
                        const parenColumn = change.range.startColumn + parenIndexInText;
                        const parenLine = change.range.startLineNumber;

                        // Get the current line content (after the change)
                        const currentLine = model.getLineContent(parenLine);

                        const parenZeroIndex = Math.max(0, parenColumn - 1);

                        // Get text before the opening parenthesis that was just inserted
                        const textBeforeParen = currentLine.substring(0, parenZeroIndex);

                        // Don't expand if inside quotes
                        if (isInsideQuotes(textBeforeParen, textBeforeParen.length)) {
                          return;
                        }

                        // Match a function name immediately before the opening parenthesis
                        // Pattern: function name followed by optional whitespace and then (
                        const functionNamePattern = /([A-Za-z_][A-Za-z0-9_]*)\s*$/;
                        const match = textBeforeParen.match(functionNamePattern);

                        if (match) {
                          const funcName = match[1];

                          // Check if this matches a function name (case-insensitive)
                          if (window.hyperFormulaFunctions && window.hyperFormulaFunctions.length > 0) {
                            const funcNameUpper = funcName.toUpperCase();
                            const matchingFunction = window.hyperFormulaFunctions.find(func =>
                              func.name.toUpperCase() === funcNameUpper
                            );

                            if (matchingFunction) {
                              // Don't expand if it's already a snippet (check if there's already placeholders after the paren)
                              const textAfterParen = currentLine.substring(parenZeroIndex + 1);
                              if (textAfterParen.includes('${') || textAfterParen.includes('$1')) {
                                return; // Already expanded
                              }

                              // Also check if the function name + ( is already part of a snippet
                              const funcStartZeroIndex = Math.max(0, parenZeroIndex - funcName.length);
                              const textWithFuncAndParen = currentLine.substring(
                                funcStartZeroIndex,
                                Math.min(currentLine.length, parenZeroIndex + change.text.length)
                              );
                              if (textWithFuncAndParen.includes('${')) {
                                return; // Already part of a snippet
                              }

                              // Check that it's not a cell reference (like A1() - though unlikely)
                              const isCellReference = /^[A-Z]{1,3}\$?\d+\$?$/i.test(funcName);
                              if (isCellReference) {
                                return; // Don't expand cell references
                              }

                              // Found a matching function - expand to snippet
                              const snippetTemplate = createFunctionSnippet(matchingFunction);
                              if (!snippetTemplate) {
                                return;
                              }

                              // Calculate the range: from start of function name to after the opening paren (and optional auto-close ))
                              const funcNameStartCol = parenColumn - funcName.length;
                              let funcNameEndCol = parenColumn + 1; // include the opening parenthesis

                              const insertedClosingIndex = change.text.indexOf(')');
                              if (insertedClosingIndex !== -1) {
                                const closingColumn = change.range.startColumn + insertedClosingIndex + 1;
                                funcNameEndCol = Math.max(funcNameEndCol, closingColumn);
                              } else {
                                const charAfterParen = currentLine.charAt(parenZeroIndex + 1);
                                if (charAfterParen === ')') {
                                  funcNameEndCol = Math.max(funcNameEndCol, parenColumn + 2);
                                }
                              }

                              const replaceRange = new monaco.Range(
                                parenLine,
                                funcNameStartCol,
                                parenLine,
                                funcNameEndCol
                              );

                              // Use snippet controller to insert the snippet
                              window.isProgrammaticCursorChange = true;

                              try {
                                const snippetController = editor.getContribution('snippetController2');
                                if (snippetController) {
                                  editor.setSelection(replaceRange);
                                  snippetController.insert(snippetTemplate);
                                  console.log('Auto-expanded function snippet:', matchingFunction.name);
                                } else {
                                  throw new Error('SnippetController2 not available');
                                }
                              } catch (e) {
                                console.error('Failed to auto-expand snippet:', e);
                              }

                              Promise.resolve().then(() => {
                                window.isProgrammaticCursorChange = false;
                              });
                            }
                          }
                        }
                      });
                    });

                    // Initial validation
                    setTimeout(validateFormula, 100);

                    console.log('HyperFormula IntelliSense registered successfully');
                  }).catch(err => {
                    console.error('Failed to load HyperFormula functions:', err);
                  });

                  // Set up Tab key handler for comma/parenthesis completion and function snippet insertion
                  // Priority: If IntelliSense is showing, let Monaco accept the suggestion (which will insert snippet)
                  // Otherwise, handle Tab for our custom completions (function snippets, commas, closing parens)
                  editor.addCommand(monaco.KeyCode.Tab, () => {
                    const model = editor.getModel();
                    if (!model) return;

                    const position = editor.getPosition();
                    if (!position) {
                      return;
                    }

                    // FIRST: Check if IntelliSense is showing
                    // If so, let Monaco handle Tab to accept the suggestion (which will insert snippet with snippet mode)
                    // Monaco's acceptSuggestionOnTab: 'on' will handle this automatically
                    // But we need to check if suggestions are visible first
                    const suggestController = editor.getContribution('suggestController');
                    if (suggestController && suggestController.widget && suggestController.widget.value) {
                      const widget = suggestController.widget.value;
                      // Check if the suggest widget is open and visible
                      if (widget && typeof widget.isOpen === 'function' && widget.isOpen()) {
                        // IntelliSense is showing - let Monaco handle Tab to accept the suggestion
                        // Don't intercept - return undefined so Monaco's default handler runs
                        // Monaco will accept the suggestion which will insert the snippet with snippet mode
                        return;
                      }
                    }

                    // IntelliSense is NOT showing - insert 4 spaces
                    editor.executeEdits('tab-indent', [{
                      range: new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column),
                      text: '    ' // 4 spaces
                    }]);
                  });

                  // Prevent IntelliSense popup when F8 is pressed (used for navigating to next problem)
                  editor.addCommand(monaco.KeyCode.F8, () => {
                    // Check if suggestions are visible and close them
                    const suggestController = editor.getContribution('suggestController');
                    if (suggestController && suggestController.widget && suggestController.widget.value) {
                      const widget = suggestController.widget.value;
                      if (widget && typeof widget.isOpen === 'function' && widget.isOpen()) {
                        widget.cancel(); // Close IntelliSense if open
                      }
                    }
                    // Prevent IntelliSense from opening by temporarily disabling quick suggestions
                    // Store original setting
                    const originalQuickSuggestions = editor.getOption(monaco.editor.EditorOption.quickSuggestions);
                    // Temporarily disable quick suggestions
                    editor.updateOptions({ quickSuggestions: false });
                    // Use setTimeout to re-enable after a brief moment (after F8 navigation completes)
                    setTimeout(() => {
                      editor.updateOptions({ quickSuggestions: originalQuickSuggestions });
                    }, 100);
                    // Let Monaco handle F8 navigation (go to next problem)
                    // Return undefined to allow default F8 behavior
                    return;
                  });

                  const handleEditorEnterKey = (event) => {
                    event.preventDefault();
                    if (currentMode === MODES.ENTER || currentMode === MODES.EDIT) {
                      const moveDelta = { row: event.shiftKey ? -1 : 1, col: 0 };
                      const committed = commitEditorChanges(moveDelta);
                      if (!committed) {
                        const activeEditor = window.monacoEditor || editor;
                        if (activeEditor && typeof activeEditor.focus === 'function') {
                          activeEditor.focus();
                        }
                      }
                      return;
                    }
                    commitEditorChanges(null);
                  };

                  editor.onKeyDown((event) => {
                    if (event.keyCode === monaco.KeyCode.Enter && !event.ctrlKey && !event.metaKey) {
                      handleEditorEnterKey(event);
                    }
                  });

                  const moveCursorToNextLineOrAppend = () => {
                    const targetEditor = window.monacoEditor || editor;
                    if (!targetEditor) return;
                    const model = targetEditor.getModel?.();
                    if (!model) return;
                    const position = targetEditor.getPosition();
                    if (!position) return;

                    const currentLine = position.lineNumber;
                    const totalLines = model.getLineCount();
                    if (currentLine < totalLines) {
                      const nextLine = currentLine + 1;
                      const nextLineMaxColumn = model.getLineMaxColumn(nextLine);
                      const newColumn = Math.min(position.column, nextLineMaxColumn);
                      targetEditor.setPosition(new monaco.Position(nextLine, newColumn));
                      updateEditorStatusDisplay({ lineNumber: nextLine, column: newColumn });
                      return;
                    }

                    const lastColumn = model.getLineMaxColumn(totalLines);
                    window.isProgrammaticCursorChange = true;
                    targetEditor.executeEdits('ctrl-enter-next-line', [{
                      range: new monaco.Range(totalLines, lastColumn, totalLines, lastColumn),
                      text: '\n',
                      forceMoveMarkers: true
                    }]);
                    const appendedLine = totalLines + 1;
                    targetEditor.setPosition(new monaco.Position(appendedLine, 1));
                    updateEditorStatusDisplay({ lineNumber: appendedLine, column: 1 });
                    Promise.resolve().then(() => {
                      window.isProgrammaticCursorChange = false;
                    });
                  };

                  editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter, () => {
                    moveCursorToNextLineOrAppend();
                  });

                  window.monacoEditor = editor;
                  console.log("Monaco Editor created successfully", editor);

                  // Add click listener to Monaco editor to enter edit mode when clicked
                  const editorDomNode = editor.getDomNode();
                  if (editorDomNode) {
                    editorDomNode.addEventListener('click', (e) => {
                      let targetPosition = null;
                      try {
                        targetPosition = editor.getTargetAtClientPoint(e.clientX, e.clientY);
                      } catch (err) {
                        // Ignore if positioning fails
                      }

                      if (currentMode === MODES.EDIT) {
                        editor.updateOptions({ cursorBlinking: 'blink' });
                        if (targetPosition && targetPosition.position) {
                          const { lineNumber, column } = targetPosition.position;
                          editor.setPosition({ lineNumber, column });
                          updateEditorStatusDisplay({ lineNumber, column });
                          if (typeof syncEditingCellDisplayWithPane === 'function') {
                            requestAnimationFrame(() => syncEditingCellDisplayWithPane(false, editor));
                          }
                        }
                        return;
                      }

                      if (!window.selectedCell) {
                        return;
                      }

                      if (currentMode === MODES.ENTER) {
                        promoteEnterModeToEdit(targetPosition);
                      } else {
                        enterEditMode(window.selectedCell, targetPosition);
                      }
                    });
                  }
                } catch (monacoError) {
                  console.error("Failed to initialize Monaco Editor:", monacoError);
                }
              }).catch(monacoError => {
                console.error("Failed to load Monaco Editor:", monacoError);
                console.error("Error details:", monacoError.stack);
                // Continue without Monaco - grid should still work
              });
            }, 100);
          }

          // Function to enter edit mode for a cell
          function enterEditMode(cell, clickPosition, options = {}) {
            if (!cell) return;

            const isHeader = cell.tagName === "TH" || cell === cell.parentElement.querySelector("td:first-child");
            if (isHeader) return;

            const desiredMode = options.modeOverride || MODES.EDIT;
            setMode(desiredMode);
            cancelFormulaSelection();

            if (selectionOverlayElement) {
              selectionOverlayElement.style.display = "none";
            }
            if (fillHandleElement) {
              fillHandleElement.style.display = "none";
            }

            clearSelection();
            selectedCells.add(cell);
            cell.classList.add("selected");
            cell.classList.add("edit-mode");
            cell.setAttribute("data-selection-edge", "top bottom left right");
            window.selectedCell = cell;
            if (desiredMode === MODES.ENTER || desiredMode === MODES.EDIT) {
              initializeEnterModePointerState(cell);
            } else {
              resetEnterModePointerState();
            }
            updateCellRangePill();
            updateHeaderHighlightsFromSelection();

            const cellRef = cell.getAttribute("data-ref");
            const hasInitialText = Object.prototype.hasOwnProperty.call(options, 'initialText');
            const shouldPreserveExistingContent = !!options.preserveExistingContent;
            let editorValue = "";

            if (!shouldPreserveExistingContent) {
              if (hasInitialText) {
                editorValue = (options.initialText || "").toString();
              } else {
                editorValue = getStoredFormulaText(cell) || "";
                if (!editorValue && cellRef && window.hf) {
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
            }

            const currentEditor = window.monacoEditor || editor;
            if (currentEditor && typeof currentEditor.setValue === 'function') {
              const model = currentEditor.getModel();
              let workingValue = editorValue;
              if (shouldPreserveExistingContent && model) {
                workingValue = model.getValue();
              }

              let isEditorEmpty = true;
              if (model) {
                const currentLines = (workingValue || "").split('\n');
                isEditorEmpty = !currentLines.some(line => line.trim().length > 0);
              }

              let fullValue = ensureSevenLines(editorValue || "");
              let targetLine = 1;
              let targetColumn = 1;

              if (clickPosition && clickPosition.position && !isEditorEmpty) {
                targetLine = clickPosition.position.lineNumber;
                targetColumn = clickPosition.position.column;
              } else if (workingValue && workingValue.trim().length > 0) {
                const lines = workingValue.split('\n');
                const lastLine = lines[lines.length - 1] || "";
                targetLine = lines.length > 0 ? lines.length : 1;
                targetColumn = Math.max(1, lastLine.length + 1);
              }

              window.isProgrammaticCursorChange = true;

              if (!shouldPreserveExistingContent) {
                currentEditor.setValue(fullValue);
              }

              currentEditor.updateOptions({ 
                cursorBlinking: 'blink',
                cursorStyle: 'line'
              });

              currentEditor.setPosition({ lineNumber: targetLine, column: targetColumn });
              if (typeof currentEditor.focus === 'function') {
                currentEditor.focus();
              }
              syncEditingCellDisplayWithPane(true, currentEditor);

              Promise.resolve().then(() => {
                window.isProgrammaticCursorChange = false;
                if (typeof updateCellChips === 'function') {
                  requestAnimationFrame(() => updateCellChips());
                }
              });

              setTimeout(() => {
                if (currentEditor && typeof currentEditor.focus === 'function') {
                  currentEditor.focus();
                  currentEditor.setPosition({ lineNumber: targetLine, column: targetColumn });
                }
              }, 50);
            }
          }

          function getFormulaPaneSnapshotText(activeEditorInstance) {
            const activeEditor = activeEditorInstance || window.monacoEditor || editor;
            if (!activeEditor || typeof activeEditor.getValue !== 'function') {
              return "";
            }
            const rawValue = activeEditor.getValue() || "";
            return rawValue.replace(/\r\n/g, '\n');
          }

          function getEditorCursorSnapshot(activeEditorInstance) {
            const activeEditor = activeEditorInstance || window.monacoEditor || editor;
            if (!activeEditor || typeof activeEditor.getPosition !== 'function') {
              return null;
            }
            const model = activeEditor.getModel?.();
            if (!model || typeof model.getOffsetAt !== 'function') {
              return null;
            }
            const selection = typeof activeEditor.getSelection === 'function' ? activeEditor.getSelection() : null;
            const position = selection ? selection.getPosition() : activeEditor.getPosition();
            if (!position) {
              return null;
            }
            const offset = model.getOffsetAt(position);
            return {
              offset,
              lineNumber: position.lineNumber,
              column: position.column
            };
          }

          function convertFormulaSegmentToHtml(segment) {
            if (!segment) return '';
            return escapeHtml(segment).replace(/\n/g, '<br>');
          }

          function renderFormulaTextWithCaret(fullText, caretInfo) {
            const text = typeof fullText === 'string' ? fullText : '';
            if (!caretInfo || typeof caretInfo.offset !== 'number') {
              const html = convertFormulaSegmentToHtml(text);
              return html || '<span class="cell-edit-caret" aria-hidden="true"></span>';
            }
            const clampedOffset = Math.max(0, Math.min(text.length, caretInfo.offset));
            const before = text.slice(0, clampedOffset);
            const after = text.slice(clampedOffset);
            const beforeHtml = convertFormulaSegmentToHtml(before);
            const afterHtml = convertFormulaSegmentToHtml(after);
            return `${beforeHtml}<span class="cell-edit-caret" aria-hidden="true"></span>${afterHtml}`;
          }

          function isMonacoEditorFocused(editorInstance) {
            if (!editorInstance) return false;
            if (typeof editorInstance.hasTextFocus === 'function') {
              return editorInstance.hasTextFocus();
            }
            const domNode = typeof editorInstance.getDomNode === 'function' ? editorInstance.getDomNode() : null;
            if (!domNode) return false;
            return domNode === document.activeElement || domNode.contains(document.activeElement);
          }

          function blurMonacoEditor(editorInstance) {
            const instance = editorInstance || window.monacoEditor || editor;
            if (!instance) return;

            if (typeof instance.blur === 'function') {
              instance.blur();
            }

            const domNode = typeof instance.getDomNode === 'function' ? instance.getDomNode() : null;
            if (!domNode) return;

            const activeElement = document.activeElement;
            if (activeElement && domNode.contains(activeElement)) {
              if (typeof activeElement.blur === 'function') {
                activeElement.blur();
              }
            }

            const textArea = domNode.querySelector('textarea');
            if (textArea && typeof textArea.blur === 'function') {
              textArea.blur();
            }
          }

          function syncEditingCellDisplayWithPane(force = false, activeEditorInstance = null) {
            if (currentMode === MODES.READY || !window.selectedCell) {
              return;
            }
            const cell = window.selectedCell;
            const display = cell.querySelector(".grid-cell-display");
            if (!display) {
              return;
            }

            if (force || !cell.hasAttribute("data-original-display-text")) {
              cell.setAttribute("data-original-display-text", display.textContent || "");
            }

            const formulaText = getFormulaPaneSnapshotText(activeEditorInstance);
            const caretInfo = getEditorCursorSnapshot(activeEditorInstance);
            const renderedHtml = renderFormulaTextWithCaret(formulaText, caretInfo);
            cell.setAttribute("data-showing-formula-text", "true");
            display.innerHTML = renderedHtml;
          }

          function clearEditingCellDisplayOverride(cell) {
            if (!cell) {
              return;
            }
            cell.removeAttribute("data-showing-formula-text");

            const display = cell.querySelector(".grid-cell-display");
            if (!display) {
              return;
            }

            const cellRef = cell.getAttribute("data-ref");
            const originalText = cell.getAttribute("data-original-display-text");

            if (cellRef && typeof updateCellDisplay === 'function' && window.hf) {
              updateCellDisplay(cellRef);
            } else if (originalText !== null) {
              display.textContent = originalText;
            }

            cell.removeAttribute("data-original-display-text");
          }

          // Function to update a cell's display from HyperFormula
          function updateCellDisplay(cellRef) {
            if (!window.hf) return;

            const address = cellRefToAddress(cellRef);
            if (!address) return;

            const [row, col] = address;
            const sheetId = 0;

            try {
              // First, check if cell has a formula
              let hasFormula = false;
              let storedFormula = null;
              try {
                storedFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
                if (storedFormula) {
                  hasFormula = true;
                }
              } catch (e) {
                // Cell doesn't have a formula, it's a plain value
                hasFormula = false;
              }

              // Get the calculated value from HyperFormula
              const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });

              // Update the display in the grid
              const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
              if (cell) {
                const isShowingFormulaText = cell.getAttribute("data-showing-formula-text") === "true";
                const display = cell.querySelector(".grid-cell-display");
                if (display) {
                  // Show the calculated result (without "=" prefix) unless we're mirroring the formula text
                  const displayValue = cellValue !== null && cellValue !== undefined ? cellValue.toString() : "";
                  if (!isShowingFormulaText) {
                    display.textContent = displayValue;
                  }

                  // Store the formula or value in data-formula attribute for quick retrieval
                  // Formulas have "=" prefix, plain values don't
                  if (hasFormula && storedFormula) {
                    cell.setAttribute("data-formula", storedFormula);
                  } else if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
                    // Store plain value (without "=" prefix) so selectCell can retrieve it
                    cell.setAttribute("data-formula", cellValue.toString());
                  } else {
                    // Empty cell - clear the attribute
                    cell.removeAttribute("data-formula");
                  }

                  // Determine if value is a number or date (right-align) vs text (left-align)
                  let isNumber = false;
                  if (cellValue !== null && cellValue !== undefined && displayValue !== '') {
                    // Check if it's a JavaScript number type
                    if (typeof cellValue === 'number') {
                      isNumber = true;
                    } else {
                      // Check if string represents a number (including dates which are serial numbers)
                      const numValue = parseFloat(displayValue);
                      if (!isNaN(numValue) && isFinite(numValue) && displayValue.trim() !== '') {
                        // Additional check: if it's a valid number string
                        const trimmed = displayValue.trim();
                        // Match numbers: integers, decimals, scientific notation, but not text that starts with numbers
                        if (/^-?\d+\.?\d*([eE][+-]?\d+)?$/.test(trimmed)) {
                          isNumber = true;
                        }
                      }
                    }
                  }

                  // Apply alignment: right for numbers/dates, left for text
                  if (isNumber) {
                    display.style.textAlign = 'right';
                    cell.classList.add('cell-numeric');
                    cell.classList.remove('cell-text');
                  } else {
                    display.style.textAlign = 'left';
                    cell.classList.add('cell-text');
                    cell.classList.remove('cell-numeric');
                  }
                }
              }
            } catch (error) {
              console.error(`Error updating cell display for ${cellRef}:`, error);
            }
          }

          // Function to get all dependent cells for a given cell address
          function getDependentCells(row, col, sheetId) {
            if (!window.hf) return [];

            const dependentCells = [];

            try {
              // Iterate through all cells that are rendered in the grid
              // This is more efficient than checking all possible cells
              const changedCellRef = addressToCellRef(row, col);

              // Find all cells in the grid that have formulas
              const allCells = gridBody.querySelectorAll('td[data-ref]');

              allCells.forEach(cellElement => {
                const cellRef = cellElement.getAttribute("data-ref");
                if (!cellRef) return;

                // Check if this cell has a stored formula
                const storedFormula = cellElement.getAttribute("data-formula");
                if (storedFormula && storedFormula.startsWith("=")) {
                  // Check if the formula references the changed cell
                  // Build a regex that matches the cell reference (handling $ for absolute references)
                  // Match both A1 and $A$1, $A1, A$1 variants
                  const baseRef = changedCellRef; // e.g., "A1"
                  const colLetter = baseRef.match(/^([A-Z]+)/i)?.[1] || "";
                  const rowNum = baseRef.match(/(\d+)$/)?.[1] || "";

                  // Create regex pattern that matches cell reference in all its forms: A1, $A$1, $A1, A$1
                  // Escape the column letter for regex
                  const escapedColLetter = colLetter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                  const pattern = new RegExp(`\\$?${escapedColLetter}\\$?${rowNum}\\b`, 'i');

                  // Check if pattern matches
                  const referencesChangedCell = pattern.test(storedFormula);

                  if (referencesChangedCell) {
                    const depAddress = cellRefToAddress(cellRef);
                    if (depAddress) {
                      const [depRow, depCol] = depAddress;
                      dependentCells.push({ row: depRow, col: depCol, ref: cellRef });
                    }
                  }
                } else {
                  // If no stored formula, try to get it from HyperFormula
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

          function normalizeValueForCell(rawValue) {
            if (rawValue === null || rawValue === undefined) {
              return "";
            }

            if (typeof rawValue === "number") {
              if (!Number.isFinite(rawValue)) {
                return "";
              }
              return rawValue.toString();
            }

            if (typeof rawValue === "boolean") {
              return rawValue ? "TRUE" : "FALSE";
            }

            if (rawValue instanceof Date) {
              const iso = rawValue.toISOString().split('T')[0];
              const escapedDate = iso.replace(/"/g, '""');
              return `="${escapedDate}"`;
            }

            const stringValue = rawValue.toString();
            const trimmedForFormulaCheck = stringValue.trimStart();
            if (trimmedForFormulaCheck.startsWith("=")) {
              return trimmedForFormulaCheck;
            }

            const escapedValue = stringValue.replace(/"/g, '""');
            return `="${escapedValue}"`;
          }

          function getRawFormulaStoreKey(cellRef, sheetId = getActiveSheetId()) {
            return `${sheetId}:${cellRef}`;
          }

          function saveRawFormula(cellRef, rawFormula, sheetId = getActiveSheetId()) {
            if (!cellRef) {
              return;
            }
            const key = getRawFormulaStoreKey(cellRef, sheetId);
            const hasContent = typeof rawFormula === "string" && rawFormula.trim().length > 0;
            if (hasContent) {
              rawFormulaStore.set(key, rawFormula);
            } else {
              rawFormulaStore.delete(key);
            }
          }

          function readRawFormula(cellRef, sheetId = getActiveSheetId()) {
            if (!cellRef) {
              return "";
            }
            const key = getRawFormulaStoreKey(cellRef, sheetId);
            const stored = rawFormulaStore.get(key);
            return typeof stored === "string" ? stored : "";
          }

          // Function to set cell value in HyperFormula and update display
          function setCellValue(cellRef, value, options = {}) {
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
              // First sheet is always at index 0
              const sheetId = 0;

              // Strip comments from formula before storing and executing
              // Comments are already stripped in Ctrl+Enter handler, but this ensures
              // comments are removed even if setCellValue is called directly
            const formulaForHyperFormula = stripComments(value);
            console.log(`[HyperFormula Formula] ${formulaForHyperFormula || ''}`);

              // Set the value in HyperFormula - address object uses col, row, sheet order
            const hfResponse = window.hf.setCellContents({ col, row, sheet: sheetId }, [[formulaForHyperFormula]]);
              console.log('[HyperFormula Response]', { cellRef, response: hfResponse });

              // Recalculate to execute formulas - HyperFormula automatically recalculates dependencies
              window.hf.rebuildAndRecalculate();

              try {
                const evaluatedValue = window.hf.getCellValue({ col, row, sheet: sheetId });
                console.log('[HyperFormula Value]', { cellRef, value: evaluatedValue });
              } catch (valueError) {
                console.warn('[HyperFormula Value Error]', { cellRef, error: valueError });
              }

              const rawFormulaToPersist = options.rawFormula !== undefined ? options.rawFormula : value;
              if (rawFormulaToPersist !== undefined) {
                saveRawFormula(cellRef, rawFormulaToPersist, sheetId);
              } else if (!value) {
                saveRawFormula(cellRef, "", sheetId);
              }

              // Update the changed cell's display
              updateCellDisplay(cellRef);

              // Store the formula text (with "=" prefix) in data attribute for later retrieval
              const cell = gridBody.querySelector(`td[data-ref="${cellRef}"]`);
              if (cell) {
                cell.setAttribute("data-formula", formulaForHyperFormula);
              }

              // Find and update all dependent cells
              const dependentCells = getDependentCells(row, col, sheetId);

              // Update displays of all dependent cells
              dependentCells.forEach(depCell => {
                updateCellDisplay(depCell.ref);

                // Also update the formula stored in data attribute if it exists
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

          // Function to get cell value from HyperFormula
          function getCellValue(cellRef) {
            if (!window.hf) {
              return null;
            }

            const address = cellRefToAddress(cellRef);
            if (!address) {
              return null;
            }

            const [row, col] = address;
            try {
              // First sheet is always at index 0
              const sheetId = 0;

              const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
              return cellValue;
            } catch (error) {
              console.error("Error getting cell value:", error);
              return null;
            }
          }

          function getSelectionBounds() {
            if (!selectedCells || selectedCells.size === 0) {
              return null;
            }
            let minRow = Number.POSITIVE_INFINITY;
            let maxRow = Number.NEGATIVE_INFINITY;
            let minCol = Number.POSITIVE_INFINITY;
            let maxCol = Number.NEGATIVE_INFINITY;
            selectedCells.forEach((cell) => {
              const rowAttr = cell.getAttribute("data-row");
              const colAttr = cell.getAttribute("data-col");
              if (rowAttr === null || colAttr === null) {
                return;
              }
              const row = parseInt(rowAttr, 10);
              const col = parseInt(colAttr, 10);
              if (Number.isNaN(row) || Number.isNaN(col)) {
                return;
              }
              minRow = Math.min(minRow, row);
              maxRow = Math.max(maxRow, row);
              minCol = Math.min(minCol, col);
              maxCol = Math.max(maxCol, col);
            });
            if (!Number.isFinite(minRow) || !Number.isFinite(minCol)) {
              return null;
            }
            return {
              minRow,
              maxRow,
              minCol,
              maxCol,
              height: maxRow - minRow + 1,
              width: maxCol - minCol + 1
            };
          }

          function clearSelectedCellContents() {
            if (!selectedCells || selectedCells.size === 0) {
              return false;
            }
            let clearedAny = false;
            selectedCells.forEach((cell) => {
              const cellRef = cell.getAttribute("data-ref");
              if (cellRef) {
                setCellValue(cellRef, "", { rawFormula: "" });
                cell.removeAttribute("data-formula");
                clearedAny = true;
              }
            });
            if (clearedAny && typeof updateCellChips === "function") {
              requestAnimationFrame(() => updateCellChips());
            }
            return clearedAny;
          }

          function wrapIndex(value, size) {
            if (!size || size <= 0) {
              return 0;
            }
            const remainder = value % size;
            return remainder < 0 ? remainder + size : remainder;
          }

          function clearFillPreview() {
            if (!fillPreviewCells) {
              return;
            }
            fillPreviewCells.forEach((cell) => cell.classList.remove("fill-preview"));
            fillPreviewCells.clear();
          }

          function updateFillPreview(range) {
            clearFillPreview();
            if (!range) {
              return;
            }
            const minRow = Math.min(range.minRow ?? range.maxRow ?? 0, range.maxRow ?? range.minRow ?? 0);
            const maxRow = Math.max(range.minRow ?? range.maxRow ?? 0, range.maxRow ?? range.minRow ?? 0);
            const minCol = Math.min(range.minCol ?? range.maxCol ?? 0, range.maxCol ?? range.minCol ?? 0);
            const maxCol = Math.max(range.minCol ?? range.maxCol ?? 0, range.maxCol ?? range.minCol ?? 0);
            for (let row = minRow; row <= maxRow; row++) {
              for (let col = minCol; col <= maxCol; col++) {
                const cell = gridBody.querySelector(`tr[data-row="${row}"] td[data-col="${col}"]`);
                if (cell) {
                  cell.classList.add("fill-preview");
                  fillPreviewCells.add(cell);
                }
              }
            }
          }

          function getCellDataSnapshot(row, col) {
            const cell = gridBody.querySelector(`tr[data-row="${row}"] td[data-col="${col}"]`);
            if (!cell) {
              return { raw: "", isFormula: false };
            }
            let rawValue = cell.getAttribute("data-formula");
            if (rawValue === null || rawValue === undefined) {
              const cellRef = cell.getAttribute("data-ref");
              if (cellRef) {
                const currentValue = getCellValue(cellRef);
                rawValue = currentValue !== null && currentValue !== undefined ? currentValue.toString() : "";
              } else {
                rawValue = "";
              }
            }
            let normalized = typeof rawValue === "string" ? rawValue : (rawValue ?? "").toString();
            if (!normalized) {
              const cellRef = cell.getAttribute("data-ref");
              if (cellRef && window.hf) {
                const address = cellRefToAddress(cellRef);
                if (address) {
                  const [cellRow, cellCol] = address;
                  try {
                    const sheetId = typeof window.hfSheetId === "number" ? window.hfSheetId : 0;
                    const hfFormula = window.hf.getCellFormula({ col: cellCol, row: cellRow, sheet: sheetId });
                    if (hfFormula) {
                      normalized = hfFormula;
                    }
                  } catch (formulaError) {
                    // Ignore if formula lookup fails
                  }
                }
              }
            }
            let isFormulaCell = typeof normalized === "string" && normalized.startsWith("=");
            return {
              raw: normalized,
              isFormula: isFormulaCell
            };
          }

          function determineFillTarget(bounds, pointerRow, pointerCol, lockedAxis) {
            if (!bounds) {
              return null;
            }
            let axis = lockedAxis || null;
            if (axis === null) {
              if (pointerRow < bounds.minRow || pointerRow > bounds.maxRow) {
                axis = "vertical";
              } else if (pointerCol < bounds.minCol || pointerCol > bounds.maxCol) {
                axis = "horizontal";
              }
            }
            if (!axis) {
              return null;
            }
            if (axis === "vertical") {
              if (pointerRow > bounds.maxRow) {
                return {
                  axis: "vertical",
                  direction: "down",
                  minRow: bounds.maxRow + 1,
                  maxRow: pointerRow,
                  minCol: bounds.minCol,
                  maxCol: bounds.maxCol
                };
              }
              if (pointerRow < bounds.minRow) {
                return {
                  axis: "vertical",
                  direction: "up",
                  minRow: pointerRow,
                  maxRow: bounds.minRow - 1,
                  minCol: bounds.minCol,
                  maxCol: bounds.maxCol
                };
              }
            } else if (axis === "horizontal") {
              if (pointerCol > bounds.maxCol) {
                return {
                  axis: "horizontal",
                  direction: "right",
                  minCol: bounds.maxCol + 1,
                  maxCol: pointerCol,
                  minRow: bounds.minRow,
                  maxRow: bounds.maxRow
                };
              }
              if (pointerCol < bounds.minCol) {
                return {
                  axis: "horizontal",
                  direction: "left",
                  minCol: pointerCol,
                  maxCol: bounds.minCol - 1,
                  minRow: bounds.minRow,
                  maxRow: bounds.maxRow
                };
              }
            }
            return null;
          }

          function buildRange(start, end, step) {
            const values = [];
            if (step > 0) {
              for (let value = start; value <= end; value += step) {
                values.push(value);
              }
            } else if (step < 0) {
              for (let value = start; value >= end; value += step) {
                values.push(value);
              }
            }
            return values;
          }

          function applyAutofillFromSelection(bounds, targetRange) {
            if (!bounds || !targetRange) {
              return;
            }
            if (targetRange.axis === "vertical") {
              applyVerticalAutofill(bounds, targetRange);
            } else if (targetRange.axis === "horizontal") {
              applyHorizontalAutofill(bounds, targetRange);
            }
          }

          function applyVerticalAutofill(bounds, targetRange) {
            const targetRows = targetRange.direction === "down"
              ? buildRange(targetRange.minRow, targetRange.maxRow, 1)
              : buildRange(targetRange.maxRow, targetRange.minRow, -1);
            if (targetRows.length === 0) {
              return;
            }
            const fillDirection = targetRange.direction === "down"
              ? AutofillDirection.FORWARD
              : AutofillDirection.BACKWARD;
            for (let col = bounds.minCol; col <= bounds.maxCol; col++) {
              const seeds = [];
              for (let row = bounds.minRow; row <= bounds.maxRow; row++) {
                seeds.push(getCellDataSnapshot(row, col));
              }
              const allValues = seeds.every((entry) => !entry.isFormula);
              if (allValues) {
                const rawSeeds = seeds.map((entry) => entry.raw ?? "");
                const fillValues = generateValueFill(rawSeeds, targetRows.length, { direction: fillDirection });
                targetRows.forEach((targetRow, index) => {
                  const valueToSet = fillValues[index] ?? "";
                  const targetRef = addressToCellRef(targetRow, col);
                  setCellValue(targetRef, valueToSet, { rawFormula: valueToSet });
                });
              } else {
                targetRows.forEach((targetRow) => {
                  const relativeIndex = wrapIndex(targetRow - bounds.minRow, seeds.length);
                  const sourceRow = bounds.minRow + relativeIndex;
                  const seed = seeds[relativeIndex] || { raw: "", isFormula: false };
                  let nextValue = seed.raw ?? "";
                  if (seed.isFormula) {
                    const rowDelta = targetRow - sourceRow;
                    nextValue = adjustFormulaReferences(nextValue, rowDelta, 0);
                  }
                  const targetRef = addressToCellRef(targetRow, col);
                  setCellValue(targetRef, nextValue, { rawFormula: nextValue });
                });
              }
            }
            const newMinRow = Math.min(bounds.minRow, targetRange.minRow);
            const newMaxRow = Math.max(bounds.maxRow, targetRange.maxRow);
            const startCell = gridBody.querySelector(`tr[data-row="${newMinRow}"] td[data-col="${bounds.minCol}"]`);
            const endCell = gridBody.querySelector(`tr[data-row="${newMaxRow}"] td[data-col="${bounds.maxCol}"]`);
            if (startCell && endCell) {
              selectRange(startCell, endCell);
            }
          }

          function applyHorizontalAutofill(bounds, targetRange) {
            const targetCols = targetRange.direction === "right"
              ? buildRange(targetRange.minCol, targetRange.maxCol, 1)
              : buildRange(targetRange.maxCol, targetRange.minCol, -1);
            if (targetCols.length === 0) {
              return;
            }
            const fillDirection = targetRange.direction === "right"
              ? AutofillDirection.FORWARD
              : AutofillDirection.BACKWARD;
            for (let row = bounds.minRow; row <= bounds.maxRow; row++) {
              const seeds = [];
              for (let col = bounds.minCol; col <= bounds.maxCol; col++) {
                seeds.push(getCellDataSnapshot(row, col));
              }
              const allValues = seeds.every((entry) => !entry.isFormula);
              if (allValues) {
                const rawSeeds = seeds.map((entry) => entry.raw ?? "");
                const fillValues = generateValueFill(rawSeeds, targetCols.length, { direction: fillDirection });
                targetCols.forEach((targetCol, index) => {
                  const valueToSet = fillValues[index] ?? "";
                  const targetRef = addressToCellRef(row, targetCol);
                  setCellValue(targetRef, valueToSet, { rawFormula: valueToSet });
                });
              } else {
                targetCols.forEach((targetCol) => {
                  const relativeIndex = wrapIndex(targetCol - bounds.minCol, seeds.length);
                  const sourceCol = bounds.minCol + relativeIndex;
                  const seed = seeds[relativeIndex] || { raw: "", isFormula: false };
                  let nextValue = seed.raw ?? "";
                  if (seed.isFormula) {
                    const colDelta = targetCol - sourceCol;
                    nextValue = adjustFormulaReferences(nextValue, 0, colDelta);
                  }
                  const targetRef = addressToCellRef(row, targetCol);
                  setCellValue(targetRef, nextValue, { rawFormula: nextValue });
                });
              }
            }
            const newMinCol = Math.min(bounds.minCol, targetRange.minCol);
            const newMaxCol = Math.max(bounds.maxCol, targetRange.maxCol);
            const startCell = gridBody.querySelector(`tr[data-row="${bounds.minRow}"] td[data-col="${newMinCol}"]`);
            const endCell = gridBody.querySelector(`tr[data-row="${bounds.maxRow}"] td[data-col="${newMaxCol}"]`);
            if (startCell && endCell) {
              selectRange(startCell, endCell);
            }
          }

          function handleFillHandleMouseDown(event) {
            if (event.button !== 0) {
              return;
            }
            if (isEditMode) {
              return;
            }
            const bounds = getSelectionBounds();
            if (!bounds) {
              return;
            }
            event.preventDefault();
            event.stopPropagation();
            fillHandleDragState = {
              bounds,
              axis: null,
              targetRange: null
            };
            document.addEventListener("mousemove", handleFillHandleMouseMove);
            document.addEventListener("mouseup", handleFillHandleMouseUp);
          }

          function handleFillHandleMouseMove(event) {
            if (!fillHandleDragState) {
              return;
            }
            const cell = getCellElementFromPoint(event.clientX, event.clientY);
            if (!cell) {
              fillHandleDragState.targetRange = null;
              updateFillPreview(null);
              return;
            }
            const row = parseInt(cell.getAttribute("data-row"), 10);
            const col = parseInt(cell.getAttribute("data-col"), 10);
            if (Number.isNaN(row) || Number.isNaN(col)) {
              fillHandleDragState.targetRange = null;
              updateFillPreview(null);
              return;
            }
            const nextRange = determineFillTarget(fillHandleDragState.bounds, row, col, fillHandleDragState.axis);
            if (nextRange && !fillHandleDragState.axis) {
              fillHandleDragState.axis = nextRange.axis;
            }
            fillHandleDragState.targetRange = nextRange;
            updateFillPreview(nextRange);
          }

          function handleFillHandleMouseUp() {
            if (!fillHandleDragState) {
              return;
            }
            document.removeEventListener("mousemove", handleFillHandleMouseMove);
            document.removeEventListener("mouseup", handleFillHandleMouseUp);
            const { bounds, targetRange } = fillHandleDragState;
            fillHandleDragState = null;
            clearFillPreview();
            if (targetRange) {
              applyAutofillFromSelection(bounds, targetRange);
            }
          }

          function getCellElementFromPoint(x, y) {
            const element = document.elementFromPoint(x, y);
            if (!element) {
              return null;
            }
            const cell = element.closest('td[data-row][data-col]');
            if (!cell || !gridBody.contains(cell)) {
              return null;
            }
            return cell;
          }

          if (fillHandleElement) {
            fillHandleElement.addEventListener("mousedown", handleFillHandleMouseDown);
          }

          function populateGridWithData(dataRows, startRef = "A1") {
            if (!window.hf) {
              console.warn("HyperFormula not initialized - skipping data population");
              return;
            }
            if (!Array.isArray(dataRows) || dataRows.length === 0) {
              return;
            }

            const startAddress = cellRefToAddress(startRef);
            if (!startAddress) {
              console.warn("Invalid start reference for data population:", startRef);
              return;
            }

            const [startRow, startCol] = startAddress;

            dataRows.forEach((rowValues, rowIndex) => {
              if (!Array.isArray(rowValues)) {
                return;
              }

              rowValues.forEach((rawValue, colIndex) => {
                const targetRef = addressToCellRef(startRow + rowIndex, startCol + colIndex);
                const normalizedValue = normalizeValueForCell(rawValue);
                setCellValue(targetRef, normalizedValue, { rawFormula: normalizedValue });
              });
            });
          }

          // Function to get cell reference from row and col (0-indexed)
          function getCellReference(row, col) {
            return addressToCellRef(row, col);
          }

          // Function to update the cell range pill
          function updateCellRangePill() {
            if (selectedCells.size === 0) {
              currentSelectionLabel = "";
              updateModeIndicatorElements();
              return;
            }

            const cellsArray = Array.from(selectedCells);
            const rows = cellsArray.map(cell => parseInt(cell.getAttribute("data-row"))).filter(r => !isNaN(r));
            const cols = cellsArray.map(cell => parseInt(cell.getAttribute("data-col"))).filter(c => !isNaN(c));

            if (rows.length === 0 || cols.length === 0) {
              currentSelectionLabel = "";
              updateModeIndicatorElements();
              return;
            }

            const minRow = Math.min(...rows);
            const maxRow = Math.max(...rows);
            const minCol = Math.min(...cols);
            const maxCol = Math.max(...cols);

            if (minRow === maxRow && minCol === maxCol) {
              currentSelectionLabel = getCellReference(minRow, minCol);
            } else {
              currentSelectionLabel = `${getCellReference(minRow, minCol)}:${getCellReference(maxRow, maxCol)}`;
            }
            updateModeIndicatorElements();
          }

          // Function to clear all selections (without updating pill)
          function clearSelection() {
            selectedCells.forEach(cell => {
              cell.classList.remove("selected");
              cell.classList.remove("edit-mode"); // Remove edit-mode class when clearing selection
              cell.removeAttribute("data-selection-edge");
            });
            selectedCells.clear();
            selectionAnchorCell = null;
            clearFillPreview();

            // Remove edit-mode class from editor-wrapper when clearing selection
            if (currentMode !== MODES.EDIT) {
              const editorWrapper = document.querySelector('.editor-wrapper');
              if (editorWrapper) {
                editorWrapper.classList.remove('editor-edit-mode');
              }
            }

            // Hide selection overlay and fill handle
            if (selectionOverlayElement) {
              selectionOverlayElement.style.display = "none";
            }
            if (fillHandleElement) {
              fillHandleElement.style.display = "none";
            }
            clearHeaderHighlights();
          }

          // Function to update the selection overlay rectangle
          function updateSelectionOverlay() {
            const overlay = selectionOverlayElement;
            if (!overlay) {
              return;
            }

            if (isEditMode) {
              if (editModeRangeHighlight.active && editModeRangeHighlight.bounds) {
                positionElementForBounds(overlay, editModeRangeHighlight.bounds);
              } else {
                overlay.style.display = "none";
              }
              if (fillHandleElement) {
                fillHandleElement.style.display = "none";
              }
              return;
            }

            if (selectedCells.size === 0) {
              overlay.style.display = "none";
              if (fillHandleElement) {
                fillHandleElement.style.display = "none";
              }
              return;
            }

            // Get all selected cells
            const cells = Array.from(selectedCells);
            if (cells.length === 0) {
              overlay.style.display = "none";
              return;
            }

            // Get the grid wrapper inner (scrollable container)
            if (!gridWrapperInner) return;

            // Get bounding rectangles
            const gridRect = gridWrapperInner.getBoundingClientRect();

            // Find the min/max positions of all selected cells
            let minLeft = Infinity;
            let minTop = Infinity;
            let maxRight = -Infinity;
            let maxBottom = -Infinity;

            cells.forEach(cell => {
              const cellRect = cell.getBoundingClientRect();
              const relativeLeft = cellRect.left - gridRect.left + gridWrapperInner.scrollLeft;
              const relativeTop = cellRect.top - gridRect.top + gridWrapperInner.scrollTop;

              minLeft = Math.min(minLeft, relativeLeft);
              minTop = Math.min(minTop, relativeTop);
              maxRight = Math.max(maxRight, relativeLeft + cellRect.width);
              maxBottom = Math.max(maxBottom, relativeTop + cellRect.height);
            });

            // Position and size the overlay
            overlay.style.left = `${minLeft}px`;
            overlay.style.top = `${minTop}px`;
            overlay.style.width = `${maxRight - minLeft}px`;
            overlay.style.height = `${maxBottom - minTop}px`;
            overlay.style.display = "block";

            // Show fill handle
            if (fillHandleElement) {
              fillHandleElement.style.display = "block";
            }
          }

          window.updateSelectionOverlay = updateSelectionOverlay;

          function getSelectionBoundsInfo() {
            if (!selectedCells || selectedCells.size === 0) {
              return null;
            }
            const rows = [];
            const cols = [];
            selectedCells.forEach(cell => {
              const rowAttr = cell.getAttribute("data-row");
              const colAttr = cell.getAttribute("data-col");
              const rowValue = rowAttr !== null ? parseInt(rowAttr, 10) : NaN;
              const colValue = colAttr !== null ? parseInt(colAttr, 10) : NaN;
              if (!Number.isNaN(rowValue)) {
                rows.push(rowValue);
              }
              if (!Number.isNaN(colValue)) {
                cols.push(colValue);
              }
            });
            if (rows.length === 0 || cols.length === 0) {
              return null;
            }
            const minRow = Math.min(...rows);
            const maxRow = Math.max(...rows);
            const minCol = Math.min(...cols);
            const maxCol = Math.max(...cols);
            return {
              minRow,
              maxRow,
              minCol,
              maxCol,
              rowCount: maxRow - minRow + 1,
              colCount: maxCol - minCol + 1
            };
          }

          function getCellFromViewportPoint(clientX, clientY) {
            const element = document.elementFromPoint(clientX, clientY);
            if (!element) {
              return null;
            }
            return element.closest('#gridBody td[data-row]');
          }

          function positionElementForBounds(element, bounds) {
            if (!element || !bounds || !gridWrapperInner) {
              return;
            }
            const topLeftCell = gridBody.querySelector(`tr[data-row="${bounds.minRow}"] td[data-col="${bounds.minCol}"]`);
            const bottomRightCell = gridBody.querySelector(`tr[data-row="${bounds.maxRow}"] td[data-col="${bounds.maxCol}"]`);
            if (!topLeftCell || !bottomRightCell) {
              element.style.display = "none";
              return;
            }
            const gridRect = gridWrapperInner.getBoundingClientRect();
            const topLeftRect = topLeftCell.getBoundingClientRect();
            const bottomRightRect = bottomRightCell.getBoundingClientRect();
            const left = topLeftRect.left - gridRect.left + gridWrapperInner.scrollLeft;
            const top = topLeftRect.top - gridRect.top + gridWrapperInner.scrollTop;
            const width = bottomRightRect.right - topLeftRect.left;
            const height = bottomRightRect.bottom - topLeftRect.top;
            element.style.left = `${left}px`;
            element.style.top = `${top}px`;
            element.style.width = `${width}px`;
            element.style.height = `${height}px`;
            element.style.display = "block";
          }

          function showSelectionDragPreview(bounds) {
            if (!selectionDragPreview) {
              return;
            }
            positionElementForBounds(selectionDragPreview, bounds);
          }

          function hideSelectionDragPreview() {
            if (selectionDragPreview) {
              selectionDragPreview.style.display = "none";
            }
          }

          function startSelectionDrag(event) {
            if (selectionDragState) {
              return;
            }
            const bounds = getSelectionBoundsInfo();
            if (!bounds) {
              return;
            }
            selectionDragState = {
              bounds,
              rowOffset: 0,
              colOffset: 0,
              dropTarget: null,
              overlayPointerEvents: selectionOverlayElement ? selectionOverlayElement.style.pointerEvents : ""
            };
            if (selectionOverlayElement) {
              selectionOverlayElement.style.pointerEvents = "none";
            }
            const startCell = getCellFromViewportPoint(event.clientX, event.clientY);
            if (startCell) {
              const startRow = parseInt(startCell.getAttribute("data-row"), 10);
              const startCol = parseInt(startCell.getAttribute("data-col"), 10);
              if (!Number.isNaN(startRow) && !Number.isNaN(startCol)) {
                selectionDragState.rowOffset = startRow - bounds.minRow;
                selectionDragState.colOffset = startCol - bounds.minCol;
              }
            }
            document.body.style.cursor = "grabbing";
            showSelectionDragPreview(bounds);
            document.addEventListener("mousemove", handleSelectionDragMove);
            document.addEventListener("mouseup", handleSelectionDragEnd);
            event.preventDefault();
            event.stopPropagation();
          }

          function handleSelectionDragMove(event) {
            if (!selectionDragState) {
              return;
            }
            const hoverCell = getCellFromViewportPoint(event.clientX, event.clientY);
            if (!hoverCell) {
              selectionDragState.dropTarget = null;
              hideSelectionDragPreview();
              return;
            }
            const hoverRow = parseInt(hoverCell.getAttribute("data-row"), 10);
            const hoverCol = parseInt(hoverCell.getAttribute("data-col"), 10);
            if (Number.isNaN(hoverRow) || Number.isNaN(hoverCol)) {
              selectionDragState.dropTarget = null;
              hideSelectionDragPreview();
              return;
            }
            const state = selectionDragState;
            const rowCount = state.bounds.rowCount;
            const colCount = state.bounds.colCount;
            const maxRowStart = typeof window.totalGridRows === 'number' ? Math.max(0, window.totalGridRows - rowCount) : null;
            const maxColStart = typeof window.totalGridColumns === 'number' ? Math.max(0, window.totalGridColumns - colCount) : null;
            let targetRow = hoverRow - state.rowOffset;
            let targetCol = hoverCol - state.colOffset;
            if (maxRowStart !== null && !Number.isNaN(maxRowStart)) {
              targetRow = Math.min(Math.max(0, targetRow), maxRowStart);
            } else {
              targetRow = Math.max(0, targetRow);
            }
            if (maxColStart !== null && !Number.isNaN(maxColStart)) {
              targetCol = Math.min(Math.max(0, targetCol), maxColStart);
            } else {
              targetCol = Math.max(0, targetCol);
            }
            state.dropTarget = { row: targetRow, col: targetCol };
            const previewBounds = {
              minRow: targetRow,
              maxRow: targetRow + rowCount - 1,
              minCol: targetCol,
              maxCol: targetCol + colCount - 1,
              rowCount,
              colCount
            };
            showSelectionDragPreview(previewBounds);
          }

          function handleSelectionDragEnd() {
            if (!selectionDragState) {
              return;
            }
            document.removeEventListener("mousemove", handleSelectionDragMove);
            document.removeEventListener("mouseup", handleSelectionDragEnd);
            document.body.style.cursor = "";
            if (selectionOverlayElement) {
              selectionOverlayElement.style.pointerEvents = selectionDragState.overlayPointerEvents || "";
            }
            hideSelectionDragPreview();
            const { bounds, dropTarget } = selectionDragState;
            selectionDragState = null;
            if (!dropTarget) {
              return;
            }
            if (dropTarget.row === bounds.minRow && dropTarget.col === bounds.minCol) {
              return;
            }
            applySelectionMove(bounds, dropTarget);
          }

          function applySelectionMove(sourceBounds, targetStart) {
            if (!window.hf) {
              console.warn("HyperFormula not initialized");
              return;
            }
            const sheetId = window.hfSheetId !== undefined ? window.hfSheetId : 0;
            const sourceRange = {
              start: { sheet: sheetId, row: sourceBounds.minRow, col: sourceBounds.minCol },
              end: { sheet: sheetId, row: sourceBounds.maxRow, col: sourceBounds.maxCol }
            };
            const destinationAddress = { sheet: sheetId, row: targetStart.row, col: targetStart.col };
            try {
              const isAllowed = typeof window.hf.isItPossibleToMoveCells === 'function'
                ? window.hf.isItPossibleToMoveCells(sourceRange, destinationAddress)
                : true;
              if (!isAllowed) {
                console.warn("HyperFormula reported moveCells operation is not allowed for the requested range.");
                return;
              }
              window.hf.moveCells(sourceRange, destinationAddress);
            } catch (error) {
              console.error("Failed to move selection via HyperFormula:", error);
              return;
            }
            const targetBounds = {
              minRow: targetStart.row,
              maxRow: targetStart.row + sourceBounds.rowCount - 1,
              minCol: targetStart.col,
              maxCol: targetStart.col + sourceBounds.colCount - 1,
              rowCount: sourceBounds.rowCount,
              colCount: sourceBounds.colCount
            };
            updateDisplaysForRanges([sourceBounds, targetBounds]);
            const startCell = gridBody.querySelector(`tr[data-row="${targetBounds.minRow}"] td[data-col="${targetBounds.minCol}"]`);
            const endCell = gridBody.querySelector(`tr[data-row="${targetBounds.maxRow}"] td[data-col="${targetBounds.maxCol}"]`);
            if (startCell && endCell) {
              selectRange(startCell, endCell);
            } else {
              updateSelectionOverlay();
              updateCellRangePill();
            }
          }

          function updateDisplaysForRanges(boundsList) {
            if (!Array.isArray(boundsList) || boundsList.length === 0) {
              return;
            }
            const refsToUpdate = new Set();
            const sheetId = window.hfSheetId !== undefined ? window.hfSheetId : 0;
            boundsList.forEach(bounds => {
              if (!bounds) {
                return;
              }
              for (let row = bounds.minRow; row <= bounds.maxRow; row++) {
                for (let col = bounds.minCol; col <= bounds.maxCol; col++) {
                  const ref = addressToCellRef(row, col);
                  if (ref) {
                    refsToUpdate.add(ref);
                  }
                  const dependents = getDependentCells(row, col, sheetId);
                  if (Array.isArray(dependents)) {
                    dependents.forEach(dep => {
                      if (dep && dep.ref) {
                        refsToUpdate.add(dep.ref);
                      }
                    });
                  }
                }
              }
            });
            refsToUpdate.forEach(ref => updateCellDisplay(ref));
          }

          function refreshCellsFromChanges(changes) {
            if (!Array.isArray(changes) || changes.length === 0) {
              return;
            }
            const refsToUpdate = new Set();
            changes.forEach(change => {
              if (!change || !change.address) {
                return;
              }
              const { row, col } = change.address;
              const ref = addressToCellRef(row, col);
              if (ref) {
                refsToUpdate.add(ref);
              }
            });
            refsToUpdate.forEach(ref => updateCellDisplay(ref));
          }

          function refreshFormulaEditorFromSelection() {
            if (!window.selectedCell) {
              return;
            }
            const currentEditor = window.monacoEditor || editor;
            if (!currentEditor || typeof currentEditor.setValue !== 'function') {
              return;
            }
            const cell = window.selectedCell;
            const isHeader = cell.tagName === "TH" || cell === cell.parentElement.querySelector("td:first-child");
            let editorValue = "";
            if (!isHeader) {
              editorValue = getStoredFormulaText(cell) || "";
              if (!editorValue && window.hf) {
                const cellRef = cell.getAttribute("data-ref");
                const address = cellRef ? cellRefToAddress(cellRef) : null;
                if (address) {
                  const [row, col] = address;
                  const sheetId = window.hfSheetId !== undefined ? window.hfSheetId : 0;
                  try {
                    const cellFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
                    if (cellFormula) {
                      editorValue = cellFormula.startsWith("=") ? cellFormula.substring(1) : cellFormula;
                      cell.setAttribute("data-formula", cellFormula);
                      if (cellRef && !readRawFormula(cellRef, sheetId)) {
                        saveRawFormula(cellRef, cellFormula, sheetId);
                      }
                    } else {
                      const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
                      if (cellValue !== null && cellValue !== undefined) {
                        editorValue = cellValue.toString();
                        cell.setAttribute("data-formula", editorValue);
                        if (cellRef && !readRawFormula(cellRef, sheetId)) {
                          saveRawFormula(cellRef, editorValue, sheetId);
                        }
                      } else {
                        cell.removeAttribute("data-formula");
                      }
                    }
                  } catch (error) {
                    console.error("Failed to refresh editor snapshot:", error);
                  }
                }
              }
            }
            const fullValue = ensureSevenLines(editorValue || "");
            const firstLine = (editorValue || "").split('\n')[0] || "";
            const targetColumn = Math.max(1, firstLine.length + 1);
            window.isProgrammaticCursorChange = true;
            currentEditor.setValue(fullValue);
            currentEditor.setPosition({ lineNumber: 1, column: targetColumn });
            Promise.resolve().then(() => {
              window.isProgrammaticCursorChange = false;
              if (typeof updateCellChips === 'function') {
                requestAnimationFrame(() => updateCellChips());
              }
            });
            if (currentMode !== MODES.EDIT) {
              currentEditor.updateOptions({ cursorBlinking: 'hidden' });
            }
          }

          function handleUndoRedoAction(action) {
            if (!window.hf) {
              return;
            }
            const canUndo = action === 'undo'
              ? (typeof window.hf.isThereSomethingToUndo === 'function' ? window.hf.isThereSomethingToUndo() : true)
              : (typeof window.hf.isThereSomethingToRedo === 'function' ? window.hf.isThereSomethingToRedo() : true);
            if (!canUndo) {
              return;
            }
            try {
              const changes = action === 'undo' ? window.hf.undo() : window.hf.redo();
              refreshCellsFromChanges(changes);
              if (selectedCells && selectedCells.size > 0) {
                updateSelectionOverlay();
                updateHeaderHighlightsFromSelection();
                refreshFormulaEditorFromSelection();
              } else {
                updateSelectionOverlay();
                updateCellRangePill();
              }
              if (typeof updateCellChips === 'function') {
                requestAnimationFrame(() => updateCellChips());
              }
            } catch (error) {
              console.error(`Failed to ${action}:`, error);
            }
          }

          // Function to select a range of cells
          function selectRange(startCell, endCell) {
            clearSelection();

            if (!startCell || !endCell) return;
            selectionAnchorCell = startCell;

            const startRow = parseInt(startCell.getAttribute("data-row"));
            const startCol = parseInt(startCell.getAttribute("data-col"));
            const endRow = parseInt(endCell.getAttribute("data-row"));
            const endCol = parseInt(endCell.getAttribute("data-col"));

            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);

            // Select all cells in the range (skip headers)
            for (let r = minRow; r <= maxRow; r++) {
              for (let c = minCol; c <= maxCol; c++) {
                const cell = gridBody.querySelector(`tr[data-row="${r}"] td[data-col="${c}"]`);
                if (cell) {
                  // Skip row headers (first column)
                  const isRowHeader = cell === cell.parentElement.querySelector("td:first-child");
                  if (!isRowHeader) {
                    cell.classList.add("selected");
                    selectedCells.add(cell);

                    // Mark edge cells for border styling
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

            // Update selection overlay
            updateSelectionOverlay();
            updateHeaderHighlightsFromSelection();

            // Update formula editor with first cell's value
            const firstCell = gridBody.querySelector(`tr[data-row="${minRow}"] td[data-col="${minCol}"]`);
            if (firstCell) {
              const isHeader = firstCell.tagName === "TH" || firstCell === firstCell.parentElement.querySelector("td:first-child");
              if (!isHeader) {
                let editorValue = getStoredFormulaText(firstCell) || "";
                if (!editorValue) {
                  const cellRefForRange = firstCell.getAttribute("data-ref");
                  if (cellRefForRange && window.hf) {
                    const address = cellRefToAddress(cellRefForRange);
                    if (address) {
                      const [row, col] = address;
                      const sheetId = typeof window.hfSheetId === "number" ? window.hfSheetId : 0;
                      try {
                        const cellFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
                        if (cellFormula) {
                          editorValue = cellFormula.startsWith("=") ? cellFormula.substring(1) : cellFormula;
                          const formulaToStore = cellFormula.startsWith("=") ? cellFormula : `=${cellFormula}`;
                          firstCell.setAttribute("data-formula", formulaToStore);
                          if (cellRefForRange && !readRawFormula(cellRefForRange, sheetId)) {
                            saveRawFormula(cellRefForRange, formulaToStore, sheetId);
                          }
                        } else {
                          const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
                          if (cellValue !== null && cellValue !== undefined) {
                            editorValue = cellValue.toString();
                            firstCell.setAttribute("data-formula", editorValue);
                            if (cellRefForRange && !readRawFormula(cellRefForRange, sheetId)) {
                              saveRawFormula(cellRefForRange, editorValue, sheetId);
                            }
                          } else {
                            firstCell.removeAttribute("data-formula");
                          }
                        }
                      } catch (rangeError) {
                        console.error("Failed to hydrate formula for selection range:", rangeError);
                      }
                    }
                  }
                }

                const currentEditor = window.monacoEditor || editor;
                if (currentEditor) {
                  // Formula starts on line 1
                  const fullValue = ensureSevenLines(editorValue || "");

                  // Calculate target cursor position before setting value
                  const line1Content = (editorValue || "").split('\n')[0] || "";
                  const targetColumn = Math.max(1, line1Content.length + 1);

                  // Remove edit-mode class from editor-wrapper to show 30% opacity when just selecting
                  const editorWrapper = document.querySelector('.editor-wrapper');
                  if (editorWrapper && currentMode !== MODES.EDIT) {
                    editorWrapper.classList.remove('editor-edit-mode');
                  }

                  // Set programmatic flag BEFORE setValue to prevent event listener from interfering
                  window.isProgrammaticCursorChange = true;

                  // Set value
                  currentEditor.setValue(fullValue);

                  // Set cursor to line 1 immediately and synchronously
                  currentEditor.setPosition({ lineNumber: 1, column: targetColumn });

                  // Reset flag after a microtask to allow normal cursor behavior
                  Promise.resolve().then(() => {
                    window.isProgrammaticCursorChange = false;
                    if (typeof updateCellChips === 'function') {
                      requestAnimationFrame(() => updateCellChips());
                    }
                  });

                  currentEditor.updateOptions({ 
                    cursorBlinking: isEditMode ? 'blink' : 'hidden'
                  });
                }
              } else {
                const currentEditor = window.monacoEditor || editor;
                if (currentEditor) {
                  // Remove edit-mode class from editor-wrapper to show 30% opacity when just selecting
                  const editorWrapper = document.querySelector('.editor-wrapper');
                  if (editorWrapper && currentMode !== MODES.EDIT) {
                    editorWrapper.classList.remove('editor-edit-mode');
                  }

                  currentEditor.updateOptions({ 
                    cursorBlinking: 'hidden' // Hide cursor in selection mode
                  });
                }
              }
              window.selectedCell = firstCell;
              if (typeof syncResultPaneWithSelection === 'function') {
                syncResultPaneWithSelection();
              }
            }
            updateCellRangePill();
          }

          function enterSelectionModeAndSelectCell(targetCell, options = {}) {
            if (!targetCell) return;
            const forceExit = options.forceExit === true;
            if (!forceExit && currentMode === MODES.EDIT) {
              return;
            }

            const wasEditing = currentMode !== MODES.READY;
            if (forceExit || wasEditing) {
              const previouslyEditingCell = wasEditing ? window.selectedCell : null;
            setMode(MODES.READY);
            if (previouslyEditingCell) {
              clearEditingCellDisplayOverride(previouslyEditingCell);
            }

            const editorWrapper = document.querySelector('.editor-wrapper');
            if (editorWrapper) {
              editorWrapper.classList.remove('editor-edit-mode');
            }

            const currentEditor = window.monacoEditor || editor;
            if (currentEditor && typeof currentEditor.updateOptions === 'function') {
              currentEditor.updateOptions({
                cursorBlinking: 'hidden',
                cursorStyle: 'line'
              });
              blurMonacoEditor(currentEditor);
              currentEditor.render(true);
              }
            }

            if (typeof updateSelectionOverlay === 'function') {
              updateSelectionOverlay();
            }

            selectCell(targetCell);
          }

          function moveSelectionBy(deltaRow = 0, deltaCol = 0, extendSelection = false) {
            const currentCell = window.selectedCell;
            if (!currentCell) return;
            const currentRow = parseInt(currentCell.getAttribute("data-row"), 10);
            const currentCol = parseInt(currentCell.getAttribute("data-col"), 10);
            if (Number.isNaN(currentRow) || Number.isNaN(currentCol)) return;

            const maxRow = typeof window.totalGridRows === 'number' ? window.totalGridRows - 1 : currentRow;
            const maxCol = typeof window.totalGridColumns === 'number' ? window.totalGridColumns - 1 : currentCol;

            const targetRow = Math.max(0, Math.min(maxRow, currentRow + deltaRow));
            const targetCol = Math.max(0, Math.min(maxCol, currentCol + deltaCol));

            const targetCell = gridBody.querySelector(`tr[data-row="${targetRow}"] td[data-col="${targetCol}"]`);
            if (!targetCell) return;

            if (extendSelection) {
              const anchor = selectionAnchorCell || currentCell;
              selectRange(anchor, targetCell);
            } else {
              enterSelectionModeAndSelectCell(targetCell);
              selectionAnchorCell = targetCell;
            }
            targetCell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
          }

          // Function to select a single cell programmatically (selection mode)
          function selectCell(cell) {
            if (!cell) return;
            clearSelection();
            selectedCells.add(cell);
            cell.classList.add("selected");
            // Only add edit-mode class if actually in edit mode
            if (isEditMode) {
              cell.classList.add("edit-mode");
            }
            // Single cell selection - all edges should have borders
            cell.setAttribute("data-selection-edge", "top bottom left right");
            window.selectedCell = cell;
            selectionAnchorCell = cell;

            // Update selection overlay
            updateSelectionOverlay();
            updateHeaderHighlightsFromSelection();

            // Log cell information
            const cellRef = cell.getAttribute("data-ref");
            const sheetName = "Sheet1";
            const sheetId = window.hfSheetId !== undefined ? window.hfSheetId : 0;
            console.log(`Cell clicked - Sheet: ${sheetName}, Cell: ${cellRef || 'N/A'}, Sheet ID: ${sheetId}`);

            // Only update formula editor for data cells (not headers)
            const isHeader = cell.tagName === "TH" || cell === cell.parentElement.querySelector("td:first-child");
            if (!isHeader) {
              let editorValue = getStoredFormulaText(cell) || "";
              if (!editorValue) {
                const cellRef = cell.getAttribute("data-ref");
                if (cellRef && window.hf) {
                  const address = cellRefToAddress(cellRef);
                  if (address) {
                    const [row, col] = address;
                    const sheetId = typeof window.hfSheetId === "number" ? window.hfSheetId : 0;
                    let hasFormula = false;
                    try {
                      const cellFormula = window.hf.getCellFormula({ col, row, sheet: sheetId });
                      if (cellFormula) {
                        hasFormula = true;
                        editorValue = cellFormula.startsWith("=") ? cellFormula.substring(1) : cellFormula;
                        const formulaToStore = cellFormula.startsWith("=") ? cellFormula : `=${cellFormula}`;
                        cell.setAttribute("data-formula", formulaToStore);
                        if (cellRef && !readRawFormula(cellRef, sheetId)) {
                          saveRawFormula(cellRef, formulaToStore, sheetId);
                        }
                      }
                    } catch (formulaError) {
                      hasFormula = false;
                    }

                    if (!hasFormula) {
                      try {
                        const cellValue = window.hf.getCellValue({ col, row, sheet: sheetId });
                        if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
                          editorValue = cellValue.toString();
                          cell.setAttribute("data-formula", editorValue);
                          if (cellRef && !readRawFormula(cellRef, sheetId)) {
                            saveRawFormula(cellRef, editorValue, sheetId);
                          }
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
                // Formula starts on line 1
                const fullValue = ensureSevenLines(editorValue || "");

                // Calculate target cursor position
                const line1Content = (editorValue || "").split('\n')[0] || "";
                const targetColumn = Math.max(1, line1Content.length + 1);
                const targetLine = 1;

                // Set programmatic flag BEFORE setValue to prevent event listener from interfering
                window.isProgrammaticCursorChange = true;

                // Remove edit-mode class from editor-wrapper to show 30% opacity when just selecting
                // This applies to cells with any content (formula or value)
                const editorWrapper = document.querySelector('.editor-wrapper');
                if (editorWrapper && currentMode !== MODES.EDIT) {
                  editorWrapper.classList.remove('editor-edit-mode');
                }

                // Set the value directly using setValue
                // This is more reliable than executeEdits
                // Use requestAnimationFrame to ensure the editor is ready
                requestAnimationFrame(() => {
                  currentEditor.setValue(fullValue);

                  // Set cursor position after setting value
                  if (window.monaco && window.monaco.Range) {
                    currentEditor.setPosition({ lineNumber: targetLine, column: targetColumn });
                  } else {
                    // Fallback if monaco isn't loaded yet
                    setTimeout(() => {
                      currentEditor.setPosition({ lineNumber: targetLine, column: targetColumn });
                    }, 0);
                  }

                  // Reset flag after operations complete
                  Promise.resolve().then(() => {
                    window.isProgrammaticCursorChange = false;
                    if (typeof updateCellChips === 'function') {
                      requestAnimationFrame(() => updateCellChips());
                    }
                  });

                  currentEditor.updateOptions({ 
                    cursorBlinking: isEditMode ? 'blink' : 'hidden',
                    cursorStyle: 'line'
                  });

                  if (currentMode !== MODES.READY && typeof currentEditor.focus === 'function') {
                    currentEditor.focus();
                  } else if (currentMode === MODES.READY) {
                    blurMonacoEditor(currentEditor);
                    currentEditor.render(true);
                  }
                });
              }
            } else {
              const currentEditor = window.monacoEditor || editor;
              if (currentEditor) {
                currentEditor.updateOptions({ 
                  cursorBlinking: 'hidden'
                });
                // Blur and force re-render to ensure cursor is hidden
                blurMonacoEditor(currentEditor);
                currentEditor.render(true);
              }
            }
            updateCellRangePill();
            if (typeof syncResultPaneWithSelection === 'function') {
              syncResultPaneWithSelection();
            }
            if (typeof updateCellChips === 'function') {
              requestAnimationFrame(() => updateCellChips());
            }
          }

          // Disable editor initially (no cell selected) - will be set when Monaco loads
          // Ctrl+Enter command is set up in the Monaco initialization above

          // Handle mouseup to end selection
          document.addEventListener("mouseup", () => {
            if (isSelecting) {
              isSelecting = false;
              selectionStart = null;
            }
          });

          // Handle keydown events
          const readyNavigationMap = {
            ArrowUp: { row: -1, col: 0 },
            ArrowDown: { row: 1, col: 0 },
            ArrowLeft: { row: 0, col: -1 },
            ArrowRight: { row: 0, col: 1 }
          };

          const editingCommitMap = (shiftKey) => ({
            Enter: { row: shiftKey ? -1 : 1, col: 0 }
          });

          const editingNavigationMap = {
            ArrowUp: { row: -1, col: 0 },
            ArrowDown: { row: 1, col: 0 },
            ArrowLeft: { row: 0, col: -1 },
            ArrowRight: { row: 0, col: 1 }
          };

          function handleEnterModePointerNavigation() {
              return false;
          }

          function getStoredFormulaText(cell) {
            if (!cell) return "";
            const cellRef = cell.getAttribute("data-ref");
            const sheetId = getActiveSheetId();
            if (cellRef) {
              const rawFormula = readRawFormula(cellRef, sheetId);
              if (rawFormula) {
                return rawFormula.startsWith("=") ? rawFormula.substring(1) : rawFormula;
              }
            }
            const storedFormula = cell.getAttribute("data-formula") || "";
            if (storedFormula && storedFormula.startsWith("=")) {
              return storedFormula.substring(1);
            }
            return storedFormula;
          }

          function cancelEditingSession() {
            if (currentMode === MODES.READY || !window.selectedCell) {
              return;
            }
            const editingCell = window.selectedCell;
            const currentEditor = window.monacoEditor || editor;
            cancelFormulaSelection();
            const lastSavedFormula = getStoredFormulaText(editingCell);
            if (currentEditor && typeof currentEditor.setValue === 'function') {
              const fullValue = ensureSevenLines(lastSavedFormula || "");
              currentEditor.setValue(fullValue);
              if (typeof updateCellChips === 'function') {
                requestAnimationFrame(() => updateCellChips());
              }
            }
            clearEditingCellDisplayOverride(editingCell);
            setMode(MODES.READY);
            enterSelectionModeAndSelectCell(editingCell, { forceExit: true });
          }

          function getNormalizedEditorValue() {
            const currentEditor = window.monacoEditor || editor;
            if (!currentEditor || typeof currentEditor.getValue !== 'function') {
              return "";
            }
            const allLines = currentEditor.getValue().split('\n');
            return stripComments(allLines.join('\n')).trim();
          }

          function getRawEditorValue() {
            const currentEditor = window.monacoEditor || editor;
            if (!currentEditor || typeof currentEditor.getValue !== 'function') {
              return "";
            }
            const fullValue = currentEditor.getValue();
            if (typeof fullValue !== "string" || fullValue.length === 0) {
              return "";
            }
            const lines = fullValue.split('\n');
            let lastContentIndex = lines.length - 1;
            while (lastContentIndex >= 0 && lines[lastContentIndex].trim().length === 0) {
              lastContentIndex--;
            }
            if (lastContentIndex < 0) {
              return "";
            }
            return lines.slice(0, lastContentIndex + 1).join('\n');
          }

          function commitEditorChanges(moveDelta = null) {
            if (!window.selectedCell) return false;
            const cell = window.selectedCell;
            const cellRef = cell.getAttribute("data-ref");
            if (!cellRef) return false;
            cancelFormulaSelection();
            let formula = getNormalizedEditorValue();
            const rawFormulaInput = getRawEditorValue();
            if (formula) {
              if (!formula.startsWith("=")) {
                formula = "=" + formula;
              }
              setCellValue(cellRef, formula, { rawFormula: rawFormulaInput });
            } else {
              setCellValue(cellRef, "", { rawFormula: "" });
            }
            clearEditingCellDisplayOverride(cell);
            setMode(MODES.READY);
            if (moveDelta) {
              moveSelectionBy(moveDelta.row, moveDelta.col, false);
            } else {
              enterSelectionModeAndSelectCell(cell, { forceExit: true });
            }
            return true;
          }

          const handleGlobalKeyDown = (e) => {
            const currentEditor = window.monacoEditor || editor;
            const isEditorFocused = isMonacoEditorFocused(currentEditor);
            const activeElement = document.activeElement;
            const editorDomNode = currentEditor && typeof currentEditor.getDomNode === 'function'
              ? currentEditor.getDomNode()
              : null;
            const editorTextArea = editorDomNode ? editorDomNode.querySelector('textarea') : null;
            const isExternalTextInputFocused = !!activeElement
              && activeElement !== editorTextArea
              && (activeElement.tagName === 'INPUT' || activeElement.tagName === 'TEXTAREA');

            if (clappyInput && activeElement === clappyInput) {
              return;
            }

            if (suppressNextDocumentEnter && e.key === 'Enter') {
              suppressNextDocumentEnter = false;
              e.preventDefault();
              return;
            }

            const keyLower = typeof e.key === 'string' ? e.key.toLowerCase() : '';
            const isModifierCombo = (e.ctrlKey || e.metaKey) && !e.altKey;
            const isUndoShortcut = isModifierCombo && keyLower === 'z' && !e.shiftKey;
            const isRedoShortcut = isModifierCombo && ((keyLower === 'y' && !e.shiftKey) || (keyLower === 'z' && e.shiftKey));

            if ((isUndoShortcut || isRedoShortcut) && !isExternalTextInputFocused) {
              if (currentMode === MODES.READY && !isEditorFocused) {
                e.preventDefault();
                handleUndoRedoAction(isUndoShortcut ? 'undo' : 'redo');
                return;
              }
            }

            if (!window.selectedCell) {
              return;
            }

            if ((currentMode === MODES.ENTER || currentMode === MODES.EDIT) && e.key === "Escape") {
              e.preventDefault();
              cancelEditingSession();
              return;
            }

            if (e.key === "F2") {
              e.preventDefault();
              if (currentMode === MODES.READY) {
                enterEditMode(window.selectedCell);
              } else if (currentMode === MODES.ENTER) {
                promoteEnterModeToEdit();
              }
              return;
            }

            if (currentMode === MODES.READY) {
              if ((e.key === "Delete" || e.key === "Backspace") && !isExternalTextInputFocused) {
                if (selectedCells && selectedCells.size > 0) {
                  e.preventDefault();
                  clearSelectedCellContents();
                }
                return;
              }
              if (readyNavigationMap[e.key]) {
                e.preventDefault();
                moveSelectionBy(readyNavigationMap[e.key].row, readyNavigationMap[e.key].col, e.shiftKey);
                return;
              }
              if (e.key === 'Tab') {
                e.preventDefault();
                moveSelectionBy(0, e.shiftKey ? -1 : 1, false);
                return;
              }
              if (e.key === 'Enter') {
                e.preventDefault();
                moveSelectionBy(e.shiftKey ? -1 : 1, 0, false);
                return;
              }

              const isPrintableChar = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
              const specialKeys = ['Tab', 'Enter', 'Escape', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown', 'Delete', 'Backspace', 'Insert', 'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12'];
              const isSpecialKey = specialKeys.includes(e.key);
              if (!isEditorFocused && isPrintableChar && !isSpecialKey) {
                e.preventDefault();
                enterEditMode(window.selectedCell, null, { modeOverride: MODES.ENTER, initialText: e.key });
              }
              return;
            }

            if (currentMode === MODES.ENTER || currentMode === MODES.EDIT) {
              if (e.key === 'F4') {
                const didToggle = toggleReferenceAtCursor();
                if (didToggle) {
                  e.preventDefault();
                  return;
                }
              }
            }

            const pointerNavigationActive = currentMode === MODES.ENTER
              || currentMode === MODES.EDIT;
            if (pointerNavigationActive) {
              const handledPointerNav = handleEnterModePointerNavigation(e);
              if (handledPointerNav) {
                return;
              }
            }
            if (currentMode === MODES.EDIT) {
              const navigationMove = editingNavigationMap[e.key];
              if (navigationMove) {
                if (isEditorFocused) {
                  return;
                }
                e.preventDefault();
                commitEditorChanges(navigationMove);
                return;
              }
            }

            if ((currentMode === MODES.ENTER || currentMode === MODES.EDIT) && !isEditorFocused) {
              if (!(e.ctrlKey || e.metaKey)) {
                const commitMap = editingCommitMap(e.shiftKey);
                if (commitMap[e.key]) {
                  e.preventDefault();
                  commitEditorChanges(commitMap[e.key]);
                  suppressNextDocumentEnter = true;
                  return;
                }
              }
            }
          };
          document.addEventListener("keydown", handleGlobalKeyDown, true);

          // Exit edit mode when clicking on a different cell (single click)
          // This is handled in the click handler by setting isEditMode = false


          function createReferenceFromCells(startCell, endCell) {
            if (!startCell || !endCell) return '';
            const startRow = parseInt(startCell.getAttribute("data-row"), 10);
            const startCol = parseInt(startCell.getAttribute("data-col"), 10);
            const endRow = parseInt(endCell.getAttribute("data-row"), 10);
            const endCol = parseInt(endCell.getAttribute("data-col"), 10);
            if ([startRow, startCol, endRow, endCol].some((value) => Number.isNaN(value))) {
              return '';
            }
            const refStart = addressToCellRef(startRow, startCol);
            const refEnd = addressToCellRef(endRow, endCol);
            return startRow === endRow && startCol === endCol ? refStart : `${refStart}:${refEnd}`;
          }

          function handleFormulaSelectionStart() {}

          function handleFormulaSelectionHover() {}

          function finalizeFormulaSelectionIfNeeded() {}

          function buildGrid(rows = 150, columns = 50) {
            console.log("buildGrid function called with rows:", rows, "columns:", columns);
            window.totalGridRows = rows;
            window.totalGridColumns = columns;
            if (!gridHeaderElement) {
              console.error("gridHeader not found!");
              return;
            }
            const headerRow = gridHeaderElement.querySelector("tr");
            if (!headerRow) {
              console.error("headerRow not found!");
              return;
            }

            // Generate column headers
            const headers = Array.from({ length: columns }, (_, i) => {
              let col = "";
              let num = i;
              do {
                col = String.fromCharCode(65 + (num % 26)) + col;
                num = Math.floor(num / 26) - 1;
              } while (num >= 0);
              return col;
            });

            // Add header cells with resizing
            // Use a shared resizing state outside the loop
            if (!window.gridColumnResizing) {
              window.gridColumnResizing = {
                current: null,
                handleMouseMove: null,
                handleMouseUp: null,
                justResized: false // Flag to prevent selection after resize
              };
            }

            headers.forEach((header, index) => {
              const th = document.createElement("th");
              th.scope = "col";
              th.textContent = header;
              th.style.width = "80px";
              th.setAttribute("data-col-index", index);

              // Add column resizing with better detection
              const handleColResize = (e) => {
                const rect = th.getBoundingClientRect();
                const rightEdge = rect.right;
                // Check if click is within 10px of right edge
                if (e.clientX >= rightEdge - 10) {
                  window.gridColumnResizing.current = {
                    th: th,
                    colIndex: index,
                    startX: e.clientX,
                    startWidth: th.offsetWidth
                  };
                  document.body.style.cursor = "col-resize";
                  e.preventDefault();
                  e.stopPropagation();
                  e.stopImmediatePropagation();
                  return true;
                }
                return false;
              };

              // Use capture phase to ensure this runs first
              th.addEventListener("mousedown", handleColResize, true);

              // Allow column headers to be selected (but not visually highlighted)
              th.addEventListener("click", (e) => {
                if (currentMode !== MODES.READY) {
                  e.preventDefault();
                  e.stopPropagation();
                  return;
                }
                // Only select if not resizing and resize didn't just happen
                if (!window.gridColumnResizing.current && !window.gridColumnResizing.justResized) {
                  selectCell(th);
                } else if (window.gridColumnResizing.justResized) {
                  // Prevent selection if we just resized
                  e.preventDefault();
                  e.stopPropagation();
                }
              });

              // Add cursor feedback for column headers
              th.addEventListener("mousemove", (e) => {
                const rect = th.getBoundingClientRect();
                const rightEdge = rect.right;
                if (e.clientX >= rightEdge - 10) {
                  th.style.cursor = "col-resize";
                } else {
                  th.style.cursor = "default";
                }
              });

              headerRow.appendChild(th);
            });

            // Set up global mouse handlers only once
            if (!window.gridColumnResizing.handleMouseMove) {
              window.gridColumnResizing.rafId = null;
              window.gridColumnResizing.handleMouseMove = (e) => {
                if (window.gridColumnResizing.current) {
                  // Cancel any pending animation frame
                  if (window.gridColumnResizing.rafId) {
                    cancelAnimationFrame(window.gridColumnResizing.rafId);
                  }

                  // Batch updates in requestAnimationFrame
                  window.gridColumnResizing.rafId = requestAnimationFrame(() => {
                    const resizing = window.gridColumnResizing.current;
                    if (!resizing) return;

                    const newWidth = resizing.startWidth + (e.clientX - resizing.startX);
                    if (newWidth >= 40) {
                      // Update header
                      resizing.th.style.width = `${newWidth}px`;
                      resizing.th.style.minWidth = `${newWidth}px`;

                      // Update all cells in this column - batch DOM updates
                      const allCells = gridBody.querySelectorAll(`td[data-col="${resizing.colIndex}"]`);
                      const widthStr = `${newWidth}px`;
                      for (let i = 0; i < allCells.length; i++) {
                        allCells[i].style.width = widthStr;
                        allCells[i].style.minWidth = widthStr;
                      }

                      // Update selection overlay to match new cell sizes (if there's a selection)
                      if (typeof updateSelectionOverlay === 'function') {
                        updateSelectionOverlay();
                      }
                    }
                    window.gridColumnResizing.rafId = null;
                  });
                }
              };

              window.gridColumnResizing.handleMouseUp = () => {
                if (window.gridColumnResizing.current) {
                  // Cancel any pending animation frame
                  if (window.gridColumnResizing.rafId) {
                    cancelAnimationFrame(window.gridColumnResizing.rafId);
                    window.gridColumnResizing.rafId = null;
                  }
                  // Set flag to prevent click from selecting
                  window.gridColumnResizing.justResized = true;
                  // Clear flag after a short delay to allow normal clicking later
                  setTimeout(() => {
                    window.gridColumnResizing.justResized = false;
                  }, 100);
                  window.gridColumnResizing.current = null;
                  document.body.style.cursor = "";
                }
              };

              document.addEventListener("mousemove", window.gridColumnResizing.handleMouseMove);
              document.addEventListener("mouseup", window.gridColumnResizing.handleMouseUp);
            }

            // Initialize row resizing state
            if (!window.gridRowResizing) {
              window.gridRowResizing = {
                current: null,
                handleMouseMove: null,
                handleMouseUp: null,
                justResized: false // Flag to prevent selection after resize
              };
            }

            // Set up row resizing handlers only once
            if (!window.gridRowResizing.handleMouseMove) {
              window.gridRowResizing.rafId = null;
              window.gridRowResizing.handleMouseMove = (e) => {
                if (window.gridRowResizing.current) {
                  // Cancel any pending animation frame
                  if (window.gridRowResizing.rafId) {
                    cancelAnimationFrame(window.gridRowResizing.rafId);
                  }

                  // Batch updates in requestAnimationFrame
                  window.gridRowResizing.rafId = requestAnimationFrame(() => {
                    const resizing = window.gridRowResizing.current;
                    if (!resizing) return;

                    const newHeight = resizing.startHeight + (e.clientY - resizing.startY);
                    if (newHeight >= 15) {
                      // Update row
                      resizing.row.style.height = `${newHeight}px`;
                      resizing.row.style.minHeight = `${newHeight}px`;

                      // Update all cells in this row - batch DOM updates
                      const allCells = resizing.row.querySelectorAll("td");
                      const heightStr = `${newHeight}px`;
                      for (let i = 0; i < allCells.length; i++) {
                        allCells[i].style.height = heightStr;
                        allCells[i].style.minHeight = heightStr;
                      }

                      // Update selection overlay to match new cell sizes (if there's a selection)
                      if (typeof updateSelectionOverlay === 'function') {
                        updateSelectionOverlay();
                      }
                    }
                    window.gridRowResizing.rafId = null;
                  });
                }
              };

              window.gridRowResizing.handleMouseUp = () => {
                if (window.gridRowResizing.current) {
                  // Cancel any pending animation frame
                  if (window.gridRowResizing.rafId) {
                    cancelAnimationFrame(window.gridRowResizing.rafId);
                    window.gridRowResizing.rafId = null;
                  }
                  // Set flag to prevent click from selecting
                  window.gridRowResizing.justResized = true;
                  // Clear flag after a short delay to allow normal clicking later
                  setTimeout(() => {
                    window.gridRowResizing.justResized = false;
                  }, 100);
                  window.gridRowResizing.current = null;
                  document.body.style.cursor = "";
                }
              };

              document.addEventListener("mousemove", window.gridRowResizing.handleMouseMove);
              document.addEventListener("mouseup", window.gridRowResizing.handleMouseUp);
            }

            const fragment = document.createDocumentFragment();

            for (let r = 0; r < rows; r += 1) {
              const rowElement = document.createElement("tr");
              rowElement.setAttribute("data-row", r);
              rowElement.style.height = "20px";

              const indexCell = document.createElement("td");
              indexCell.textContent = r + 1;
              indexCell.setAttribute("data-row-index", r);
              indexCell.style.pointerEvents = "auto";

              // Add row resizing - handle both direct clicks and ::after pseudo-element
              const handleRowResize = (e) => {
                const rect = indexCell.getBoundingClientRect();
                const bottomEdge = rect.bottom;
                // Check if click is within 10px of bottom edge
                if (e.clientY >= bottomEdge - 10) {
                  window.gridRowResizing.current = {
                    row: rowElement,
                    rowIndex: r,
                    startY: e.clientY,
                    startHeight: rowElement.offsetHeight
                  };
                  document.body.style.cursor = "row-resize";
                  e.preventDefault();
                  e.stopPropagation();
                  e.stopImmediatePropagation();
                }
              };

              // Use capture phase to ensure this runs first
              indexCell.addEventListener("mousedown", handleRowResize, true);

              // Also handle mousemove to show cursor
              indexCell.addEventListener("mousemove", (e) => {
                const rect = indexCell.getBoundingClientRect();
                const bottomEdge = rect.bottom;
                if (e.clientY >= bottomEdge - 10) {
                  indexCell.style.cursor = "row-resize";
                } else {
                  indexCell.style.cursor = "default";
                }
              });

              // Allow row headers to be selected (but not visually highlighted)
              indexCell.addEventListener("click", (e) => {
                if (currentMode !== MODES.READY) {
                  e.preventDefault();
                  e.stopPropagation();
                  return;
                }
                // Only select if not resizing and resize didn't just happen
                if (!window.gridRowResizing.current && !window.gridRowResizing.justResized) {
                selectCell(indexCell);
                } else if (window.gridRowResizing.justResized) {
                  // Prevent selection if we just resized
                  e.preventDefault();
                  e.stopPropagation();
                }
              });

              rowElement.appendChild(indexCell);

              for (let c = 0; c < columns; c += 1) {
                const cell = document.createElement("td");
                cell.setAttribute("data-row", r);
                cell.setAttribute("data-col", c);
                cell.setAttribute("data-ref", `${headers[c]}${r + 1}`);
                cell.style.width = "80px"; // Set initial width
                cell.style.height = "20px"; // Set initial height

                const display = document.createElement("span");
                display.className = "grid-cell-display";
                display.textContent = "";

                cell.appendChild(display);

                // Handle drag selection
                cell.addEventListener("mousedown", (e) => {
                  // Don't start selection if resizing
                  if (window.gridColumnResizing.current || window.gridRowResizing.current) {
                    return;
                  }

                  // Don't prevent default on double click
                  if (e.detail === 2) {
                    return;
                  }

                  isSelecting = true;
                  selectionStart = cell;
                  selectCell(cell);
                  e.preventDefault();
                });

                cell.addEventListener("mouseenter", (e) => {
                  if (isSelecting && selectionStart) {
                    selectRange(selectionStart, cell);
                  }
                });

                // Single click handler (when not dragging) - selection mode
                cell.addEventListener("click", (e) => {
                  if (currentMode !== MODES.READY) {
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                  }
                  if (isSelecting) {
                    return;
                  }
                  if (e.detail > 1) {
                    return;
                  }
                  enterSelectionModeAndSelectCell(cell);
                });

                // Double click handler - edit mode
                cell.addEventListener("dblclick", (e) => {
                  if (currentMode !== MODES.READY) {
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                  }
                  e.preventDefault();
                  e.stopPropagation();
                  e.stopImmediatePropagation();
                  enterEditMode(cell);
                });

                rowElement.appendChild(cell);
              }

              fragment.appendChild(rowElement);
            }

            gridBody.appendChild(fragment);
          }

        initResizablePanes();

        initClappyChat({
          transcript: clappyTranscript,
          form: clappyForm,
          input: clappyInput,
          chatToggle: chatCollapseToggle,
          consoleEl: clappyConsole,
          chatContainer: mainChatContainer,
          middleSection,
          gridContainer
        });

          // Build grid immediately - don't wait for Monaco
          console.log("About to call buildGrid()");
          buildGrid();
          console.log("buildGrid() called");

          // Start the grid with the sample fake data set (row-major from A1)
          populateGridWithData(SAMPLE_FAKE_DATA, "A1");

          // Select A1 (first cell) after grid is built
          const firstCell = gridBody.querySelector("tr:first-child td[data-col='0']");
          if (firstCell) {
            selectCell(firstCell);
          }

          // Initialize custom overlay scrollbars
          initCustomScrollbars();

          // Initialize sliding panes toggle functionality
          // formulaEditor already declared above, reuse it
          const paneHeaders = document.querySelectorAll('.pane-header');

          function updatePaneHeight(pane) {
            if (!pane || !formulaEditor) return;
            const formulaEditorHeight = formulaEditor.offsetHeight;
            const headerHeight = 22; // Height of collapsed pane
            const expandedHeight = formulaEditorHeight * 0.2; // 20% of formula editor

            // Use requestAnimationFrame to ensure smooth animation
            requestAnimationFrame(() => {
              if (pane.classList.contains('collapsed')) {
                pane.style.height = `${headerHeight}px`;
              } else {
                pane.style.height = `${expandedHeight}px`;
              }

              // After the pane height settles, relayout Monaco so it fills the new space
              requestAnimationFrame(() => {
                if (window.monacoEditor && typeof window.monacoEditor.layout === 'function') {
                  window.monacoEditor.layout();
                }
              });
            });
          }

          // Update pane heights on resize
          function updateAllPaneHeights() {
            document.querySelectorAll('.sliding-pane').forEach(pane => {
              updatePaneHeight(pane);
            });
          }

          paneHeaders.forEach(header => {
            header.addEventListener('click', () => {
              const pane = header.closest('.sliding-pane');
              if (pane) {
                pane.classList.toggle('collapsed');
                  updatePaneHeight(pane);
                  if (header.dataset.pane === 'result' && typeof window.updateResultHeaderValue === "function") {
                    window.updateResultHeaderValue();
                  }
              }
            });
          });

          // Update heights on window resize
          window.addEventListener('resize', updateAllPaneHeights);

          // Initial update - set all panes to collapsed state
          // Disable transitions on initial load to prevent animation
          document.querySelectorAll('.sliding-pane').forEach(pane => {
            if (!pane.classList.contains('collapsed')) {
              pane.classList.add('collapsed');
            }
            // Temporarily disable transition for initial setup
            pane.style.transition = 'none';
            pane.style.height = '22px';
            // Re-enable transition after a brief delay
            requestAnimationFrame(() => {
              pane.style.transition = '';
            });
          });
          if (typeof window.updateResultHeaderValue === "function") {
            window.updateResultHeaderValue();
          }
        } catch (error) {
          console.error("Error initializing app:", error);
        }
      }


