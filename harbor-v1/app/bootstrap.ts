import { HyperFormulaService } from './core/services/HyperFormulaService';
import { GridService } from './core/services/GridService';
import { SheetService } from './core/services/SheetService';

/**
 * Minimal bootstrap - initializes core services only
 * No event listeners or interactivity - those will be added by controllers later
 */
export async function initializeApp() {
  try {
    // 1. Initialize HyperFormula
    const hfService = new HyperFormulaService();
    await hfService.initialize();

    // 2. Initialize Grid
    const gridContainer = document.getElementById('handsontableRoot');
    if (!gridContainer) {
      throw new Error('Grid container not found');
    }
    const gridService = new GridService(gridContainer, hfService);
    gridService.initialize();

    // 3. Initialize Sheet Service
    const sheetService = new SheetService(hfService, gridService);
    sheetService.initialize();

    // 4. Return services for use by controllers
    return {
      hfService,
      gridService,
      sheetService,
    };
  } catch (error) {
    console.error('Error initializing app:', error);
    throw error;
  }
}

