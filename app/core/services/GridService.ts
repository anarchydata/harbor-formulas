import type Handsontable from 'handsontable';
import { BaseService } from '../BaseService';
import { createHandsontableGrid } from '../../grid/handsontableAdapter';
import type { HyperFormulaService } from './HyperFormulaService';

/**
 * Service for managing the Handsontable grid
 * Handles grid creation and basic operations
 */
export class GridService extends BaseService {
  private hot: Handsontable | null = null;
  private container: HTMLElement | null = null;

  constructor(
    container: HTMLElement,
    private hfService: HyperFormulaService
  ) {
    super();
    this.container = container;
  }

  initialize(): void {
    if (this.initialized || !this.container) {
      return;
    }

    if (!this.hfService.isInitialized()) {
      throw new Error('HyperFormulaService must be initialized first');
    }

    // Create Handsontable grid
    const gridController = createHandsontableGrid({
      container: this.container,
      rows: 150,
      columns: 50,
      // Event handlers will be added by controllers later
      onCellRender: undefined,
      onSelection: undefined,
      onDoubleClick: undefined,
      onBeforeKeyDown: undefined,
    });

    this.hot = gridController.hot;

    this.initialized = true;
  }

  destroy(): void {
    if (this.hot) {
      this.hot.destroy();
      this.hot = null;
    }
    this.initialized = false;
  }

  getInstance(): Handsontable {
    if (!this.hot) {
      throw new Error('Grid not initialized');
    }
    return this.hot;
  }

  getContainer(): HTMLElement {
    if (!this.container) {
      throw new Error('Container not set');
    }
    return this.container;
  }
}

