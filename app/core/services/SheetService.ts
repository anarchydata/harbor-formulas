import { BaseService } from '../BaseService';
import type { HyperFormulaService } from './HyperFormulaService';
import type { GridService } from './GridService';

/**
 * Service for managing sheets
 * Handles sheet creation, deletion, and switching
 */
export class SheetService extends BaseService {
  constructor(
    private hfService: HyperFormulaService,
    private gridService: GridService
  ) {
    super();
  }

  initialize(): void {
    if (this.initialized) {
      return;
    }

    if (!this.hfService.isInitialized()) {
      throw new Error('HyperFormulaService must be initialized first');
    }

    if (!this.gridService.isInitialized()) {
      throw new Error('GridService must be initialized first');
    }

    this.initialized = true;
  }

  destroy(): void {
    // Cleanup if needed
    this.initialized = false;
  }

  getActiveSheetId(): number {
    return this.hfService.getSheetId();
  }

  getSheetNames(): string[] {
    return this.hfService.getSheetNames();
  }

  addSheet(name: string): string {
    return this.hfService.addSheet(name);
  }

  renameSheet(sheetId: number, name: string): void {
    this.hfService.renameSheet(sheetId, name);
  }
}

