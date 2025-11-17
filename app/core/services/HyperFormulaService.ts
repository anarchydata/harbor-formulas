import type { HyperFormula } from 'hyperformula';
import { BaseService } from '../BaseService';

/**
 * Service for managing HyperFormula instance
 * Handles initialization and sheet management
 */
export class HyperFormulaService extends BaseService {
  private hf: HyperFormula | null = null;
  private sheetId: number = 0;
  private namedRanges: Set<string> = new Set();
  private rawFormulaStore: Map<string, string> = new Map();

  async initialize(): Promise<void> {
    if (this.initialized) {
      return;
    }

    const { HyperFormula } = await import('hyperformula');

    this.hf = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
      undoLimit: 500,
      useArrayArithmetic: true,
    });

    // Initialize named ranges storage
    this.namedRanges = new Set();

    // Store raw formulas (including comments) keyed by sheet + cell reference
    this.rawFormulaStore = new Map();

    // Add Sheet1
    this.hf.addSheet('Sheet1');
    this.sheetId = 0;

    // Verify the sheet exists
    const sheetNames = this.hf.getSheetNames();
    if (sheetNames.length === 0 || sheetNames[0] !== 'Sheet1') {
      throw new Error(
        `Sheet1 was not created successfully. Available sheets: ${sheetNames.join(', ')}`
      );
    }

    // Store globally for backward compatibility during migration
    window.hf = this.hf;
    window.namedRanges = this.namedRanges;
    window.rawFormulaStore = this.rawFormulaStore;
    window.hfSheetId = this.sheetId;

    this.initialized = true;
  }

  destroy(): void {
    // Cleanup if needed
    this.hf = null;
    this.initialized = false;
  }

  getInstance(): HyperFormula {
    if (!this.hf) {
      throw new Error('HyperFormula not initialized');
    }
    return this.hf;
  }

  getSheetId(): number {
    return this.sheetId;
  }

  addSheet(name: string): string {
    if (!this.hf) {
      throw new Error('HyperFormula not initialized');
    }
    return this.hf.addSheet(name);
  }

  renameSheet(sheetId: number, name: string): void {
    if (!this.hf) {
      throw new Error('HyperFormula not initialized');
    }
    this.hf.renameSheet(sheetId, name);
  }

  getSheetNames(): string[] {
    if (!this.hf) {
      throw new Error('HyperFormula not initialized');
    }
    return this.hf.getSheetNames();
  }

  getNamedRanges(): Set<string> {
    return this.namedRanges;
  }

  getRawFormulaStore(): Map<string, string> {
    return this.rawFormulaStore;
  }
}

