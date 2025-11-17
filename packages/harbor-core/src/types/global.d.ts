import type { HyperFormula } from 'hyperformula';

// Global window properties (will be replaced with proper services later)
declare global {
  interface Window {
    hf?: HyperFormula;
    rawFormulaStore?: Map<string, string>;
    namedRanges?: Set<string>;
    hfSheetId?: number;
    selectedCell?: HTMLElement | null;
    // Add more as needed for future features
  }
}

export {};

