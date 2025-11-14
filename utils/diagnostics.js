/**
 * Diagnostics provider for formula validation
 * Validates Excel formulas and marks errors with red squiggles
 */
export function createDiagnosticsProvider(monaco) {
  return {
    provideDiagnostics: function(model, lastResult) {
      const diagnostics = [];
      const text = model.getValue();
      
      // Skip first line (the "=" visual indicator line)
      const lines = text.split('\n');
      if (lines.length < 2) return { markers: [] };
      
      // Get formula starting from line 2
      const formulaText = lines.slice(1).join('\n').trim();
      if (!formulaText) return { markers: [] };
      
      // Check for common syntax errors
      // 1. Check for periods in function arguments (should be commas)
      // Need to check if we're in a function context and if parentheses are closed
      const openParens = (formulaText.match(/\(/g) || []).length;
      const closeParens = (formulaText.match(/\)/g) || []).length;
      const hasUnclosedParens = openParens > closeParens;
      
      // Pattern: value followed by period (in function context)
      // Match periods that appear after values in function calls
      // Note: Single quotes are reserved by Excel for sheet names
      const periodPattern = /(\w+|"[^"]*"|[A-Z]+\$?\d+\$?)\s*\./g;
      let match;
      while ((match = periodPattern.exec(formulaText)) !== null) {
        // Calculate line and column in the editor
        const beforeMatch = formulaText.substring(0, match.index);
        const formulaLines = beforeMatch.split('\n');
        const formulaLineIndex = formulaLines.length - 1;
        const editorLineNumber = formulaLineIndex + 2; // +2 because line 1 is skipped, +1 for 1-based
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
      
      // 2. Check for double commas
      const doubleCommaRegex = /,\s*,/g;
      while ((match = doubleCommaRegex.exec(formulaText)) !== null) {
        const beforeMatch = formulaText.substring(0, match.index);
        const formulaLines = beforeMatch.split('\n');
        const formulaLineIndex = formulaLines.length - 1;
        const editorLineNumber = formulaLineIndex + 2; // +2 because line 1 is skipped, +1 for 1-based
        const column = formulaLines[formulaLineIndex].length + 1; // +1 for 1-based
        
        diagnostics.push({
          severity: monaco.MarkerSeverity.Warning,
          startLineNumber: editorLineNumber,
          startColumn: column,
          endLineNumber: editorLineNumber,
          endColumn: column + 1,
          message: 'Unexpected comma',
          source: 'Formula Validator',
          code: 'DOUBLE_COMMA'
        });
      }
      
      // 3. Validate formula syntax with HyperFormula if available
      // Only validate if the formula looks complete (has closing parens, etc.)
      if (window.hf && formulaText) {
        try {
          // Ensure formula starts with '='
          const testFormula = formulaText.startsWith('=') ? formulaText : '=' + formulaText;
          
          // Check for common syntax errors first (parentheses, periods, etc.)
          const openParens = (formulaText.match(/\(/g) || []).length;
          const closeParens = (formulaText.match(/\)/g) || []).length;
          
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
                    startLineNumber: lineNum + 1, // +1 because line 1 is the "=" line
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
              let unmatchedCount = 0;
              let firstUnmatched = -1;
              for (let i = 0; i < formulaText.length; i++) {
                if (formulaText[i] === '(') unmatchedCount++;
                if (formulaText[i] === ')') {
                  unmatchedCount--;
                  if (unmatchedCount < 0 && firstUnmatched === -1) {
                    firstUnmatched = i;
                    break;
                  }
                }
              }
              
              if (firstUnmatched !== -1) {
                const formulaLines = formulaText.substring(0, firstUnmatched + 1).split('\n');
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
      
      return { markers: diagnostics };
    }
  };
}

