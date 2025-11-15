/**
 * Diagnostics provider for formula validation
 * Validates Excel formulas and marks errors with red squiggles
 */
function isInsideQuotes(text, position) {
  if (!text || typeof position !== 'number' || position < 0) {
    return false;
  }

  let inDoubleQuotes = false;
  let escaped = false;
  const limit = Math.min(position, text.length);

  for (let i = 0; i < limit; i++) {
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
      inDoubleQuotes = !inDoubleQuotes;
    }
  }

  return inDoubleQuotes;
}

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

    // Handle Excel-style escaped quotes ("") and sheet names wrapped in ''
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

function findTrailingCommaPositions(text) {
  const positions = [];
  let inDoubleQuotes = false;
  let inSingleQuotes = false;
  let braceDepth = 0;
  let parenDepth = 0;

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
      parenDepth++;
      continue;
    }

    if (char === ')') {
      if (parenDepth > 0) {
        parenDepth--;
      }
      continue;
    }

    if (char === ',' && parenDepth > 0) {
      let lookaheadIndex = i + 1;
      while (lookaheadIndex < text.length && /\s/.test(text[lookaheadIndex])) {
        lookaheadIndex++;
      }
      if (text[lookaheadIndex] === ')') {
        positions.push(i);
      }
    }
  }

  return positions;
}

function findDanglingEndCommaIndex(text) {
  for (let i = text.length - 1; i >= 0; i--) {
    const char = text[i];
    if (/\s/.test(char)) {
      continue;
    }
    if (char === ',' && !isInsideQuotes(text, i)) {
      return i;
    }
    break;
  }
  return -1;
}

function stripLineComments(text) {
  if (!text) {
    return '';
  }

  let result = '';
  let inDoubleQuotes = false;
  let inSingleQuotes = false;
  let inLineComment = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inLineComment) {
      if (char === '\n' || char === '\r') {
        inLineComment = false;
        result += char;
      } else {
        result += ' ';
      }
      continue;
    }

    if (!inSingleQuotes && char === '"') {
      if (nextChar === '"') {
        result += char;
        i++;
        result += text[i];
        continue;
      }
      inDoubleQuotes = !inDoubleQuotes;
      result += char;
      continue;
    }

    if (!inDoubleQuotes && char === "'") {
      if (nextChar === "'") {
        result += char;
        i++;
        result += text[i];
        continue;
      }
      inSingleQuotes = !inSingleQuotes;
      result += char;
      continue;
    }

    if (!inSingleQuotes && !inDoubleQuotes && char === '/' && nextChar === '/') {
      inLineComment = true;
      result += ' ';
      i++;
      result += ' ';
      continue;
    }

    result += char;
  }

  return result;
}

export function createDiagnosticsProvider(monaco) {
  return {
    provideDiagnostics: function(model, lastResult) {
      const diagnostics = [];
      const text = model.getValue();
      const sanitizedText = stripLineComments(text);
      
      const lines = sanitizedText.split('\n');
      if (!lines.length) return { markers: [] };

      const formulaText = sanitizedText;
      const trimmedFormulaText = formulaText.trim();
      if (!trimmedFormulaText) return { markers: [] };

      const lineOffset = 0;
      const toEditorLineFromIndex = (lineIndexZeroBased) => lineIndexZeroBased + 1 + lineOffset;
      const toEditorLineFromCount = (lineCountOneBased) => lineCountOneBased + lineOffset;
      
      // Check for common syntax errors
      // 1. Check for periods in function arguments (should be commas)
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

        // Ignore dots that are part of a string literal (e.g., "0.00")
        if (isInsideQuotes(formulaText, tokenStartIndex) || isInsideQuotes(formulaText, dotIndex)) {
          continue;
        }

        // Ignore periods that are clearly part of a quoted string literal
        if (match[1]?.startsWith('"') && match[1]?.endsWith('"')) {
          continue;
        }
        // Ignore periods when there's a newline between the token and the period
        if (match[0].includes('\n')) {
          continue;
        }

        // Calculate line and column in the editor
        const beforeMatch = formulaText.substring(0, match.index);
        const formulaLines = beforeMatch.split('\n');
        const formulaLineIndex = formulaLines.length - 1;
        const editorLineNumber = toEditorLineFromIndex(formulaLineIndex);
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
              const functionStartEditorLine = toEditorLineFromIndex(functionStartLineIndex);
              const functionStartColumn = functionStartLines[functionStartLineIndex].length + 1;
              
              // Find the end of the function (last closing paren)
              const lastCloseParen = formulaText.lastIndexOf(')');
              const functionEndLines = formulaText.substring(0, lastCloseParen + 1).split('\n');
              const functionEndLineIndex = functionEndLines.length - 1;
              const functionEndEditorLine = toEditorLineFromIndex(functionEndLineIndex);
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
        const editorLineNumber = toEditorLineFromIndex(formulaLineIndex);
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

      // 3. Check for trailing commas before closing parentheses (only when parentheses are balanced)
      const trailingCommaPositions = findTrailingCommaPositions(formulaText);
      trailingCommaPositions.forEach((commaIndex) => {
        const beforeMatch = formulaText.substring(0, commaIndex);
        const formulaLines = beforeMatch.split('\n');
        const formulaLineIndex = formulaLines.length - 1;
        const editorLineNumber = toEditorLineFromIndex(formulaLineIndex);
        const column = formulaLines[formulaLineIndex].length + 1;

        diagnostics.push({
          severity: monaco.MarkerSeverity.Error,
          startLineNumber: editorLineNumber,
          startColumn: column,
          endLineNumber: editorLineNumber,
          endColumn: column + 1,
          message: 'Unexpected trailing comma before closing parenthesis',
          source: 'Formula Validator',
          code: 'TRAILING_COMMA'
        });
      });
      
      if (openParens === closeParens) {
        const danglingEndCommaIndex = findDanglingEndCommaIndex(formulaText);
        if (danglingEndCommaIndex !== -1) {
          const beforeMatch = formulaText.substring(0, danglingEndCommaIndex);
          const formulaLines = beforeMatch.split('\n');
          const formulaLineIndex = formulaLines.length - 1;
          const editorLineNumber = toEditorLineFromIndex(formulaLineIndex);
          const column = formulaLines[formulaLineIndex].length + 1;

          diagnostics.push({
            severity: monaco.MarkerSeverity.Error,
            startLineNumber: editorLineNumber,
            startColumn: column,
            endLineNumber: editorLineNumber,
            endColumn: column + 1,
            message: 'Unexpected trailing comma at end of formula',
            source: 'Formula Validator',
            code: 'END_TRAILING_COMMA'
          });
        }
      }
      
      // 4. Validate formula syntax with HyperFormula if available
      // Only validate if the formula looks complete (has closing parens, etc.)
      if (window.hf && trimmedFormulaText) {
        try {
          // Ensure formula starts with '='
          const testFormula = trimmedFormulaText.startsWith('=') ? trimmedFormulaText : '=' + trimmedFormulaText;
          
          // Check for common syntax errors first (parentheses, periods, etc.)
          const { openParens, closeParens } = analyzeParentheses(formulaText);
          
          // Only validate with HyperFormula if the formula looks syntactically complete
          // Don't validate incomplete formulas (missing closing parens, etc.)
          // This prevents false positives for formulas that are still being typed
          const isLikelyComplete = openParens === closeParens && trimmedFormulaText.length > 0;
          
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
                const formulaStart = trimmedFormulaText.startsWith('=') ? 1 : 0;
                diagnostics.push({
                  severity: monaco.MarkerSeverity.Error,
                  startLineNumber: lineOffset + 1,
                  startColumn: formulaStart + 1,
                  endLineNumber: lineOffset + 1,
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
                    startLineNumber: toEditorLineFromCount(lineNum),
                    startColumn: colNum + 1,
                    endLineNumber: toEditorLineFromCount(lineNum),
                    endColumn: colNum + 2,
                    message: 'Missing closing parenthesis',
                    source: 'Formula Validator',
                    code: 'MISSING_PARENTHESIS'
                  });
                }
              }
            } else if (closeParens > openParens) {
              // Extra closing parenthesis - always mark this as an error
              const { firstExtraClosingIndex } = analyzeParentheses(formulaText);
              
              if (firstExtraClosingIndex !== -1) {
                const formulaLines = formulaText.substring(0, firstExtraClosingIndex + 1).split('\n');
                const lineNum = formulaLines.length;
                const colNum = formulaLines[formulaLines.length - 1].length;
                
                diagnostics.push({
                  severity: monaco.MarkerSeverity.Error,
                  startLineNumber: toEditorLineFromCount(lineNum),
                  startColumn: colNum,
                  endLineNumber: toEditorLineFromCount(lineNum),
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

