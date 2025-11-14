/**
 * Helper function to create snippet template from function signature
 * This must be defined before it's used in completion providers and Tab handlers
 */
export function createFunctionSnippet(func) {
  const signature = func.signature;
  const match = signature.match(/^(\w+)\s*\((.*)\)$/);
  if (!match) {
    return `${func.name}($1)`;
  }

  const params = match[2];
  if (!params || params.trim() === '') {
    return `${func.name}()`;
  }

  const paramList = params.split(',').map((p) => p.trim());
  let placeholderIndex = 1;
  const snippetParams = paramList.map((param) => {
    if (param.endsWith('...')) {
      return `\${${placeholderIndex++}:${param.replace('...', '')}}`;
    }
    return `\${${placeholderIndex++}:${param}}`;
  });

  return `${func.name}(${snippetParams.join(', ')})`;
}

export function ensureSevenLines(value) {
  const lines = value.split('\n');
  while (lines.length < 7) {
    lines.push('');
  }
  return lines.slice(0, 7).join('\n');
}

export function stripComments(formulaText) {
  if (!formulaText) return formulaText;

  let result = '';
  let inString = false;
  let inBlockComment = false;
  let inInlineComment = false;
  let escaped = false;

  for (let i = 0; i < formulaText.length; i++) {
    const char = formulaText[i];
    const nextChar = i < formulaText.length - 1 ? formulaText[i + 1] : '';

    if (escaped) {
      escaped = false;
      if (!inBlockComment && !inInlineComment) {
        result += char;
      }
      continue;
    }

    if (char === '\\') {
      escaped = true;
      if (!inBlockComment && !inInlineComment) {
        result += char;
      }
      continue;
    }

    if (char === '"' && !inBlockComment && !inInlineComment) {
      inString = !inString;
      result += char;
      continue;
    }

    if (inString) {
      result += char;
      continue;
    }

    if (!inBlockComment && !inInlineComment && char === '/' && nextChar === '*') {
      inBlockComment = true;
      i++;
      continue;
    }

    if (inBlockComment && char === '*' && nextChar === '/') {
      inBlockComment = false;
      i++;
      continue;
    }

    if (!inBlockComment && !inInlineComment && char === '/' && nextChar === '/') {
      inInlineComment = true;
      i++;
      continue;
    }

    if (char === '\n' && inInlineComment) {
      inInlineComment = false;
      result += char;
      continue;
    }

    if (!inBlockComment && !inInlineComment) {
      result += char;
    }
  }

  return result.trim();
}

export function escapeHtml(text = '') {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

export function colToLetter(col) {
  let letter = '';
  let num = col;
  do {
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26) - 1;
  } while (num >= 0);
  return letter;
}

export function letterToCol(letter) {
  let col = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function cellRefToAddress(cellRef) {
  if (!cellRef) return null;
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  const col = letterToCol(match[1]);
  const row = parseInt(match[2], 10) - 1;
  if (Number.isNaN(row) || Number.isNaN(col)) return null;
  return [row, col];
}

export function addressToCellRef(row, col) {
  return `${colToLetter(col)}${row + 1}`;
}

export function isInsideQuotes(text, position) {
  if (!text) return false;
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

export function detectCellReferences(text) {
  if (!text) return [];

  const cellRefs = [];
  const cellPattern = /(?:'[^']+'!)?(?:(?:([A-Z]+\$?[0-9]+\$?)(?::([A-Z]+\$?[0-9]+\$?))*)|([A-Z]+)\$?:([A-Z]+)\$?)/gi;
  let match;

  while ((match = cellPattern.exec(text)) !== null) {
    if (isInsideQuotes(text, match.index)) {
      continue;
    }

    const fullMatch = match[0];
    let cellRef = fullMatch.includes('!') ? fullMatch.substring(fullMatch.indexOf('!') + 1) : fullMatch;

    if (match[1]) {
      cellRef = match[1];
      if (match[2]) {
        cellRef += `:${match[2]}`;
        const secondColonIndex = fullMatch.indexOf(':', fullMatch.indexOf(':') + 1);
        if (secondColonIndex !== -1) {
          cellRef += fullMatch.substring(secondColonIndex);
        }
      }
    } else if (match[3] && match[4]) {
      cellRef = `${match[3]}:${match[4]}`;
    }

    const hasDigits = /[0-9]/.test(cellRef);
    const hasColon = cellRef.includes(':');
    if (!hasDigits && !hasColon) {
      continue;
    }

    const startIndex = match.index;
    const endIndex = startIndex + fullMatch.length;

    cellRefs.push({
      text: cellRef,
      start: startIndex,
      end: endIndex,
      range: cellRef,
      fullMatch
    });
  }

  return cellRefs;
}

const CUSTOM_FUNCTIONS = [
  {
    name: 'LET',
    signature: 'LET(name1, value1, calculation_or_name2, ...)',
    description: 'Assigns names to calculation results for reuse within a formula.'
  }
];

export function extendWithCustomFunctions(functions = []) {
  const merged = Array.isArray(functions) ? [...functions] : [];
  const seen = new Set(merged.map((f) => f.name?.toUpperCase()).filter(Boolean));

  CUSTOM_FUNCTIONS.forEach((func) => {
    const upperName = func.name.toUpperCase();
    if (!seen.has(upperName)) {
      merged.push(func);
      seen.add(upperName);
    }
  });

  return merged;
}

/**
 * Base chip colors used across the experience (excluding pure black).
 * Organized list so consumers can reference consistent color tokens.
 */
export const BASE_CHIP_COLORS = [
  { name: 'Brick', hex: '#8C3B2A' },
  { name: 'Orchid', hex: '#9A5FA7' },
  { name: 'Indigo', hex: '#4B4C97' },
  { name: 'Harbor Green', hex: '#6B9B62' },
  { name: 'Citron', hex: '#D7DB8A' },
  { name: 'Charcoal', hex: '#777777' },
  { name: 'Fog', hex: '#BEBEBE' },
  { name: 'Olive Bark', hex: '#68551A' },
  { name: 'Copper', hex: '#A56C45' },
  { name: 'Rose', hex: '#C9877C' },
  { name: 'Lavender', hex: '#8C86C8' },
  { name: 'Mint', hex: '#B8DCAD' },
  { name: 'Aqua', hex: '#88C1CB' },
  { name: 'Slate', hex: '#8F8F8F' },
  { name: 'Snow', hex: '#FFFFFF' }
];

