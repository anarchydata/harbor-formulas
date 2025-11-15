import { colToLetter, letterToCol, isInsideQuotes } from './helpers.js';

const NUMBER_REGEX = /^[+-]?\d+(?:\.\d+)?$/;
const DATE_CANDIDATE_REGEX = /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/;
const TEXT_NUMBER_REGEX = /^(.*?)(\d+)$/;
const ISO_DATE_REGEX = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const MDY_DATE_REGEX = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
const MS_PER_DAY = 24 * 60 * 60 * 1000;

const defaultLists = [
  createList('daysShort', ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']),
  createList('daysLong', ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']),
  createList('monthsShort', ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']),
  createList('monthsLong', ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])
].filter(Boolean);

let customLists = [];

export const AutofillDirection = Object.freeze({
  FORWARD: 'forward',
  BACKWARD: 'backward'
});

export function setCustomAutofillLists(lists = []) {
  if (!Array.isArray(lists)) {
    customLists = [];
    return;
  }

  customLists = lists
    .map((entry, index) => {
      if (Array.isArray(entry)) {
        return createList(`custom_${index}`, entry);
      }
      if (entry && Array.isArray(entry.values)) {
        return createList(entry.name || `custom_${index}`, entry.values);
      }
      return null;
    })
    .filter(Boolean);
}

export function detectValuePattern(seedsInput = []) {
  const seeds = Array.isArray(seedsInput) ? seedsInput.map((value) => (value ?? '').toString()) : [''];

  if (seeds.length === 0) {
    return { type: 'copy', seedValue: '' };
  }

  const trimmed = seeds.map((value) => value.trim());

  if (seeds.length === 1) {
    const singleDatePattern = detectSingleDateSeries(trimmed[0]);
    if (singleDatePattern) {
      return singleDatePattern;
    }
    const singleListPattern = detectKnownListSeries(seeds);
    if (singleListPattern) {
      return singleListPattern;
    }
    const singleTextNumberPattern = detectTextNumberSeries(seeds);
    if (singleTextNumberPattern) {
      return singleTextNumberPattern;
    }
    return { type: 'copy', seedValue: seeds[0] ?? '' };
  }

  if (trimmed.every((value) => NUMBER_REGEX.test(value))) {
    const numbers = trimmed.map((value) => Number(value));
    const step = numbers[1] - numbers[0];
    const consistent = numbers.every((value, idx, arr) => idx === 0 || nearlyEqual(value - arr[idx - 1], step));
    if (consistent) {
      const decimalPlaces = trimmed.reduce((max, value) => Math.max(max, countDecimalPlaces(value)), 0);
      return {
        type: 'numberSeries',
        numbers,
        firstValue: numbers[0],
        lastValue: numbers[numbers.length - 1],
        step,
        decimalPlaces
      };
    }
  }

  const datePattern = detectDateSeries(trimmed);
  if (datePattern) {
    return datePattern;
  }

  const textNumberPattern = detectTextNumberSeries(seeds);
  if (textNumberPattern) {
    return textNumberPattern;
  }

  const knownListPattern = detectKnownListSeries(seeds);
  if (knownListPattern) {
    return knownListPattern;
  }

  return { type: 'repeat', seeds: [...seeds] };
}

export function generateValueFill(seedsInput = [], count = 0, options = {}) {
  const seeds = Array.isArray(seedsInput) ? seedsInput.map((value) => (value ?? '').toString()) : [''];
  const pattern = detectValuePattern(seeds);
  const direction = options.direction === AutofillDirection.BACKWARD ? AutofillDirection.BACKWARD : AutofillDirection.FORWARD;

  if (count <= 0) {
    return [];
  }

  switch (pattern.type) {
    case 'numberSeries':
      return generateNumberSeries(pattern, count, direction);
    case 'dateSeries':
      return generateDateSeries(pattern, count, direction);
    case 'textIncrement':
      return generateTextNumberSeries(pattern, count, direction);
    case 'knownList':
      return generateKnownListSeries(pattern, count, direction);
    case 'repeat':
      return generateRepeatSeries(pattern, count, direction);
    case 'copy':
    default:
      return Array.from({ length: count }, () => pattern.seedValue ?? seeds[0] ?? '');
  }
}

export function adjustFormulaReferences(formulaText, rowDelta = 0, colDelta = 0) {
  if (typeof formulaText !== 'string' || !formulaText.startsWith('=')) {
    return formulaText ?? '';
  }

  const referencePattern = /\b(\$?)([A-Z]+)(\$?)(\d+)\b/gi;
  return formulaText.replace(referencePattern, (match, colAbs, colLetters, rowAbs, rowDigits, offset) => {
    if (isInsideQuotes(formulaText, offset)) {
      return match;
    }

    const originalColIndex = letterToCol(colLetters);
    const originalRowIndex = parseInt(rowDigits, 10) - 1;
    if (Number.isNaN(originalColIndex) || Number.isNaN(originalRowIndex)) {
      return match;
    }

    let nextColIndex = originalColIndex;
    let nextRowIndex = originalRowIndex;

    if (!colAbs) {
      nextColIndex += colDelta;
    }
    if (!rowAbs) {
      nextRowIndex += rowDelta;
    }

    if (nextColIndex < 0 || nextRowIndex < 0) {
      return match;
    }

    const nextColLetters = colToLetter(nextColIndex);
    const rowPart = `${rowAbs ? '$' : ''}${nextRowIndex + 1}`;
    return `${colAbs ? '$' : ''}${nextColLetters}${rowPart}`;
  });
}

function createList(name, values) {
  if (!Array.isArray(values) || values.length === 0) {
    return null;
  }
  const normalized = values.map((value) => (value ?? '').toString());
  return {
    name,
    values: normalized,
    lowerValues: normalized.map((value) => value.toLowerCase())
  };
}

function nearlyEqual(a, b, epsilon = 1e-9) {
  return Math.abs(a - b) <= epsilon;
}

function countDecimalPlaces(value) {
  const match = value.match(/\.(\d+)/);
  return match ? match[1].length : 0;
}

function detectDateSeries(values) {
  if (values.length < 2 || !values.every((value) => DATE_CANDIDATE_REGEX.test(value))) {
    return null;
  }

  const parsed = values.map(parseSupportedDate);
  if (parsed.some((entry) => !entry)) {
    return null;
  }

  const formatSignature = parsed[0].format.signature;
  if (!parsed.every((entry) => entry.format.signature === formatSignature)) {
    return null;
  }

  const serials = parsed.map((entry) => entry.serial);
  const step = serials[1] - serials[0];
  const consistent = serials.every((value, idx, arr) => idx === 0 || value - arr[idx - 1] === step);
  if (!consistent) {
    return null;
  }

  return {
    type: 'dateSeries',
    firstSerial: serials[0],
    lastSerial: serials[serials.length - 1],
    step,
    format: parsed[0].format
  };
}

function detectSingleDateSeries(value) {
  if (!value || !DATE_CANDIDATE_REGEX.test(value)) {
    return null;
  }

  const parsed = parseSupportedDate(value);
  if (!parsed) {
    return null;
  }

  return {
    type: 'dateSeries',
    firstSerial: parsed.serial,
    lastSerial: parsed.serial,
    step: 1,
    format: parsed.format
  };
}

function parseSupportedDate(value) {
  const trimmed = value.trim();
  let match = trimmed.match(ISO_DATE_REGEX);
  if (match) {
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    const day = parseInt(match[3], 10);
    if (!isValidDateParts(year, month, day)) {
      return null;
    }
    const serial = Date.UTC(year, month - 1, day) / MS_PER_DAY;
    return {
      serial,
      format: {
        type: 'ymd',
        separator: '-',
        order: ['Y', 'M', 'D'],
        lengths: [match[1].length, match[2].length, match[3].length],
        signature: `ymd-${match[1].length}-${match[2].length}-${match[3].length}`
      }
    };
  }

  match = trimmed.match(MDY_DATE_REGEX);
  if (match) {
    const month = parseInt(match[1], 10);
    const day = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);
    if (!isValidDateParts(year, month, day)) {
      return null;
    }
    const serial = Date.UTC(year, month - 1, day) / MS_PER_DAY;
    return {
      serial,
      format: {
        type: 'mdy',
        separator: '/',
        order: ['M', 'D', 'Y'],
        lengths: [match[1].length, match[2].length, match[3].length],
        signature: `mdy-${match[1].length}-${match[2].length}-${match[3].length}`
      }
    };
  }

  return null;
}

function isValidDateParts(year, month, day) {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return false;
  }
  if (month < 1 || month > 12) {
    return false;
  }
  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return day >= 1 && day <= daysInMonth;
}

function detectTextNumberSeries(seeds) {
  if (seeds.length === 0) {
    return null;
  }

  const parsed = seeds.map((value) => {
    const asString = (value ?? '').toString();
    const leadingWhitespace = asString.match(/^\s*/)?.[0] ?? '';
    const trailingWhitespace = asString.match(/\s*$/)?.[0] ?? '';
    const trimmed = asString.trim();
    return {
      match: trimmed.match(TEXT_NUMBER_REGEX),
      leadingWhitespace,
      trailingWhitespace
    };
  });

  if (parsed.some((entry) => !entry.match || !entry.match[1])) {
    return null;
  }

  const prefix = parsed[0].match[1];
  if (!parsed.every((entry) => entry.match[1] === prefix)) {
    return null;
  }

  const numbers = parsed.map((entry) => parseInt(entry.match[2], 10));
  if (numbers.some((num) => Number.isNaN(num))) {
    return null;
  }

  const step = numbers.length >= 2 ? numbers[1] - numbers[0] : 1;
  const consistent = numbers.length <= 2 || numbers.every((value, idx, arr) => idx === 0 || value - arr[idx - 1] === step);
  if (!consistent) {
    return null;
  }

  const padLength = Math.max(...parsed.map((entry) => entry.match[2].length));
  return {
    type: 'textIncrement',
    prefix,
    numbers,
    firstNumber: numbers[0],
    lastNumber: numbers[numbers.length - 1],
    step,
    padLength,
    leadingWhitespace: parsed[0].leadingWhitespace || '',
    trailingWhitespace: parsed[0].trailingWhitespace || ''
  };
}

function detectKnownListSeries(seeds) {
  const parsed = seeds.map((value) => {
    const asString = (value ?? '').toString();
    const leadingWhitespace = asString.match(/^\s*/)?.[0] ?? '';
    const trailingWhitespace = asString.match(/\s*$/)?.[0] ?? '';
    const trimmed = asString.trim();
    return {
      lower: trimmed.toLowerCase(),
      trimmed,
      leadingWhitespace,
      trailingWhitespace
    };
  });
  const trimmedLower = parsed.map((entry) => entry.lower);
  const lists = [...defaultLists, ...customLists];

  for (const list of lists) {
    const indexes = trimmedLower.map((value) => list.lowerValues.indexOf(value));
    if (indexes.some((idx) => idx === -1)) {
      continue;
    }

    const sequential = indexes.every((value, idx) => {
      if (idx === 0) return true;
      const expected = (indexes[idx - 1] + 1) % list.values.length;
      return value === expected;
    });

    if (sequential) {
      return {
        type: 'knownList',
        list,
        firstIndex: indexes[0],
        lastIndex: indexes[indexes.length - 1],
        casingSample: parsed[0].trimmed,
        leadingWhitespace: parsed[0].leadingWhitespace || '',
        trailingWhitespace: parsed[0].trailingWhitespace || ''
      };
    }
  }

  return null;
}

function generateNumberSeries(pattern, count, direction) {
  const results = [];
  if (direction === AutofillDirection.FORWARD) {
    for (let i = 1; i <= count; i++) {
      const nextValue = pattern.lastValue + pattern.step * i;
      results.push(formatNumber(nextValue, pattern.decimalPlaces));
    }
  } else {
    for (let i = 1; i <= count; i++) {
      const nextValue = pattern.firstValue - pattern.step * i;
      results.push(formatNumber(nextValue, pattern.decimalPlaces));
    }
  }
  return results;
}

function formatNumber(value, decimalPlaces) {
  if (decimalPlaces > 0) {
    return value.toFixed(decimalPlaces);
  }
  return Number.isInteger(value) ? value.toString() : value.toString();
}

function generateDateSeries(pattern, count, direction) {
  const results = [];
  if (direction === AutofillDirection.FORWARD) {
    for (let i = 1; i <= count; i++) {
      const serial = pattern.lastSerial + pattern.step * i;
      results.push(formatDateFromSerial(serial, pattern.format));
    }
  } else {
    for (let i = 1; i <= count; i++) {
      const serial = pattern.firstSerial - pattern.step * i;
      results.push(formatDateFromSerial(serial, pattern.format));
    }
  }
  return results;
}

function formatDateFromSerial(serial, format) {
  const timestamp = serial * MS_PER_DAY;
  const date = new Date(timestamp);

  const parts = {
    Y: date.getUTCFullYear(),
    M: date.getUTCMonth() + 1,
    D: date.getUTCDate()
  };

  const segments = format.order.map((token, index) => {
    const rawValue = parts[token];
    const minLength = format.lengths[index] || (token === 'Y' ? 4 : 2);
    return rawValue.toString().padStart(minLength, '0');
  });

  return segments.join(format.separator);
}

function generateTextNumberSeries(pattern, count, direction) {
  const results = [];
  const leadingWhitespace = pattern.leadingWhitespace || '';
  const trailingWhitespace = pattern.trailingWhitespace || '';
  if (direction === AutofillDirection.FORWARD) {
    for (let i = 1; i <= count; i++) {
      const nextNumber = pattern.lastNumber + pattern.step * i;
      results.push(`${leadingWhitespace}${pattern.prefix}${padNumber(nextNumber, pattern.padLength)}${trailingWhitespace}`);
    }
  } else {
    for (let i = 1; i <= count; i++) {
      const nextNumber = pattern.firstNumber - pattern.step * i;
      results.push(`${leadingWhitespace}${pattern.prefix}${padNumber(nextNumber, pattern.padLength)}${trailingWhitespace}`);
    }
  }
  return results;
}

function padNumber(value, length) {
  const isNegative = value < 0;
  const absolute = Math.abs(value);
  const padded = absolute.toString().padStart(length, '0');
  return isNegative ? `-${padded}` : padded;
}

function generateKnownListSeries(pattern, count, direction) {
  const values = pattern.list.values;
  const length = values.length;
  const results = [];
  const leadingWhitespace = pattern.leadingWhitespace || '';
  const trailingWhitespace = pattern.trailingWhitespace || '';
  if (direction === AutofillDirection.FORWARD) {
    for (let i = 1; i <= count; i++) {
      const index = (pattern.lastIndex + i) % length;
      const nextValue = applyCasing(values[index], pattern.casingSample);
      results.push(`${leadingWhitespace}${nextValue}${trailingWhitespace}`);
    }
  } else {
    for (let i = 1; i <= count; i++) {
      const index = ((pattern.firstIndex - i) % length + length) % length;
      const nextValue = applyCasing(values[index], pattern.casingSample);
      results.push(`${leadingWhitespace}${nextValue}${trailingWhitespace}`);
    }
  }
  return results;
}

function applyCasing(value, sample) {
  if (!sample) {
    return value;
  }
  if (sample === sample.toUpperCase()) {
    return value.toUpperCase();
  }
  if (sample === sample.toLowerCase()) {
    return value.toLowerCase();
  }
  if (
    sample[0] &&
    sample[0] === sample[0].toUpperCase() &&
    sample.slice(1) === sample.slice(1).toLowerCase()
  ) {
    return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
  }
  return value;
}

function generateRepeatSeries(pattern, count, direction) {
  const { seeds } = pattern;
  const length = seeds.length || 1;
  if (direction === AutofillDirection.FORWARD) {
    return Array.from({ length: count }, (_, idx) => seeds[idx % length] ?? '');
  }
  return Array.from({ length: count }, (_, idx) => {
    const offset = (idx % length) + 1;
    const index = (length - offset) % length;
    return seeds[index] ?? '';
  });
}

