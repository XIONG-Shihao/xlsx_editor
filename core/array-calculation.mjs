/** Authored array context is a protection hint, never a spill/result proof. */
export const ARRAY_CALCULATION_REASON = 'Array/spill calculation requires qualified output ownership';
const producers = new Set([
  'ANCHORARRAY',
  'BYCOL',
  'BYROW',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'DROP',
  'EXPAND',
  'FILTER',
  'HSTACK',
  'MAKEARRAY',
  'MAP',
  'RANDARRAY',
  'SCAN',
  'SEQUENCE',
  'SORT',
  'SORTBY',
  'TAKE',
  'TEXTSPLIT',
  'TOCOL',
  'TOROW',
  'UNIQUE',
  'VSTACK',
  'WRAPCOLS',
  'WRAPROWS',
]);
const scalarConsumers = new Set([
  'SUM',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'PRODUCT',
  'SUMIF',
  'SUMIFS',
  'AVERAGEIF',
  'AVERAGEIFS',
  'COUNTIF',
  'COUNTIFS',
  'SUMPRODUCT',
  'ROWS',
  'COLUMNS',
  'AND',
  'OR',
  'XOR',
  'CONCAT',
  'CONCATENATE',
  'TEXTJOIN',
  'AGGREGATE',
  'SUBTOTAL',
  'VLOOKUP',
  'HLOOKUP',
  'MATCH',
]);
function scalarIndex(text, open) {
  const args = [];
  let start = open + 1,
    depth = 0,
    brackets = 0,
    quote = '';
  for (let i = start; i < text.length; i++) {
    const char = text[i];
    if (quote) {
      if (char === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '[' || char === '{') {
      brackets++;
      continue;
    }
    if (char === ']' || char === '}') {
      brackets--;
      continue;
    }
    if (brackets) continue;
    if (char === '(') depth++;
    else if (char === ')') {
      if (depth) depth--;
      else {
        args.push(text.slice(start, i));
        return args.length === 3 && args.slice(1).every((value) => /^\s*\+?[1-9]\d*\s*$/.test(value));
      }
    } else if (char === ',' && !depth) {
      args.push(text.slice(start, i));
      start = i + 1;
    }
  }
  return false;
}

export function arrayFormulaContext(formula) {
  const text = String(formula),
    frames = [];
  let quote = '',
    square = 0,
    conservative = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quote) {
      if (char === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '[') {
      square++;
      continue;
    }
    if (char === ']') {
      square = Math.max(0, square - 1);
      continue;
    }
    if (square) continue;
    if (char === '(') {
      let end = i - 1;
      while (end >= 0 && /\s/.test(text[end])) end--;
      let start = end;
      while (start >= 0 && /[A-Za-z0-9_.\\]/.test(text[start])) start--;
      const name = text
        .slice(start + 1, end + 1)
        .toUpperCase()
        .replace(/^(?:_XLFN\.|_XLWS\.)+/, '');
      if (producers.has(name)) return { blocked: true, conservative: true, reason: ARRAY_CALCULATION_REASON };
      const scalarCell = name === 'INDEX' && scalarIndex(text, i);
      if (
        ['INDEX', 'XLOOKUP', 'LOOKUP', 'TRANSPOSE', 'ROW', 'COLUMN'].includes(name) &&
        !scalarCell &&
        !frames.some((frame) => scalarConsumers.has(frame) || frame === 'INDEX_CELL')
      )
        conservative = true;
      frames.push(scalarCell ? 'INDEX_CELL' : name);
    } else if (char === ')') frames.pop();
    else if (char === '@' || (char === '#' && /[A-Za-z0-9_)\]]/.test(text.slice(0, i).trimEnd().at(-1) ?? '')))
      return { blocked: true, conservative: true, reason: ARRAY_CALCULATION_REASON };
    else if (
      (char === ':' || char === '{') &&
      !frames.some((frame) => scalarConsumers.has(frame) || frame === 'INDEX_CELL')
    )
      conservative = true;
  }
  return { blocked: false, conservative };
}
export function hasArrayContext(view) {
  for (const sheet of view.sheets)
    for (const cell of Object.values(sheet.cells))
      if (cell.formulaType === 'array' || (cell.formula && arrayFormulaContext(cell.formula).conservative)) return true;
  return (view.namedRanges ?? []).some((name) => arrayFormulaContext(name.formula).conservative);
}
export function hasUnqualifiedArrayResults(view) {
  return view.sheets.some((sheet) =>
    Object.values(sheet.cells).some((cell) => cell.calculation?.error?.startsWith(ARRAY_CALCULATION_REASON))
  );
}
