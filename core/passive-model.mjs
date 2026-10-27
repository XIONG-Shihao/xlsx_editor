/** Typed sparse cell projections for passive readers. Saved caches stay explicit. */
import { DOMParser } from '@xmldom/xmldom';
import { cellPosition, rangeBounds, columnName } from './workbook-commands.mjs';
import { PASSIVE_FORMAT_LIMITS as LIMITS } from './passive-container.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const errors = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};
const fail = (message) => {
  throw refusal('package-refused', 'Passive workbook: ' + message);
};
const nodes = (element) => Array.from(element.childNodes).filter((node) => node.nodeType === 1);
const text = (value, label) => {
  if (typeof value !== 'string' || value.length > LIMITS.text || value.includes('\0')) fail('invalid ' + label);
  return value;
};
function xml(bytes) {
  const bad = () => fail('invalid worksheet XML'),
    raw = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(raw)) bad();
  return new DOMParser({ errorHandler: { warning: bad, error: bad, fatalError: bad } }).parseFromString(
    raw,
    'application/xml'
  ).documentElement;
}
/** Capture original numeric lexemes and formula/cache presence before a parser
 * can round decimal text or conflate a missing result with an ordinary blank. */
function rawSheet(parts, name) {
  if (!parts) return null;
  const workbook = xml(parts['xl/workbook.xml']),
    sheets = nodes(workbook).filter((n) => n.namespaceURI === NS && n.localName === 'sheets');
  if (sheets.length !== 1) fail('ambiguous source worksheets');
  const sheet = nodes(sheets[0]).filter(
    (n) => n.namespaceURI === NS && n.localName === 'sheet' && n.getAttribute('name') === name
  );
  if (sheet.length !== 1) fail('source worksheet identity differs');
  const rel = xml(parts['xl/_rels/workbook.xml.rels']),
    id = sheet[0].getAttributeNS(R, 'id'),
    edges = nodes(rel).filter(
      (n) => n.namespaceURI === PR && n.localName === 'Relationship' && n.getAttribute('Id') === id
    );
  if (edges.length !== 1 || edges[0].getAttribute('TargetMode')) fail('source worksheet relationship differs');
  const target = edges[0].getAttribute('Target'),
    stack = target.startsWith('/') ? [] : ['xl'];
  for (const part of target.split('/'))
    if (part === '..') stack.pop();
    else if (part && part !== '.') stack.push(part);
  const path = stack.join('/');
  if (!parts[path]) fail('missing source worksheet');
  const root = xml(parts[path]);
  if (root.namespaceURI !== NS || root.localName !== 'worksheet') fail('invalid source worksheet');
  const result = new Map();
  for (const row of nodes(root)
    .filter((n) => n.namespaceURI === NS && n.localName === 'sheetData')
    .flatMap(nodes))
    for (const cell of nodes(row)) {
      if (cell.namespaceURI !== NS || cell.localName !== 'c') fail('extended source cells are unsupported');
      const ref = cell.getAttribute('r');
      cellPosition(ref);
      if (result.has(ref)) fail('duplicate source cell');
      const values = nodes(cell).filter((n) => n.namespaceURI === NS && n.localName === 'v'),
        formulas = nodes(cell).filter((n) => n.namespaceURI === NS && n.localName === 'f');
      if (values.length > 1 || formulas.length > 1) fail('ambiguous source formula/cache');
      result.set(ref, {
        type: cell.getAttribute('t') || 'n',
        hasValue: values.length === 1,
        raw: values[0]?.textContent,
        formula: formulas[0]?.textContent,
        formulaType: formulas[0]?.getAttribute('t') || undefined,
      });
      if (result.size > LIMITS.storedCells) fail('source cell budget exceeded');
    }
  return result;
}
export function passiveCell(source, revision, raw) {
  if (!source || typeof source !== 'object') fail('invalid parsed cell');
  let value = source.v,
    type = source.t,
    formula = raw?.formula === '' && raw.formulaType === 'shared' ? source.f : (raw?.formula ?? source.f);
  const formulaPresent = formula !== undefined || source.F !== undefined;
  let invalid = null,
    cacheError;
  if (raw && ['n', 'b', 'e', 'str'].includes(raw.type)) {
    type = raw.type === 'str' ? 's' : raw.type;
    if (!raw.hasValue) {
      value = undefined;
      if (formulaPresent) invalid = 'No stored formula result is present';
      else type = 'z';
    } else if (type === 'n') {
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw.raw ?? ''))
        invalid = 'Malformed stored numeric result';
      else value = Number(raw.raw);
    } else if (type === 'b') {
      if (!['0', '1'].includes(raw.raw)) invalid = 'Malformed stored boolean result';
      else value = raw.raw === '1';
    } else value = raw.raw;
  }
  if (type === 'e') {
    cacheError =
      typeof value === 'number'
        ? errors[value]
        : typeof value === 'string' && /^#[A-Z0-9/?!_]+$/.test(value)
          ? value
          : undefined;
    if (!cacheError) invalid = 'Unrecognized stored error result';
    value = cacheError ?? null;
  } else if (type === 'n') {
    if (typeof value !== 'number' || !Number.isFinite(value)) invalid = 'Malformed or nonfinite stored numeric result';
  } else if (type === 'b') {
    if (typeof value !== 'boolean') invalid = 'Malformed stored boolean result';
  } else if (['s', 'str'].includes(type)) {
    if (typeof value !== 'string') invalid = 'Malformed stored text result';
    else text(value, 'stored text');
  } else if (type === 'z' || (type === undefined && value === undefined)) value = null;
  else invalid = 'Unqualified stored cell type';
  const result = {
    value: invalid ? null : (value ?? null),
    dataType:
      type === 'n'
        ? 'number'
        : type === 'b'
          ? 'boolean'
          : type === 'e'
            ? 'error'
            : ['s', 'str'].includes(type)
              ? 'text'
              : 'blank',
    ...(typeof source.z === 'string' ? { format: text(source.z, 'number format') } : {}),
    ...(type === 'n' && raw?.raw !== undefined ? { rawValue: raw.raw } : {}),
  };
  if (formula !== undefined) {
    result.formula = text(formula, 'formula');
    if (!formula) invalid = 'Empty stored formula';
  }
  if (source.F !== undefined) result.formulaReference = text(source.F, 'array range');
  if (raw?.formulaType) result.formulaType = raw.formulaType;
  else if (source.F !== undefined) result.formulaType = 'array';
  if (formulaPresent) {
    if (invalid || value === undefined || type === 'z')
      result.calculation = { status: 'unresolved', error: invalid ?? 'No stored formula result is present', revision };
    else result.calculation = { status: 'stored', value, revision, ...(cacheError ? { error: cacheError } : {}) };
    if (!invalid) result.cachedValue = value;
  } else if (invalid) {
    result.error = invalid;
    result.dataType = 'error';
  }
  if (cacheError) result.error = cacheError;
  return result;
}
export function passiveWorkbookModel(book, { format, parts, revision }) {
  if (!Array.isArray(book?.SheetNames) || !book.SheetNames.length || book.SheetNames.length > LIMITS.sheets)
    fail('worksheet count exceeds its budget');
  const names = new Set(),
    sheets = [],
    records = new Map();
  let total = 0,
    stored = 0,
    unresolved = 0;
  for (let index = 0; index < book.SheetNames.length; index++) {
    const name = text(book.SheetNames[index], 'worksheet name');
    if (!name || name.length > 31 || names.has(name.toLowerCase()) || !Object.hasOwn(book.Sheets, name))
      fail('ambiguous worksheet identity');
    names.add(name.toLowerCase());
    const source = book.Sheets[name],
      raw = format === 'xlsm' ? rawSheet(parts, name) : null,
      entries = [],
      positions = new Set();
    let rows = 1,
      columns = 1,
      formulas = 0;
    if (source['!ref']) {
      const bounds = rangeBounds(source['!ref']);
      rows = bounds.last.row + 1;
      columns = bounds.last.col + 1;
    }
    for (const ref of Object.keys(source)) {
      if (ref.startsWith('!')) continue;
      const position = cellPosition(ref),
        cell = passiveCell(source[ref], revision, raw?.get(ref));
      entries.push({ ref, row: position.row, col: position.col, cell });
      positions.add(ref);
      delete source[ref];
      rows = Math.max(rows, position.row + 1);
      columns = Math.max(columns, position.col + 1);
      if (++total > LIMITS.storedCells) fail('stored cell budget exceeded');
      if (cell.formula !== undefined || cell.formulaType) {
        formulas++;
        if (cell.calculation?.status === 'stored') stored++;
        else unresolved++;
      }
    }
    if (raw && [...raw.keys()].some((ref) => !positions.has(ref))) fail('parser omitted a stored source cell');
    entries.sort((a, b) => a.row - b.row || a.col - b.col);
    const merges = (source['!merges'] ?? []).map((merge) => {
      if (
        !Number.isSafeInteger(merge.s?.r) ||
        !Number.isSafeInteger(merge.s?.c) ||
        !Number.isSafeInteger(merge.e?.r) ||
        !Number.isSafeInteger(merge.e?.c) ||
        merge.s.r < 0 ||
        merge.s.c < 0 ||
        merge.e.r < merge.s.r ||
        merge.e.c < merge.s.c ||
        merge.e.r > 1048575 ||
        merge.e.c > 16383
      )
        fail('invalid merged range');
      return columnName(merge.s.c) + (merge.s.r + 1) + ':' + columnName(merge.e.c) + (merge.e.r + 1);
    });
    if (merges.length > 10000) fail('merged range budget exceeded');
    const widths = [];
    for (let col = 0; col < (source['!cols']?.length ?? 0); col++) {
      if (col > 16383) fail('column metadata exceeds its budget');
      const item = source['!cols'][col];
      if (item) {
        const width = item.wch ?? item.width ?? (item.wpx === undefined ? undefined : item.wpx / 7);
        if (width !== undefined && (!Number.isFinite(width) || width < 0 || width > 1000)) fail('invalid column width');
        widths.push({ min: col + 1, max: col + 1, width: width ?? 8.43, hidden: !!item.hidden });
      }
    }
    const rowHeights = [],
      hiddenRows = [];
    for (let row = 0; row < (source['!rows']?.length ?? 0); row++) {
      if (row > 1048575) fail('row metadata exceeds its budget');
      const item = source['!rows'][row];
      if (item) {
        if (item.hidden) hiddenRows.push(row);
        const height = item.hpt ?? (item.hpx === undefined ? undefined : item.hpx * 0.75);
        if (height !== undefined) {
          if (!Number.isFinite(height) || height < 0 || height > 409.5) fail('invalid row height');
          rowHeights.push({ row, height });
        }
      }
    }
    const sheet = {
      name,
      state:
        book.Workbook?.Sheets?.[index]?.Hidden === 2
          ? 'veryHidden'
          : book.Workbook?.Sheets?.[index]?.Hidden === 1
            ? 'hidden'
            : 'visible',
      path: 'passive-sheet-' + index,
      rows,
      columns,
      cells: {},
      merges,
      widths,
      hiddenRows,
      rowHeights,
      storedCells: entries.length,
      storedFormulaCount: formulas,
      date1904: !!book.Workbook?.WBProps?.date1904,
    };
    sheets.push(sheet);
    records.set(name, entries);
    delete book.Sheets[name];
  }
  return { sheets, records, summary: { calculated: 0, stored, unresolved, error: 0, circular: 0 } };
}
