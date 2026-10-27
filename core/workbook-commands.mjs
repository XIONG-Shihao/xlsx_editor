import { validateValueType } from './value-paste.mjs';
import { duplicateCommand } from './duplicate-model.mjs';
/** Portable command validation shared by the editor and standalone MCP.
 * A batch is expanded against a private projection before either owner accepts it.
 */
import { hasExactNumericValue } from './calculation.mjs';
import { normalizeSortKeys } from './sort-keys.mjs';
import { validateFilterCriteria } from './filter-model.mjs';
import { refusal } from './refusal-codes.mjs';
export function columnName(index) {
  let result = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  return result;
}

export function cellPosition(ref) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(String(ref));
  if (!match) throw refusal('internal-invalid-field', `Invalid cell address: ${ref}`);
  const col = [...match[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (row >= 1048576 || col >= 16384) throw refusal('workbook-limit', 'Cell is outside workbook limits');
  return { row, col };
}

export function rangeBounds(range) {
  const [first, last = first] = String(range).split(':');
  if (String(range).split(':').length > 2) throw refusal('internal-invalid-field', 'Invalid range');
  const a = cellPosition(first),
    b = cellPosition(last);
  if (a.row > b.row || a.col > b.col) throw refusal('internal-invalid-field', 'Invalid range');
  return {
    first: a,
    last: b,
    rows: b.row - a.row + 1,
    columns: b.col - a.col + 1,
    size: (b.row - a.row + 1) * (b.col - a.col + 1),
  };
}

export function rangeCells(range) {
  const { first: a, last: b } = rangeBounds(range);
  const result = [];
  for (let row = a.row; row <= b.row; row++)
    for (let col = a.col; col <= b.col; col++) result.push(`${columnName(col)}${row + 1}`);
  return result;
}

export function scalarCell(cell) {
  return cell?.value ?? null;
}

/** Authored literals only. Formula copying uses its existing structural path;
 * no stored/current derived output becomes a literal through this helper. */
export function literalFillSnapshot(cell) {
  if (
    cell &&
    (cell.formula || cell.formulaType || cell.formulaOutput || cell.spill || cell.scenarioResult || cell.calculation)
  )
    throw refusal('formula-unsupported', 'Source formula output copy is not qualified');
  const value = cell?.value ?? null;
  if (typeof value === 'number' && (!Number.isFinite(value) || !hasExactNumericValue(value, cell?.rawValue)))
    throw refusal('workbook-unsupported', 'Fill cannot copy a numeric value whose stored precision would change');
  return validateValueType({ value, ...(cell?.dataType === 'error' ? { valueType: 'error' } : {}) });
}

export function keyCellValue(cell) {
  if (cell?.formula || cell?.formulaType)
    throw refusal('workbook-bad-input', 'Business keys must be literal identifiers');
  if (typeof cell?.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue))
    throw refusal(
      'workbook-unsupported',
      'Business key loses stored numeric precision; use a qualified text identifier'
    );
  return scalarCell(cell);
}

export function assertXmlText(value) {
  // With the Unicode flag, valid surrogate pairs are single code points and do
  // not match the lone-surrogate range. Tabs, LF and CR are legal XML text.
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value))
    throw refusal('package-unsupported', 'Unsupported XML character in workbook text');
}

function assertGuard(cell, edit) {
  if (Object.hasOwn(edit, 'expectedValue') && !Object.is(scalarCell(cell), edit.expectedValue))
    throw refusal('workbook-refused', `Conflict: unexpected value in ${edit.sheet}!${edit.cell}`);
  if (
    Object.hasOwn(edit, 'expectedFormula') &&
    (cell?.formula?.replace(/^=/, '') ?? null) !== (edit.expectedFormula?.replace(/^=/, '') ?? null)
  )
    throw refusal('formula-refused', `Conflict: unexpected formula in ${edit.sheet}!${edit.cell}`);
  if (Object.hasOwn(edit, 'expectedRawValue') && (cell?.rawValue ?? null) !== edit.expectedRawValue)
    throw refusal('workbook-refused', `Conflict: unexpected stored value in ${edit.sheet}!${edit.cell}`);
}

function validateCellEdit(edit) {
  validateValueType(edit);
  const ref = String(edit.cell).replaceAll('$', '').toUpperCase();
  cellPosition(ref);
  if (
    ![
      'value',
      'formula',
      'bold',
      'italic',
      'underline',
      'fontFamily',
      'fontSize',
      'fontColor',
      'fillColor',
      'horizontalAlign',
      'verticalAlign',
      'wrapText',
      'border',
      'numberFormat',
      'styleId',
    ].some((key) => edit[key] !== undefined)
  )
    throw refusal('workbook-bad-input', 'Cell edit has no change');
  if (edit.styleId !== undefined && (!Number.isInteger(edit.styleId) || edit.styleId < 0))
    throw refusal('internal-invalid-field', 'Invalid source style');
  if (edit.value !== undefined && edit.formula !== undefined)
    throw refusal('formula-bad-input', 'Pass a literal value or a formula, not both');
  if (edit.value !== undefined && edit.value !== null && !['string', 'number', 'boolean'].includes(typeof edit.value))
    throw refusal('internal-invalid-field', 'Invalid cell value');
  if (typeof edit.value === 'number' && !Number.isFinite(edit.value))
    throw refusal('workbook-bad-input', 'Cell numbers must be finite');
  if (typeof edit.value === 'string' && edit.value.length > 32767)
    throw refusal('workbook-limit', 'Cell text exceeds 32767 characters');
  if (typeof edit.value === 'string') assertXmlText(edit.value);
  if (typeof edit.formula === 'string') assertXmlText(edit.formula);
  if (
    edit.formula !== undefined &&
    (typeof edit.formula !== 'string' ||
      !edit.formula.replace(/^=/, '') ||
      edit.formula.length > 8192 ||
      /https?:|file:|(?:DDE|WEBSERVICE|HYPERLINK)\s*\(|\[[^\]]*\.(?:xlsx?|xlsm)\]/i.test(edit.formula))
  )
    throw refusal('formula-unsupported', 'External or invalid formulas are unsupported');
  for (const key of ['bold', 'italic', 'underline', 'wrapText'])
    if (edit[key] !== undefined && typeof edit[key] !== 'boolean')
      throw refusal('internal-invalid-field', `Invalid ${key}`);
  if (
    edit.fontFamily !== undefined &&
    (typeof edit.fontFamily !== 'string' || !edit.fontFamily.trim() || edit.fontFamily.length > 100)
  )
    throw refusal('internal-invalid-field', 'Invalid font family');
  if (edit.fontFamily !== undefined) assertXmlText(edit.fontFamily);
  if (edit.fontSize !== undefined && (!Number.isFinite(edit.fontSize) || edit.fontSize < 6 || edit.fontSize > 72))
    throw refusal('internal-invalid-field', 'Invalid font size');
  for (const key of ['fontColor', 'fillColor'])
    if (edit[key] !== undefined && !/^[0-9a-fA-F]{6}$/.test(edit[key]))
      throw refusal('internal-invalid-field', `Invalid ${key}`);
  if (edit.horizontalAlign !== undefined && !['left', 'center', 'right'].includes(edit.horizontalAlign))
    throw refusal('internal-invalid-field', 'Invalid horizontal alignment');
  if (edit.verticalAlign !== undefined && !['top', 'center', 'bottom'].includes(edit.verticalAlign))
    throw refusal('internal-invalid-field', 'Invalid vertical alignment');
  if (edit.border !== undefined && !['none', 'all'].includes(edit.border))
    throw refusal('internal-invalid-field', 'Invalid border');
  if (
    edit.numberFormat !== undefined &&
    ![
      'general',
      'number',
      'percent',
      'date',
      'text',
      'currency',
      'accounting',
      'thousands',
      'decimal0',
      'decimal1',
      'decimal2',
    ].includes(edit.numberFormat)
  )
    throw refusal('style-unsupported', 'Unsupported number format');
  return { ...edit, cell: ref };
}

/** Excel's UI shorthand is not the persisted OOXML spelling. Normalize only
 * single-column current-row selectors with a proven table context. Quotes and
 * other selectors are left intact; ambiguous/complex shorthand rejects before
 * publication. This metadata-only step is shared by both workbook owners. */
export function canonicalizeTableShorthand(formula, book, owner, reference) {
  if (!formula.includes('[@')) return formula;
  const position = cellPosition(reference),
    pieces = [];
  let copied = 0;
  for (let i = 0; i < formula.length; i++) {
    if (formula[i] === '"' || formula[i] === "'") {
      const quote = formula[i];
      for (i++; i < formula.length; i++)
        if (formula[i] === quote) {
          if (formula[i + 1] === quote) i++;
          else break;
        }
      continue;
    }
    if (formula[i] !== '[') continue;
    let end = i,
      depth = 0;
    for (; end < formula.length; end++) {
      if (formula[end] === "'") {
        end++;
        continue;
      }
      if (formula[end] === '[') depth++;
      if (formula[end] === ']' && --depth === 0) break;
    }
    const spec = formula.slice(i, end + 1);
    if (!spec.startsWith('[@')) {
      i = end;
      continue;
    }
    const selector = /^\[@(?:\[([^\[\]]+)\]|([^\[\]]+))\]$/.exec(spec);
    const column = selector?.[1] ?? selector?.[2];
    if (!column?.trim() || /['#@,:\r\n]/.test(column))
      throw refusal('table-unsupported', 'Complex table shorthand requires qualified serialization');
    let start = i;
    while (start > 0 && /[A-Za-z0-9_.]/.test(formula[start - 1])) start--;
    const name = formula.slice(start, i),
      prefix = formula.slice(0, start).trimEnd();
    if (
      start < copied ||
      (name && !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) ||
      (prefix && !/[=+\-*/^&%,;(<>{}]$/.test(prefix))
    )
      throw refusal('table-unsupported', 'Unsupported table shorthand qualifier');
    const candidates = (book.tables ?? []).filter((table) => {
      if (name) return table.name.toLowerCase() === name.toLowerCase();
      if (table.sheet !== owner) return false;
      const { first, last } = rangeBounds(table.ref);
      return (
        position.row > first.row && position.row <= last.row && position.col >= first.col && position.col <= last.col
      );
    });
    if (candidates.length !== 1) throw refusal('table-bad-input', 'Table shorthand needs one unambiguous table');
    const table = candidates[0],
      { first, last } = rangeBounds(table.ref);
    if (
      !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(table.name) ||
      table.sheet !== owner ||
      (table.headerRowCount ?? 1) !== 1 ||
      (table.totalsRowCount ?? 0) !== 0 ||
      position.row <= first.row ||
      position.row > last.row
    )
      throw refusal('table-unsupported', 'Table shorthand is outside a qualified data row');
    const columns = table.columns.filter(
      (item) => (typeof item === 'string' ? item : item.name).toLowerCase() === column.toLowerCase()
    );
    if (columns.length !== 1 || table.columns.length !== last.col - first.col + 1)
      throw refusal('table-bad-input', 'Table shorthand needs one valid column');
    const header = typeof columns[0] === 'string' ? columns[0] : columns[0].name;
    if (/[\[\]'#@,:\r\n]/.test(header))
      throw refusal('table-unsupported', 'Complex table header requires qualified serialization');
    pieces.push(formula.slice(copied, start), `${table.name}[[#This Row],[${header}]]`);
    copied = end + 1;
    i = end;
  }
  if (!copied) return formula;
  const canonical = pieces.join('') + formula.slice(copied);
  if (canonical.length > 8192) throw refusal('table-limit', 'Canonical table formula exceeds the formula length limit');
  return canonical;
}

export function resolveWorkbookEdits(book, commands) {
  if (
    Array.isArray(commands) &&
    commands.some((command) =>
      [
        'setSheetBackground',
        'removeSheetBackground',
        'applyPageTheme',
        'booleanCell',
        'setHyperlink',
        'removeHyperlink',
      ].includes(command?.type)
    )
  )
    throw refusal('hyperlink-bad-input', 'Boolean-cell and hyperlink authoring require the native XLSX owner.');
  if (Array.isArray(commands) && commands.some((command) => command?.type === 'clearHyperlinks'))
    throw refusal(
      'hyperlink-unsupported',
      'Clear Hyperlinks requires the native XLSX owner; portable cell projection is unsupported.'
    );
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  if (book.readOnly) throw refusal('workbook-protected', `Read-only workbook: ${book.limitations.join('; ')}`);
  const sheets = new Map(book.sheets.map((sheet) => [sheet.name, { ...sheet, cells: sheet.cells, writes: new Map() }]));
  const edits = [];
  const validateSheetName = (name) => {
    if (
      typeof name !== 'string' ||
      !name.trim() ||
      name.length > 31 ||
      /[\\/:?*\[\]]/.test(name) ||
      name.startsWith("'") ||
      name.endsWith("'")
    )
      throw refusal('internal-invalid-field', 'Invalid worksheet name');
    assertXmlText(name);
  };
  function append(raw) {
    if (['addSheet', 'renameSheet', 'duplicateSheet'].includes(raw.operation)) {
      const destination = raw.operation === 'addSheet' ? raw.sheet : raw.newName;
      validateSheetName(destination);
      if (raw.operation !== 'addSheet' && !sheets.has(raw.sheet))
        throw refusal('workbook-not-found', `Sheet not found: ${raw.sheet}`);
      if (
        [...sheets.keys()].some(
          (name) => name.toLocaleLowerCase() === destination.toLocaleLowerCase() && name !== raw.sheet
        )
      )
        throw refusal('workbook-duplicate', `Worksheet already exists: ${destination}`);
      if (raw.operation === 'addSheet' && sheets.has(raw.sheet))
        throw refusal('workbook-duplicate', `Worksheet already exists: ${raw.sheet}`);
      if (raw.operation === 'duplicateSheet' && destination.toLocaleLowerCase() === raw.sheet.toLocaleLowerCase())
        throw refusal('workbook-duplicate', `Worksheet already exists: ${destination}`);
      if (raw.operation === 'renameSheet' && destination === raw.sheet)
        throw refusal('name-unsupported', 'Worksheet name is unchanged');
      edits.push({ ...raw });
      return;
    }
    const sheet = sheets.get(raw.sheet);
    if (!sheet) throw refusal('workbook-not-found', `Sheet not found: ${raw.sheet}`);
    if (raw.operation === 'removeDuplicates') {
      edits.push(duplicateCommand(raw));
      return;
    }
    if (raw.operation) {
      if (
        ![
          'insertRow',
          'deleteRow',
          'insertColumn',
          'deleteColumn',
          'sortRows',
          'filterRows',
          'reapplyFilter',
          'clearFilter',
          'removeFilter',
          'resizeColumn',
          'resizeRow',
          'freeze',
          'unhideAll',
          'mergeCells',
          'unmergeCells',
          'addConditionalRule',
          'createTable',
        ].includes(raw.operation)
      )
        throw refusal('workbook-unsupported', `Unsupported workbook operation: ${raw.operation}`);
      if (raw.operation === 'unhideAll' && Object.keys(raw).sort().join(',') !== 'operation,sheet')
        throw refusal('workbook-bad-input', 'Unhide Rows and Columns requires only operation and sheet');
      if (raw.operation === 'sortRows') normalizeSortKeys(raw);
      if (raw.operation === 'filterRows') validateFilterCriteria(raw.range, raw.column, raw.values);
      if (
        ['reapplyFilter', 'clearFilter', 'removeFilter'].includes(raw.operation) &&
        (Object.keys(raw).length !== 3 ||
          !Object.hasOwn(raw, 'expectedFilterRevision') ||
          !/^[a-f0-9]{64}$/.test(raw.expectedFilterRevision))
      )
        throw refusal('internal-invalid-field', 'Invalid accepted filter command fields');
      if (raw.operation === 'createTable') {
        const bounds = rangeBounds(raw.range);
        if (
          bounds.rows < 2 ||
          bounds.size > 10000 ||
          typeof raw.tableName !== 'string' ||
          !/^[A-Za-z_][A-Za-z0-9_.]{0,254}$/.test(raw.tableName) ||
          /^[A-Za-z]{1,3}[1-9]\d*$/.test(raw.tableName)
        )
          throw refusal('internal-invalid-field', 'Invalid table name or range');
        assertXmlText(raw.tableName);
      }
      if (raw.operation === 'addConditionalRule') {
        if (
          rangeBounds(raw.range).size > 10000 ||
          !['greaterThan', 'lessThan', 'duplicateValues'].includes(raw.rule) ||
          !/^[0-9a-fA-F]{6}$/.test(raw.fillColor) ||
          (raw.rule !== 'duplicateValues' && !Number.isFinite(raw.threshold))
        )
          throw refusal('internal-invalid-field', 'Invalid or oversized conditional formatting rule');
      }
      if (['mergeCells', 'unmergeCells'].includes(raw.operation) && rangeBounds(raw.range).size < 2)
        throw refusal('merge-bad-input', 'Select at least two cells to merge or unmerge');
      if (['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'].includes(raw.operation)) {
        if (
          !Number.isInteger(raw.index) ||
          raw.index < 0 ||
          raw.index >= (raw.operation.endsWith('Row') ? 1048576 : 16384)
        )
          throw refusal('workbook-malformed', 'Invalid row or column index');
        if (book.structuralReason) throw refusal('workbook-refused', book.structuralReason);
      }
      edits.push({
        ...raw,
        ...(raw.operation === 'sortRows' && Object.hasOwn(raw, 'keys')
          ? { keys: normalizeSortKeys(raw).map((key) => ({ ...key })) }
          : {}),
      });
      // Structural batches are resolved by the preserving writer, which knows all
      // workbook relationships. Do not manufacture a partial projected shift here.
      return;
    }
    const edit = validateCellEdit(raw);
    if (edit.formula !== undefined)
      edit.formula = canonicalizeTableShorthand(edit.formula, book, edit.sheet, edit.cell);
    const previous = sheet.writes.get(edit.cell) ?? sheet.cells[edit.cell] ?? { value: null };
    assertGuard(previous, edit);
    sheet.writes.set(edit.cell, {
      ...previous,
      ...edit,
      ...(edit.value !== undefined
        ? {
            formula: undefined,
            cachedValue: null,
            calculation: undefined,
            rawValue:
              typeof edit.value === 'number'
                ? String(edit.value)
                : typeof edit.value === 'boolean'
                  ? String(Number(edit.value))
                  : null,
            dataType:
              edit.valueType === 'error'
                ? 'error'
                : edit.value === null
                  ? 'blank'
                  : typeof edit.value === 'number'
                    ? 'number'
                    : typeof edit.value === 'boolean'
                      ? 'boolean'
                      : 'text',
          }
        : {}),
      ...(edit.formula !== undefined
        ? { formula: edit.formula.replace(/^=/, ''), rawValue: null, cachedValue: null, calculation: undefined }
        : {}),
    });
    edits.push(edit);
  }
  for (const command of commands) {
    if (command.type === 'updateRows') {
      const sheet = sheets.get(command.sheet);
      if (!sheet) throw refusal('workbook-not-found', `Sheet not found: ${command.sheet}`);
      const bounds = rangeBounds(command.range);
      const keyColumn = cellPosition(`${command.keyColumn}1`).col;
      if (keyColumn < bounds.first.col || keyColumn > bounds.last.col)
        throw refusal('workbook-unsupported', 'Key column is outside the requested range');
      const keys = new Map();
      for (let row = bounds.first.row; row <= bounds.last.row; row++) {
        const ref = `${columnName(keyColumn)}${row + 1}`;
        const value = keyCellValue(sheet.writes.get(ref) ?? sheet.cells[ref]);
        if (keys.has(value)) throw refusal('workbook-duplicate', `Duplicate key: ${String(value)}`);
        keys.set(value, cellPosition(ref).row + 1);
      }
      const requested = new Set();
      if (!Array.isArray(command.updates) || !command.updates.length)
        throw refusal('workbook-bad-input', 'Keyed updates must not be empty');
      for (const update of command.updates) {
        if (requested.has(update.key))
          throw refusal('workbook-duplicate', `Duplicate requested key: ${String(update.key)}`);
        requested.add(update.key);
        const row = keys.get(update.key);
        if (!row) throw refusal('internal-missing-field', `Missing key: ${String(update.key)}`);
        if (!Array.isArray(update.cells) || !update.cells.length)
          throw refusal('workbook-bad-input', 'Keyed update cells must not be empty');
        for (const { column, ...cell } of update.cells) {
          if (!Object.hasOwn(cell, 'expectedValue') && !Object.hasOwn(cell, 'expectedFormula'))
            throw refusal('formula-bad-input', 'Keyed updates require expected old values or formulas');
          const ref = `${column}${row}`;
          const position = cellPosition(ref);
          if (position.col < bounds.first.col || position.col > bounds.last.col)
            throw refusal('workbook-unsupported', `Update is outside the keyed range: ${ref}`);
          append({ ...cell, sheet: command.sheet, cell: ref });
        }
      }
    } else if (command.type === 'recalculate') {
      continue;
    } else if (command.type === 'clearRange') {
      for (const cell of rangeCells(command.range)) append({ sheet: command.sheet, cell, value: null });
    } else if (command.type === 'formatRange') {
      for (const cell of rangeCells(command.range)) append({ ...command.format, sheet: command.sheet, cell });
    } else if (command.type === 'setCells') {
      if (!Array.isArray(command.cells)) throw refusal('workbook-bad-input', 'setCells requires cells');
      for (const cell of command.cells) append({ ...cell, sheet: command.sheet });
    } else if (command.type !== undefined)
      throw refusal('workbook-unsupported', `Unsupported workbook command: ${command.type}`);
    else append(command);
  }
  return edits;
}
