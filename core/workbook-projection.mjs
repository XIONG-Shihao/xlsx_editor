import { normalizeValuePasteCell, createValuePasteSourceContext } from './value-paste.mjs';
import {
  resolveWorkbookEdits,
  literalFillSnapshot,
  cellPosition,
  columnName,
  rangeCells,
  rangeBounds as parseRange,
} from './workbook-commands.mjs';
import { hasExactNumericValue } from './calculation.mjs';
import { normalizeSortKeys } from './sort-keys.mjs';
import { prepareFilterState, filterDefinition, baselineHidden } from './filter-model.mjs';
import { refusal } from './refusal-codes.mjs';
const address = (row, col) => columnName(col) + (row + 1);
const own = (obj, key) => Object.hasOwn(obj, key);
function rangeBounds(range, includeCells = false) {
  return { ...parseRange(range), ...(includeCells ? { cells: rangeCells(range) } : {}) };
}
function scalar(value) {
  return value === null || ['string', 'number', 'boolean'].includes(typeof value);
}
function rejectSpecial(view, sheet) {
  if (view.readOnly) throw refusal('workbook-protected', 'Workbook is read-only');
  if (Object.values(sheet.cells).some((cell) => cell.formulaType))
    throw refusal('formula-unsupported', 'Special formulas require a qualified spreadsheet operation');
}
function setCell(sheet, edit) {
  const old = sheet.cells[edit.cell] ?? { value: null },
    cell = { ...old };
  if (old.formulaType && (own(edit, 'value') || own(edit, 'formula')))
    throw refusal('table-unsupported', 'Editing a shared, array or data-table formula cell is not qualified');
  if (edit.styleId === 0) {
    for (const key of [
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
    ])
      delete cell[key];
    cell.format = 0;
  }
  if (own(edit, 'value')) {
    cell.value = edit.value;
    cell.formula = undefined;
    cell.formulaType = undefined;
    cell.calculation = undefined;
    delete cell.cachedResult;
    cell.cachedValue = null;
    delete cell.error;
    delete cell.formulaOutput;
    delete cell.scenarioResult;
    cell.rawValue = edit.value === null ? null : String(edit.value);
    cell.dataType =
      edit.valueType === 'error'
        ? 'error'
        : edit.value === null
          ? 'blank'
          : typeof edit.value === 'number'
            ? 'number'
            : typeof edit.value === 'boolean'
              ? 'boolean'
              : 'text';
    if (edit.valueType === 'error') cell.error = edit.value;
  }
  if (edit.formula !== undefined) {
    cell.formula = edit.formula.replace(/^=/, '');
    cell.value = null;
    cell.calculation = undefined;
    delete cell.cachedResult;
    cell.cachedValue = null;
    cell.rawValue = null;
    cell.dataType = 'blank';
  }
  for (const key of [
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
  ])
    if (edit[key] !== undefined) cell[key] = edit[key];
  if (edit.styleId !== undefined) cell.styleId = edit.styleId;
  if (edit.numberFormat !== undefined)
    cell.format = {
      general: 0,
      number: 2,
      percent: 10,
      date: 14,
      text: 49,
      currency: '$#,##0.00',
      accounting: 44,
      thousands: 3,
      decimal0: 1,
      decimal1: '0.0',
      decimal2: 2,
    }[edit.numberFormat];
  sheet.cells[edit.cell] = cell;
  const p = cellPosition(edit.cell);
  sheet.rows = Math.max(sheet.rows ?? 1, p.row + 1);
  sheet.columns = Math.max(sheet.columns ?? 1, p.col + 1);
}
function preciseKey(cell) {
  if (!cell?.formula && typeof cell?.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue))
    throw refusal('workbook-unsupported', 'A numeric key cannot be compared without changing its stored precision');
}
const sortCollator = new Intl.Collator('en', { sensitivity: 'accent' });
function sortRows(view, sheet, edit, engine) {
  rejectSpecial(view, sheet);
  const { first, last } = rangeBounds(edit.range),
    sortKeys = normalizeSortKeys(edit).map((key) => ({ ...key, col: cellPosition(`${key.column}1`).col }));
  if (
    (view.tables ?? []).some((table) => table.sheet === sheet.name) ||
    (view.namedRanges ?? []).length ||
    sheet.merges?.length
  )
    throw refusal('sort-unsupported', 'Sorting tables, named ranges or merged sheets is not qualified');
  for (const ref of Object.keys(sheet.cells)) {
    const p = cellPosition(ref);
    if (p.row >= first.row && p.row <= last.row && (p.col < first.col || p.col > last.col))
      throw refusal('sort-bad-input', 'Sort must include every used column in each selected row');
  }
  for (const owner of view.sheets)
    for (const [ref, cell] of Object.entries(owner.cells))
      if (cell.formula) {
        const p = cellPosition(ref),
          inside = owner.name === sheet.name && p.row >= first.row && p.row <= last.row;
        const info = engine.formulaReferences(cell.formula);
        if (info.unsupported)
          throw refusal(
            'sort-unsupported',
            'Sort cannot safely update named, table, dynamic or external formula references'
          );
        for (const dep of info.references) {
          const r = dep.range,
            same = (dep.sheetName || owner.name) === sheet.name,
            intersects = same && r.startRow <= last.row && r.endRow >= first.row;
          // A formula outside the sorted rows keeps its text and is recalculated, as in Excel.
          if (
            inside &&
            intersects &&
            (r.startRow !== p.row ||
              r.endRow !== p.row ||
              [1, 3].includes(r.startAbsoluteRefType) ||
              [1, 3].includes(r.endAbsoluteRefType))
          )
            throw refusal('sort-bad-input', 'Sorted formulas must use relative references to their own row');
          if (
            inside &&
            !intersects &&
            (![1, 3].includes(r.startAbsoluteRefType) || ![1, 3].includes(r.endAbsoluteRefType))
          )
            throw refusal('sort-unsupported', 'References outside a sorted row must use absolute rows');
        }
      }
  const rows = Array.from({ length: last.row - first.row + 1 }, (_, i) => first.row + i);
  const keys = sortKeys.map((key) => {
    const values = new Map(
      rows.map((row) => {
        const cell = sheet.cells[address(row, key.col)];
        preciseKey(cell);
        if (cell?.formula && (cell.calculation?.status !== 'calculated' || cell.calculation.pending))
          throw refusal('sort-unsupported', 'Sort key has an unresolved calculation');
        const value = cell?.formula ? cell.calculation.value : (cell?.value ?? null),
          rank =
            cell?.dataType === 'error'
              ? 3
              : value === null || value === ''
                ? 4
                : typeof value === 'number'
                  ? 0
                  : typeof value === 'string'
                    ? 1
                    : 2;
        return [row, { value, rank }];
      })
    );
    return { ...key, values };
  });
  rows.sort((a, b) => {
    for (const key of keys) {
      const { value: x, rank: ar } = key.values.get(a),
        { value: y, rank: br } = key.values.get(b);
      if (ar === 4 || br === 4) {
        if (ar !== br) return ar === 4 ? 1 : -1;
        continue;
      }
      const compare =
        ar - br || (ar === 0 ? x - y : ar === 1 ? sortCollator.compare(x, y) : ar === 2 ? Number(x) - Number(y) : 0);
      if (compare) return key.direction === 'desc' ? -compare : compare;
    }
    return a - b;
  });
  sheet.rowOrder = rows.map((from, index) => ({ from, to: first.row + index }));
  const destinations = new Map(sheet.rowOrder.map((item) => [item.from, item.to]));
  sheet.rowHeights = (sheet.rowHeights ?? [])
    .map((item) => ({ ...item, row: destinations.get(item.row) ?? item.row }))
    .sort((a, b) => a.row - b.row);
  sheet.hiddenRows = (sheet.hiddenRows ?? []).map((row) => destinations.get(row) ?? row).sort((a, b) => a - b);
  sheet.visibleRows = (sheet.visibleRows ?? []).map((row) => destinations.get(row) ?? row).sort((a, b) => a - b);
  const cells = { ...sheet.cells },
    requests = [],
    targets = [];
  for (let offset = 0; offset < rows.length; offset++)
    for (let col = first.col; col <= last.col; col++) {
      const dest = address(first.row + offset, col),
        source = sheet.cells[address(rows[offset], col)];
      delete cells[dest];
      if (!source) continue;
      cells[dest] = { ...source, calculation: undefined };
      if (source.formula) {
        requests.push({ formula: source.formula, columnOffset: 0, rowOffset: first.row + offset - rows[offset] });
        targets.push(dest);
      }
    }
  const formulas = engine.shiftFormulas(requests);
  targets.forEach((ref, i) => {
    cells[ref].formula = formulas[i];
  });
  sheet.cells = cells;
}
/** Portable, atomic logical projection. No file bytes, durable revisions or engine
 * objects escape this boundary. The preserving writer additionally checks OOXML. */
export async function projectWorkbookCommands(input, commands, engine, { filterStates = {} } = {}) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  if (commands.some((command) => command?.type === 'clearHyperlinks'))
    throw refusal(
      'hyperlink-unsupported',
      'Clear Hyperlinks requires the native XLSX owner; portable projection is unsupported.'
    );
  let view = structuredClone(input);
  const edits = [];
  if (commands.some((command) => command?.type === 'fillAcrossSheets'))
    throw refusal(
      'workbook-unsupported',
      'Across Worksheets requires the native XLSX owner; portable projection is unsupported.'
    );
  for (const command of commands) {
    if (command.type === 'fillSeries') {
      const sheet = view.sheets.find((s) => s.name === command.sheet);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      const b = rangeBounds(command.range, true),
        vertical = b.columns === 1 && b.rows >= 3,
        horizontal = b.rows === 1 && b.columns >= 3;
      if (!vertical && !horizontal)
        throw refusal('workbook-bad-input', 'Fill Series needs one row or column with two seeds and a destination');
      const at = (i) => address(b.first.row + (vertical ? i : 0), b.first.col + (horizontal ? i : 0));
      const first = sheet.cells[at(0)],
        second = sheet.cells[at(1)];
      for (const c of [first, second])
        if (
          !c ||
          c.formula ||
          c.formulaType ||
          typeof c.value !== 'number' ||
          !hasExactNumericValue(c.value, c.rawValue)
        )
          throw refusal('workbook-bad-input', 'Fill Series needs exact numeric seeds');
      const writes = [];
      for (let i = 2; i < (vertical ? b.rows : b.columns); i++) {
        const cell = at(i),
          prior = sheet.cells[cell];
        if (prior?.formula || prior?.formulaType || prior?.value != null)
          throw refusal('workbook-conflict', `Fill Series would overwrite ${cell}`);
        const value = Number((first.value + (second.value - first.value) * i).toPrecision(15));
        if (!Number.isFinite(value))
          throw refusal('workbook-limit', 'Fill Series produced a number outside workbook limits');
        writes.push({ sheet: sheet.name, cell, value });
      }
      const normalized = resolveWorkbookEdits(view, writes);
      normalized.forEach((edit) => setCell(sheet, edit));
      edits.push(...normalized);
      continue;
    }
    if (command.type === 'pasteFormulas') {
      const sheet = view.sheets.find((item) => item.name === command.sheet);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      rejectSpecial(view, sheet);
      const source = rangeBounds(command.source),
        target = rangeBounds(command.range, true);
      if (source.rows !== target.rows || source.columns !== target.columns || target.cells.length > 1000)
        throw refusal(
          'paste-limit',
          'Paste Formulas requires matching source and target rectangles of at most 1,000 cells'
        );
      const requests = [],
        targets = [];
      for (const ref of target.cells) {
        const position = cellPosition(ref),
          from = address(
            source.first.row + position.row - target.first.row,
            source.first.col + position.col - target.first.col
          );
        const formula = sheet.cells[from]?.formula;
        if (!formula) continue;
        requests.push({
          formula,
          rowOffset: position.row - cellPosition(from).row,
          columnOffset: position.col - cellPosition(from).col,
        });
        targets.push(ref);
      }
      if (!requests.length) throw refusal('formula-bad-input', 'Copied range contains no formulas');
      const shifted = engine.shiftFormulas(requests);
      const writes = targets.map((cell, index) => ({ sheet: sheet.name, cell, formula: shifted[index] }));
      const normalized = resolveWorkbookEdits(view, writes);
      normalized.forEach((edit) => setCell(sheet, edit));
      edits.push(...normalized);
      continue;
    }
    if (command.operation === 'createTable') {
      resolveWorkbookEdits(view, [command]);
      const sheet = view.sheets.find((item) => item.name === command.sheet),
        bounds = rangeBounds(command.range, true);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      if ((view.tables ?? []).some((table) => table.name.toLowerCase() === command.tableName.toLowerCase()))
        throw refusal('name-duplicate', 'Table name already exists');
      const intersects = (range) => {
        const other = rangeBounds(range);
        return (
          bounds.first.row <= other.last.row &&
          other.first.row <= bounds.last.row &&
          bounds.first.col <= other.last.col &&
          other.first.col <= bounds.last.col
        );
      };
      if (
        (view.tables ?? []).some((table) => table.sheet === sheet.name && intersects(table.ref)) ||
        (sheet.merges ?? []).some(intersects)
      )
        throw refusal('merge-conflict', 'Table overlaps a table or merged range');
      if (bounds.cells.some((ref) => sheet.cells[ref]?.formula || sheet.cells[ref]?.formulaType))
        throw refusal('table-unsupported', 'Table creation over formulas is not qualified');
      const columns = [],
        seen = new Set();
      for (let col = bounds.first.col; col <= bounds.last.col; col++) {
        const value = sheet.cells[address(bounds.first.row, col)]?.value;
        if (typeof value !== 'string' || !value.trim() || seen.has(value.toLowerCase()))
          throw refusal('table-bad-input', 'Table headers must be unique nonempty text');
        seen.add(value.toLowerCase());
        columns.push({ name: value });
      }
      view.tables = [
        ...(view.tables ?? []),
        {
          name: command.tableName,
          sheet: sheet.name,
          ref: command.range,
          columns,
          headerRowCount: 1,
          totalsRowCount: 0,
        },
      ];
      sheet.tableRanges = [...(sheet.tableRanges ?? []), { name: command.tableName, ref: command.range }];
      edits.push(command);
      continue;
    }
    if (command.operation === 'addConditionalRule') {
      resolveWorkbookEdits(view, [command]);
      const sheet = view.sheets.find((item) => item.name === command.sheet);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      sheet.conditionalRules = [
        ...(sheet.conditionalRules ?? []),
        {
          range: command.range,
          rule: command.rule,
          threshold: command.threshold,
          fillColor: command.fillColor.toUpperCase(),
        },
      ];
      edits.push(command);
      continue;
    }
    if (['mergeCells', 'unmergeCells'].includes(command.operation)) {
      resolveWorkbookEdits(view, [command]);
      const sheet = view.sheets.find((item) => item.name === command.sheet),
        target = rangeBounds(command.range, true);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      const intersects = (range) => {
        const other = rangeBounds(range);
        return (
          target.first.row <= other.last.row &&
          other.first.row <= target.last.row &&
          target.first.col <= other.last.col &&
          other.first.col <= target.last.col
        );
      };
      if (command.operation === 'mergeCells') {
        if (
          (view.tables ?? []).some((table) => table.sheet === sheet.name) ||
          (view.namedRanges ?? []).length ||
          Object.values(sheet.cells).some((cell) => cell.formula || cell.formulaType)
        )
          throw refusal('name-unsupported', 'Merging sheets with formulas, names or tables is not qualified');
        if ((sheet.merges ?? []).some(intersects))
          throw refusal('merge-conflict', 'Merge overlaps an existing merged range');
        for (const ref of target.cells.slice(1))
          if (sheet.cells[ref]?.value != null)
            throw refusal('merge-conflict', `Merge would hide a nonempty cell: ${ref}`);
        sheet.merges = [...(sheet.merges ?? []), command.range];
      } else {
        if (!(sheet.merges ?? []).includes(command.range))
          throw refusal('merge-bad-input', 'Select the exact merged range to unmerge');
        sheet.merges = sheet.merges.filter((range) => range !== command.range);
      }
      edits.push(command);
      continue;
    }
    if (['addSheet', 'renameSheet', 'duplicateSheet'].includes(command.operation)) {
      resolveWorkbookEdits(view, [command]);
      if (command.operation === 'addSheet')
        view.sheets.push({
          name: command.sheet,
          state: 'visible',
          path: '',
          cells: {},
          rows: 20,
          columns: 8,
          widths: [],
          merges: [],
          hiddenRows: [],
          visibleRows: [],
          rowHeights: [],
          sheetFormat: { rowsHiddenByDefault: false },
          normalFont: {
            ...(view.sheets[0]?.normalFont ?? { fontFamily: 'Calibri', fontSize: 11, bold: false, italic: false }),
          },
        });
      else {
        const source = view.sheets.find((item) => item.name === command.sheet);
        if (command.operation === 'duplicateSheet') {
          if (
            Object.values(source.cells).some((cell) => cell.formula || cell.formulaType) ||
            source.merges?.length ||
            (view.tables ?? []).some((table) => table.sheet === source.name)
          )
            throw refusal('merge-unsupported', 'Duplicating a sheet with formulas, merges or tables is not qualified');
          view.sheets.push({ ...structuredClone(source), name: command.newName, path: '' });
        } else {
          if (
            view.sheets.some((sheet) => Object.values(sheet.cells).some((cell) => cell.formula || cell.formulaType)) ||
            (view.namedRanges ?? []).length ||
            (view.tables ?? []).length
          )
            throw refusal(
              'name-unsupported',
              'Renaming a sheet with workbook formulas, names or tables is not qualified'
            );
          source.name = command.newName;
        }
      }
      edits.push(command);
      continue;
    }
    if (command.type === 'pasteValues') {
      const sourceSheet = view.sheets.find((s) => s.name === (command.sourceSheet ?? command.sheet)),
        sheet = view.sheets.find((s) => s.name === command.sheet);
      if (!sheet || !sourceSheet) throw refusal('workbook-not-found', 'Sheet not found');
      const source = rangeBounds(command.source),
        target = rangeBounds(command.range, true),
        writes = [],
        context = createValuePasteSourceContext(sourceSheet),
        snapshots = new Map();
      const overlaps = (range) => {
        const area = rangeBounds(range);
        return (
          area.first.row <= target.last.row &&
          area.last.row >= target.first.row &&
          area.first.col <= target.last.col &&
          area.last.col >= target.first.col
        );
      };
      if ((sheet.merges ?? []).some(overlaps))
        throw refusal('merge-conflict', 'Paste Values destination intersects merged cells');
      if ((view.tables ?? []).some((table) => table.sheet === sheet.name && overlaps(table.ref)))
        throw refusal('table-conflict', 'Paste Values destination intersects a table');
      for (const ref of target.cells) {
        const p = cellPosition(ref),
          row = source.first.row + ((p.row - target.first.row) % source.rows),
          col = source.first.col + ((p.col - target.first.col) % source.columns),
          cell = sourceSheet.cells[address(row, col)] ?? { value: null };
        const sourceRef = address(row, col);
        if (!snapshots.has(sourceRef))
          snapshots.set(sourceRef, normalizeValuePasteCell(cell, { ...context, cell: sourceRef }));
        writes.push({ sheet: sheet.name, cell: ref, ...snapshots.get(sourceRef) });
      }
      const normalized = resolveWorkbookEdits(view, writes);
      normalized.forEach((edit) => setCell(sheet, edit));
      edits.push(...normalized);
      continue;
    }
    if (command.type === 'fillRange') {
      const sheet = view.sheets.find((item) => item.name === command.sheet);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      rejectSpecial(view, sheet);
      const source = rangeBounds(command.source),
        target = rangeBounds(command.range, true),
        requests = [],
        writes = [],
        styles = [];
      for (const ref of target.cells) {
        const p = cellPosition(ref),
          row = source.first.row + ((p.row - target.first.row) % (source.last.row - source.first.row + 1)),
          col = source.first.col + ((p.col - target.first.col) % (source.last.col - source.first.col + 1)),
          cell = sheet.cells[address(row, col)] ?? { value: null };
        if (!cell.formula && typeof cell.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue))
          throw refusal('workbook-unsupported', 'Fill cannot copy a numeric value whose stored precision would change');
        const write = {
          sheet: sheet.name,
          cell: ref,
          ...(cell.formula ? { formula: cell.formula } : literalFillSnapshot(cell)),
          ...(cell.styleId !== undefined
            ? { styleId: cell.styleId }
            : {
                ...(cell.bold !== undefined ? { bold: cell.bold } : {}),
                ...(cell.italic !== undefined ? { italic: cell.italic } : {}),
              }),
        };
        styles.push(Object.fromEntries(['styleId', 'format', 'bold', 'italic'].map((field) => [field, cell[field]])));
        if (cell.formula) {
          write.formulaIndex = requests.length;
          requests.push({ formula: cell.formula, columnOffset: p.col - col, rowOffset: p.row - row });
        }
        writes.push(write);
      }
      const shifted = engine.shiftFormulas(requests);
      for (const write of writes) {
        if (own(write, 'formulaIndex')) {
          write.formula = shifted[write.formulaIndex];
          delete write.formulaIndex;
        }
      }
      const normalized = resolveWorkbookEdits(view, writes);
      normalized.forEach((edit, index) => {
        setCell(sheet, edit);
        Object.assign(sheet.cells[edit.cell], styles[index]);
      });
      edits.push(...normalized);
      continue;
    }
    const normalized = resolveWorkbookEdits(
      { ...view, structuralReason: command.operation === 'insertRow' ? undefined : view.structuralReason },
      [command]
    );
    for (const edit of normalized) {
      let sheet = view.sheets.find((item) => item.name === edit.sheet);
      if (!edit.operation) setCell(sheet, edit);
      else if (edit.operation === 'insertRow') {
        view = await engine.transformWorkbook(view, [edit]);
        const target = view.sheets.find((item) => item.name === edit.sheet),
          count = edit.count ?? 1;
        target.rowHeights = (sheet.rowHeights ?? []).map((item) => ({
          ...item,
          row: item.row >= edit.index ? item.row + count : item.row,
        }));
        target.hiddenRows = (sheet.hiddenRows ?? []).map((row) => (row >= edit.index ? row + count : row));
        target.visibleRows = (sheet.visibleRows ?? []).map((row) => (row >= edit.index ? row + count : row));
      } else if (edit.operation === 'sortRows') {
        const calculated = await engine.calculateWorkbook(view);
        view = calculated;
        sortRows(
          view,
          view.sheets.find((item) => item.name === edit.sheet),
          edit,
          engine
        );
      } else if (['filterRows', 'reapplyFilter', 'clearFilter', 'removeFilter'].includes(edit.operation)) {
        if (!['clearFilter', 'removeFilter'].includes(edit.operation) && view.calculationPolicy?.mode !== 'manual')
          view = await engine.calculateWorkbook(view);
        sheet = view.sheets.find((item) => item.name === edit.sheet);
        const prior = filterStates[edit.sheet];
        if (sheet.filter && !prior && sheet.filter.reason !== 'No active owned filter criteria')
          throw refusal('workbook-refused', sheet.filter.reason ?? 'Owned filter baseline is unavailable');
        const originalHidden = new Set(sheet.hiddenRows ?? []);
        const b = rangeBounds(edit.range ?? prior?.model.range),
          baseline = {
            format: sheet.sheetFormat?.rowsHiddenByDefault ? { zeroHeight: '1' } : {},
            rows: Array.from(new Set([...(sheet.hiddenRows ?? []), ...(sheet.visibleRows ?? [])]))
              .filter((r) => r > b.first.row && r <= b.last.row)
              .sort((a, b) => a - b)
              .map((r) => [r, originalHidden.has(r) ? '1' : '0']),
          },
          record = prepareFilterState(edit, prior, baseline),
          { first, last } = rangeBounds(record.model.range),
          col = cellPosition(`${record.model.column}1`).col,
          values = filterDefinition(record).values,
          baselineRows = new Map(record.model.baseline.rows);
        const hidden = (sheet.hiddenRows ?? []).filter((row) => row <= first.row || row > last.row),
          visible = (sheet.visibleRows ?? []).filter((row) => row <= first.row || row > last.row);
        for (let row = first.row + 1; row <= last.row; row++) {
          let fails = false;
          if (values.length) {
            const cell = sheet.cells[address(row, col)];
            preciseKey(cell);
            if (cell?.dataType === 'error' && !cell.formula)
              throw refusal('filter-unsupported', 'Filtering typed errors requires separate qualification');
            if (cell?.formula && (cell.calculation?.status !== 'calculated' || cell.calculation.pending))
              throw refusal('filter-unsupported', 'Filter key has an unresolved calculation');
            const value = cell?.formula ? cell.calculation.value : (cell?.value ?? null);
            fails = !values.some((expected) => Object.is(value, expected));
          }
          if (baselineHidden(record, row, baselineRows) || fails) hidden.push(row);
          else if (baselineRows.get(row) != null) visible.push(row);
        }
        sheet.hiddenRows = hidden;
        sheet.visibleRows = visible;
        if (edit.operation === 'removeFilter') delete sheet.filter;
        else sheet.filter = filterDefinition(record);
      } else if (edit.operation === 'resizeColumn') {
        if (
          !Number.isInteger(edit.index) ||
          edit.index < 0 ||
          edit.index > 16383 ||
          !Number.isFinite(edit.width) ||
          edit.width < 8 ||
          edit.width > 100
        )
          throw refusal('internal-invalid-field', 'Invalid column width');
        let hidden = false;
        const widths = [];
        for (const width of sheet.widths ?? []) {
          if (edit.index + 1 < width.min || edit.index + 1 > width.max) widths.push(width);
          else {
            hidden = width.hidden ?? false;
            if (width.min < edit.index + 1) widths.push({ ...width, max: edit.index });
            if (width.max > edit.index + 1) widths.push({ ...width, min: edit.index + 2 });
          }
        }
        widths.push({ min: edit.index + 1, max: edit.index + 1, width: edit.width, hidden });
        sheet.widths = widths.sort((a, b) => a.min - b.min);
      } else if (edit.operation === 'resizeRow') {
        if (
          !Number.isInteger(edit.index) ||
          edit.index < 0 ||
          edit.index > 1048575 ||
          !Number.isFinite(edit.height) ||
          edit.height < 8 ||
          edit.height > 200
        )
          throw refusal('internal-invalid-field', 'Invalid row height');
        sheet.rowHeights = [
          ...(sheet.rowHeights ?? []).filter((item) => item.row !== edit.index),
          { row: edit.index, height: edit.height },
        ].sort((a, b) => a.row - b.row);
      } else if (edit.operation === 'unhideAll') {
        if (view.readOnly || sheet.filter || (view.tables ?? []).some((t) => t.sheet === sheet.name))
          throw refusal('filter-protected', 'Unhide requires an unprotected worksheet without linked filters/tables');
        const rh = sheet.sheetFormat?.defaultRowHeight > 0 ? sheet.sheetFormat.defaultRowHeight : 15,
          cw = sheet.sheetFormat?.defaultColumnWidth > 0 ? sheet.sheetFormat.defaultColumnWidth : 8.43;
        sheet.sheetFormat = {
          ...sheet.sheetFormat,
          defaultRowHeight: rh,
          defaultColumnWidth: cw,
          rowsHiddenByDefault: false,
        };
        sheet.hiddenRows = [];
        sheet.visibleRows = [];
        sheet.rowHeights = (sheet.rowHeights ?? []).map((r) => ({ ...r, height: r.height === 0 ? rh : r.height }));
        sheet.widths = (sheet.widths ?? []).map((c) => ({ ...c, hidden: false, width: c.width === 0 ? cw : c.width }));
      } else if (edit.operation === 'freeze') {
        if (
          !Number.isInteger(edit.rows) ||
          edit.rows < 0 ||
          edit.rows > 100 ||
          !Number.isInteger(edit.columns) ||
          edit.columns < 0 ||
          edit.columns > 30
        )
          throw refusal('layout-malformed', 'Invalid freeze boundary');
        sheet.freeze = { rows: edit.rows, columns: edit.columns };
      } else {
        if (view.structuralReason) throw refusal('workbook-refused', view.structuralReason);
        rejectSpecial(view, sheet);
        if (
          view.sheets.some((item) => Object.values(item.cells).some((cell) => cell.formula)) ||
          (view.namedRanges ?? []).length ||
          (view.tables ?? []).length
        )
          throw refusal('formula-unsupported', 'This structural operation is not qualified for formula workbooks');
        const row = edit.operation.endsWith('Row'),
          remove = edit.operation.startsWith('delete'),
          cells = {};
        for (const [ref, cell] of Object.entries(sheet.cells)) {
          const p = cellPosition(ref),
            n = row ? p.row : p.col;
          if (remove && n === edit.index) continue;
          if (n >= edit.index) {
            if (row) p.row += remove ? -1 : 1;
            else p.col += remove ? -1 : 1;
          }
          cells[address(p.row, p.col)] = cell;
        }
        sheet.cells = cells;
      }
      edits.push(edit);
    }
  }
  // Every formula result belongs to the pre-edit state until the caller calculates.
  for (const sheet of view.sheets)
    for (const cell of Object.values(sheet.cells)) if (cell.formula || cell.formulaType) delete cell.calculation;
  return { view, edits };
}

/** Validate ordinary cell batches against only their addressed coordinates. This
 * is the same guard/normalization path as full projection, with no sheet scan.
 * A null result delegates structural, keyed, fill and recalculation commands to
 * projectWorkbookCommands. No input cell or accepted history frame is mutated. */
export function projectCellCommands(input, commands) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  if (commands.some((command) => command?.type === 'clearHyperlinks'))
    throw refusal(
      'hyperlink-unsupported',
      'Clear Hyperlinks requires the native XLSX owner; portable projection is unsupported.'
    );
  const cellsBySheet = new Map(input.sheets.map((sheet) => [sheet.name, {}]));
  for (const command of commands) {
    if (command.operation || !['setCells', 'formatRange', undefined].includes(command.type)) return null;
    const refs =
      command.type === 'setCells'
        ? Array.isArray(command.cells)
          ? command.cells.map((cell) => cell.cell)
          : []
        : command.type === 'formatRange'
          ? rangeCells(command.range)
          : [command.cell];
    for (const ref of refs) {
      const p = cellPosition(ref),
        canonical = address(p.row, p.col),
        sheet = input.sheets.find((item) => item.name === command.sheet);
      if (!sheet) throw refusal('workbook-not-found', 'Sheet not found');
      cellsBySheet.get(sheet.name)[canonical] = sheet.cells[canonical];
    }
  }
  const bounded = {
    ...input,
    sheets: input.sheets.map((sheet) => ({ ...sheet, cells: cellsBySheet.get(sheet.name) })),
  };
  const edits = resolveWorkbookEdits(bounded, commands),
    changes = new Map();
  for (const edit of edits) {
    const sheet = bounded.sheets.find((item) => item.name === edit.sheet);
    setCell(sheet, edit);
    changes.set(`${edit.sheet}!${edit.cell}`, { sheet: edit.sheet, cell: edit.cell, value: sheet.cells[edit.cell] });
  }
  return { edits, changes: [...changes.values()] };
}
