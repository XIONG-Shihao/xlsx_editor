/** Compact table definition edits; worksheet values remain in their native owner. */
import { nodeStart } from './xml-position.mjs';
import { createHash } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import { openZipParts } from './zip-parts.mjs';
import { strFromU8 } from 'fflate';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import { packageReadOnlyReasons } from './package-capabilities.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const children = (node) => Array.from(node.childNodes).filter((node) => node.nodeType === 1);
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const escape = (value) =>
  value.replace(/[&<>"\r]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;' })[c]);
function parse(bytes, root, ns = NS) {
  const source = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('table-malformed', 'Table XML entities are protected');
  const fail = (message) => {
      throw refusal('package-malformed', message);
    },
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (value) => value,
      errorHandler: { warning: fail, error: fail, fatalError: fail },
    }).parseFromString(source, 'application/xml');
  if (doc.documentElement?.namespaceURI !== ns || doc.documentElement.localName !== root)
    throw refusal('table-malformed', 'Invalid table metadata namespace/root');
  return { source, doc };
}
function attrs(node, allowed) {
  if (
    Array.from(node.attributes).some(
      (attr) =>
        attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
        (attr.namespaceURI || !allowed.includes(attr.localName))
    )
  )
    throw refusal('table-unsupported', 'Extended table attributes require qualification');
}
function leaf(node, allowed) {
  attrs(node, allowed);
  if (children(node).length || node.textContent.trim())
    throw refusal('table-protected', 'Extended table metadata is protected');
}
function name(value) {
  if (
    typeof value !== 'string' ||
    value.length > 255 ||
    !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(value) ||
    /^_xlnm\./i.test(value) ||
    /^(?:R|C|[A-Za-z]{1,3}\d+|R\d+C\d+)$/i.test(value)
  )
    throw refusal('table-unsupported', 'Use a table name without cell addresses or reserved names');
  return value;
}
function rawStart(source, node) {
  const start = nodeStart(source, node);
  if (!source.startsWith('<' + node.tagName, start))
    throw refusal('table-refused', 'Table XML lacks an exact raw boundary');
  return start;
}
function opening(source, node, changes) {
  const start = rawStart(source, node),
    end = xmlElementEnd(source, start),
    raw = source.slice(start, end),
    close = raw.indexOf('>');
  if (close < 0) throw refusal('internal-invalid-field', 'Invalid table opening');
  let head = raw.slice(0, close + 1);
  for (const [key, value] of Object.entries(changes)) {
    const regex = new RegExp('(\\s' + key + '\\s*=\\s*)(["\\\'])(.*?)\\2', 's');
    if (!regex.test(head)) throw refusal('internal-missing-field', 'Missing table ' + key);
    head = head.replace(regex, (_all, prefix, quote) => prefix + quote + escape(value) + quote);
  }
  return { start, end: start + close + 1, text: head };
}
function qualified(parts, view) {
  if (packageReadOnlyReasons(parts).length) throw refusal('table-protected', 'Read-only package cannot change a table');
  if (view.namedRanges?.length || view.sheets.some((sheet) => sheet.extended || sheet.specialFormulaCount))
    throw refusal('table-unsupported', 'Special formulas or name dependencies require table-edit qualification');
  for (const [part, data] of Object.entries(parts)) {
    if (part.endsWith('.rels')) {
      const root = parse(data, 'Relationships', REL).doc.documentElement;
      for (const entry of children(root)) {
        if (
          entry.namespaceURI !== REL ||
          entry.localName !== 'Relationship' ||
          entry.getAttribute('TargetMode') === 'External'
        )
          throw refusal('table-protected', 'External/opaque table dependency is protected');
        const type = entry.getAttribute('Type');
        if (
          !/(?:\/(?:officeDocument|worksheet|table|styles|sharedStrings|theme|sheetMetadata|extended-properties)|\/metadata\/(?:core-properties|thumbnail))$/.test(
            type
          )
        )
          throw refusal('table-unsupported', 'Object, pivot or external table dependencies require qualification');
      }
    }
    if (
      view.sheets.some((sheet) => sheet.path === part) &&
      /<(?:[A-Za-z_][\w.-]*:)?(?:formula1|formula2|formula|conditionalFormatting|dataValidations|drawing|legacyDrawing|oleObjects|controls|extLst)\b/.test(
        strFromU8(data)
      )
    )
      throw refusal('table-unsupported', 'Rule/object table dependencies require qualification');
  }
}
function context(bytes, view, command) {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > 16 * 1024 ** 2)
    throw refusal('table-limit', 'Table metadata exceeds its budget');
  const parts = openZipParts(bytes, 16 * 1024 ** 2);
  qualified(parts, view);
  const found = (view.tables ?? []).filter(
    (table) => table.sheet === command.sheet && table.name.toLowerCase() === String(command.tableName).toLowerCase()
  );
  if (found.length !== 1) throw refusal('internal-missing-field', 'Table identity is missing or ambiguous');
  const table = found[0],
    sheet = view.sheets.find((sheet) => sheet.name === table.sheet);
  if (!sheet || !parts[table.path]) throw refusal('internal-missing-field', 'Table source is missing');
  const { source, doc } = parse(parts[table.path], 'table'),
    root = doc.documentElement;
  attrs(root, ['id', 'name', 'displayName', 'ref', 'headerRowCount', 'totalsRowCount', 'totalsRowShown']);
  if (
    root.getAttribute('name') !== table.name ||
    root.getAttribute('displayName') !== table.name ||
    root.getAttribute('ref') !== table.ref ||
    Number(root.getAttribute('headerRowCount') || 1) !== 1 ||
    Number(root.getAttribute('totalsRowCount') || 0) !== 0
  )
    throw refusal('table-unsupported', 'Only ordinary one-header tables without totals are qualified');
  const direct = children(root);
  if (
    direct.some(
      (node) => node.namespaceURI !== NS || !['autoFilter', 'tableColumns', 'tableStyleInfo'].includes(node.localName)
    ) ||
    new Set(direct.map((node) => node.localName)).size !== direct.length
  )
    throw refusal('table-unsupported', 'Extended table children require qualification');
  const columns = direct.find((node) => node.localName === 'tableColumns');
  if (!columns) throw refusal('internal-missing-field', 'Missing table columns');
  attrs(columns, ['count']);
  const entries = children(columns),
    b = rangeBounds(table.ref);
  if (
    entries.length !== b.columns ||
    entries.length !== Number(columns.getAttribute('count')) ||
    b.size > 10000 ||
    b.rows < 2 ||
    b.columns > 1000 ||
    entries.some((node) => node.namespaceURI !== NS || node.localName !== 'tableColumn')
  )
    throw refusal('table-unsupported', 'Inconsistent table column count');
  const ids = new Set(),
    names = new Set();
  for (const entry of entries) {
    leaf(entry, ['id', 'name']);
    const id = Number(entry.getAttribute('id')),
      label = entry.getAttribute('name');
    if (!Number.isSafeInteger(id) || id < 1 || ids.has(id) || !label || names.has(label.toLowerCase()))
      throw refusal('table-duplicate', 'Duplicate/invalid table column identity');
    ids.add(id);
    names.add(label.toLowerCase());
  }
  const filter = direct.find((node) => node.localName === 'autoFilter');
  if (filter) {
    leaf(filter, ['ref']);
    if (filter.getAttribute('ref') !== table.ref)
      throw refusal('table-unsupported', 'Filtered table resize requires qualification');
  }
  const style = direct.find((node) => node.localName === 'tableStyleInfo');
  if (style) leaf(style, ['name', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes']);
  return { parts, table, sheet, root, filter, source, b, entries };
}
export function isIndexedTableCommand(command) {
  return ['renameTable', 'resizeTable', 'renameTableColumn', 'insertTableColumn', 'deleteTableColumn'].includes(
    command?.operation
  );
}
export async function indexedTablePlan(bytes, view, command, { readRange } = {}) {
  if (
    !isIndexedTableCommand(command) ||
    Object.keys(command).some(
      (key) =>
        !['operation', 'sheet', 'tableName', 'expectedTableRef', 'newName', 'range', 'columnName', 'index'].includes(
          key
        )
    )
  )
    throw refusal('internal-invalid-field', 'Invalid table command');
  name(command.tableName);
  const ctx = context(bytes, view, command);
  if (command.expectedTableRef !== ctx.table.ref) throw refusal('table-stale', 'Conflict: table range changed');
  let patches = [],
    extra = {};
  if (['insertTableColumn', 'deleteTableColumn'].includes(command.operation)) {
    const remove = command.operation === 'deleteTableColumn';
    if (
      command.range !== undefined ||
      (remove && (command.index !== undefined || command.newName !== undefined)) ||
      (!remove && command.columnName !== undefined)
    )
      throw refusal('table-malformed', 'Unexpected table column structure fields');
    const label = (value) => {
      if (
        typeof value !== 'string' ||
        !value.trim() ||
        value !== value.trim() ||
        Buffer.byteLength(value, 'utf8') > 255 ||
        /[\[\]':,#@\x00-\x1F]/.test(value)
      )
        throw refusal(
          'table-bad-input',
          'Use an ordinary column label without selector punctuation or control characters'
        );
      return value;
    };
    const text = label(remove ? command.columnName : command.newName),
      index = remove ? ctx.entries.findIndex((node) => node.getAttribute('name') === text) : command.index;
    if (
      !Number.isSafeInteger(index) ||
      index < 0 ||
      index > ctx.entries.length ||
      (remove && index >= ctx.entries.length)
    )
      throw refusal('table-limit', 'Table column index/label changed or is out of bounds');
    if (remove && ctx.entries.length === 1) throw refusal('table-bad-input', 'Keep at least one table column');
    if (!remove && ctx.entries.some((node) => node.getAttribute('name').toLowerCase() === text.toLowerCase()))
      throw refusal('table-conflict', 'Table column label collision');
    const last = ctx.b.last.col + (remove ? -1 : 1),
      count = ctx.b.columns + (remove ? -1 : 1),
      range = columnName(ctx.b.first.col) + (ctx.b.first.row + 1) + ':' + columnName(last) + (ctx.b.last.row + 1),
      next = rangeBounds(range);
    if (last >= 16384 || count > 1000 || next.size > 10000)
      throw refusal('table-limit', 'Table column change exceeds the supported bounds');
    if (
      (view.tables ?? []).some(
        (table) =>
          table.path !== ctx.table.path && table.sheet === ctx.table.sheet && overlaps(next, rangeBounds(table.ref))
      ) ||
      (ctx.sheet.merges ?? []).some((ref) => overlaps(next, rangeBounds(ref)))
    )
      throw refusal('table-conflict', 'Table column change intersects another table or merge');
    if (!remove) {
      if (typeof readRange !== 'function') throw refusal('table-bad-input', 'Table growth needs native cell admission');
      for (let first = ctx.b.first.row; first <= ctx.b.last.row; first += 1000) {
        const growth =
            columnName(last) + (first + 1) + ':' + columnName(last) + (Math.min(ctx.b.last.row, first + 999) + 1),
          reply = await readRange(ctx.table.sheet, growth);
        if (
          Object.values(reply.cells ?? {}).some(
            (cell) =>
              cell.formula || cell.formulaType || cell.spill || (cell.value !== null && cell.value !== undefined)
          )
        )
          throw refusal('table-unsupported', 'Table growth cannot absorb existing content or spill ownership');
      }
    }
    const columns = children(ctx.root).find((node) => node.localName === 'tableColumns');
    patches = [opening(ctx.source, ctx.root, { ref: range }), opening(ctx.source, columns, { count: String(count) })];
    if (ctx.filter) patches.push(opening(ctx.source, ctx.filter, { ref: range }));
    if (remove) {
      const start = rawStart(ctx.source, ctx.entries[index]);
      patches.push({ start, end: xmlElementEnd(ctx.source, start), text: '' });
    } else {
      const id = Math.max(...ctx.entries.map((node) => Number(node.getAttribute('id')))) + 1;
      if (!Number.isSafeInteger(id)) throw refusal('table-limit', 'Table column id exceeds its bound');
      const start =
        index < ctx.entries.length
          ? rawStart(ctx.source, ctx.entries[index])
          : xmlElementEnd(ctx.source, rawStart(ctx.source, ctx.entries.at(-1)));
      patches.push({ start, end: start, text: `<${ctx.entries[0].tagName} id="${id}" name="${escape(text)}"/>` });
    }
    extra = { columnChange: { index, remove, label: text }, range, sheets: view.sheets.map((sheet) => sheet.path) };
  } else if (command.operation === 'renameTableColumn') {
    if (command.index !== undefined)
      throw refusal('table-bad-input', 'Column rename does not accept an insertion index');
    if (command.range !== undefined) throw refusal('table-bad-input', 'Column rename does not accept a range');
    const label = (value) => {
      if (
        typeof value !== 'string' ||
        !value.trim() ||
        value !== value.trim() ||
        Buffer.byteLength(value, 'utf8') > 255 ||
        /[\[\]':,#@\x00-\x1F]/.test(value)
      )
        throw refusal(
          'table-bad-input',
          'Use an ordinary column label without selector punctuation or control characters'
        );
      return value;
    };
    const prior = label(command.columnName),
      next = label(command.newName),
      index = ctx.entries.findIndex((node) => node.getAttribute('name') === prior);
    if (index < 0) throw refusal('table-stale', 'Conflict: table column changed');
    if (prior === next) throw refusal('table-unsupported', 'Choose a different table column label');
    if (ctx.entries.some((node, i) => i !== index && node.getAttribute('name').toLowerCase() === next.toLowerCase()))
      throw refusal('table-conflict', 'Table column label collision');
    patches = [opening(ctx.source, ctx.entries[index], { name: next })];
    extra = { renameTableColumn: { index, old: prior, new: next }, sheets: view.sheets.map((sheet) => sheet.path) };
  } else if (command.operation === 'renameTable') {
    if (command.index !== undefined)
      throw refusal('table-bad-input', 'Table rename does not accept an insertion index');
    if (command.columnName !== undefined) throw refusal('table-bad-input', 'Table rename does not accept a column');
    if (command.range !== undefined) throw refusal('table-bad-input', 'Rename does not accept a new range');
    const next = name(command.newName);
    if (next === ctx.table.name) throw refusal('table-unsupported', 'Choose a different table name');
    if (
      (view.tables ?? []).some(
        (table) => table.path !== ctx.table.path && table.name.toLowerCase() === next.toLowerCase()
      ) ||
      (view.namedRanges ?? []).some((entry) => entry.name.toLowerCase() === next.toLowerCase())
    )
      throw refusal('table-conflict', 'Table/name collision');
    patches = [opening(ctx.source, ctx.root, { name: next, displayName: next })];
    extra = { renameTable: { old: ctx.table.name, new: next }, sheets: view.sheets.map((sheet) => sheet.path) };
  } else {
    if (command.index !== undefined || command.newName !== undefined || command.columnName !== undefined)
      throw refusal('table-bad-input', 'Resize does not accept a new name or column');
    const b = rangeBounds(command.range);
    if (
      b.first.row !== ctx.b.first.row ||
      b.first.col !== ctx.b.first.col ||
      b.last.col !== ctx.b.last.col ||
      b.rows < 2 ||
      b.size > 10000
    )
      throw refusal(
        'table-refused',
        'Resize currently retains the header origin and complete column set with <=10,000 positions'
      );
    if (b.last.row === ctx.b.last.row) throw refusal('table-unsupported', 'Choose a different table size');
    if (
      (view.tables ?? []).some(
        (table) =>
          table.path !== ctx.table.path && table.sheet === ctx.table.sheet && overlaps(b, rangeBounds(table.ref))
      ) ||
      (ctx.sheet.merges ?? []).some((range) => overlaps(b, rangeBounds(range)))
    )
      throw refusal('table-conflict', 'Table resize intersects another table or merge');
    if (b.last.row > ctx.b.last.row) {
      if (typeof readRange !== 'function') throw refusal('table-bad-input', 'Table growth needs native cell admission');
      const cols = b.columns,
        rows = Math.max(1, Math.floor(1000 / cols));
      for (let first = ctx.b.last.row + 1; first <= b.last.row; first += rows) {
        const range =
            columnName(b.first.col) +
            (first + 1) +
            ':' +
            columnName(b.last.col) +
            (Math.min(b.last.row, first + rows - 1) + 1),
          reply = await readRange(ctx.table.sheet, range);
        if (
          Object.values(reply.cells ?? {}).some(
            (cell) =>
              cell.formula || cell.spill || cell.formulaType || (cell.value !== null && cell.value !== undefined)
          )
        )
          throw refusal('table-unsupported', 'Table growth cannot absorb existing content or spill ownership');
      }
    }
    patches = [opening(ctx.source, ctx.root, { ref: command.range })];
    if (ctx.filter) patches.push(opening(ctx.source, ctx.filter, { ref: command.range }));
    extra = { range: command.range };
  }
  let replacement = ctx.source;
  for (const patch of patches.sort((a, b) => b.start - a.start))
    replacement = replacement.slice(0, patch.start) + patch.text + replacement.slice(patch.end);
  return {
    kind: command.operation,
    sheet: ctx.sheet.path,
    tablePart: ctx.table.path,
    tableName: ctx.table.name,
    priorRef: ctx.table.ref,
    priorTableSha256: sha(ctx.parts[ctx.table.path]),
    ...extra,
    replacements: { [ctx.table.path]: replacement },
    removed: [],
    copies: {},
    fragments: {},
  };
}
function overlaps(a, b) {
  return (
    a.first.row <= b.last.row && a.last.row >= b.first.row && a.first.col <= b.last.col && a.last.col >= b.first.col
  );
}
