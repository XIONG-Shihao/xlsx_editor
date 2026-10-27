/** Preserve ordinary neighboring cells while replacing only owned pivot output. */
import { lineStarts } from './xml-position.mjs';
import { DOMParser } from '@xmldom/xmldom';
import { rangeBounds, cellPosition, columnName } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  elements = (n) => Array.from(n.childNodes ?? []).filter((n) => n.nodeType === 1);
function parse(raw) {
  const errors = [],
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (v) => v,
      errorHandler: { warning: (e) => errors.push(e), error: (e) => errors.push(e), fatalError: (e) => errors.push(e) },
    }).parseFromString(raw, 'application/xml');
  if (
    errors.length ||
    doc.documentElement.namespaceURI !== NS ||
    doc.documentElement.localName !== 'worksheet' ||
    Buffer.byteLength(raw) > 16 * 1024 ** 2 ||
    /<!DOCTYPE|<!ENTITY/i.test(raw)
  )
    throw refusal('pivot-malformed', 'Invalid pivot target worksheet');
  return doc.documentElement;
}
function span(raw, node) {
  const starts = lineStarts(raw);
  const first = starts[node.lineNumber - 1] + node.columnNumber - 1;
  if (!raw.startsWith('<' + node.tagName, first))
    throw refusal('internal-missing-field', 'Missing pivot target raw boundary');
  return [first, xmlElementEnd(raw, first)];
}
function direct(root, tag) {
  const found = elements(root).filter((n) => n.namespaceURI === NS && n.localName === tag);
  if (found.length !== 1) throw refusal('pivot-ambiguous', 'Ambiguous pivot target ' + tag);
  return found[0];
}
function attrs(node, keys) {
  for (const a of Array.from(node.attributes ?? []))
    if (!a.name.startsWith('xmlns') && !keys.includes(a.name))
      throw refusal('pivot-unsupported', 'Extended neighboring pivot target attributes');
}
function position(ref) {
  const p = cellPosition(ref);
  if (columnName(p.col) + (p.row + 1) !== ref) throw refusal('pivot-malformed', 'Noncanonical pivot target coordinate');
  return p;
}
const contains = (b, p) => p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
function ordinary(cell) {
  attrs(cell, ['r', 's', 't']);
  if (cell.hasAttribute('s') && cell.getAttribute('s') !== '0')
    throw refusal('pivot-unsupported', 'Neighboring pivot styles require qualification');
  const kids = elements(cell);
  if (
    kids.some((n) => n.namespaceURI !== NS || !['f', 'v', 'is'].includes(n.localName)) ||
    kids.length > 2 ||
    kids.filter((n) => n.localName === 'f').length > 1
  )
    throw refusal('pivot-unsupported', 'Opaque neighboring pivot cell');
  for (const child of kids) {
    if (child.localName === 'is') {
      attrs(child, []);
      const ts = elements(child);
      if (ts.length !== 1 || ts[0].namespaceURI !== NS || ts[0].localName !== 't' || elements(ts[0]).length)
        throw refusal('pivot-protected', 'Rich neighboring pivot text is protected');
      attrs(ts[0], ['xml:space']);
    } else {
      attrs(child, []);
      if (elements(child).length) throw refusal('pivot-unsupported', 'Extended neighboring pivot cell');
    }
  }
}
function cells(root) {
  const data = direct(root, 'sheetData'),
    seen = new Set(),
    rows = new Map(),
    entries = [];
  attrs(data, []);
  for (const row of elements(data)) {
    if (row.namespaceURI !== NS || row.localName !== 'row')
      throw refusal('pivot-malformed', 'Invalid pivot target row');
    attrs(row, ['r']);
    const number = Number(row.getAttribute('r'));
    if (!Number.isSafeInteger(number) || number < 1 || number > 1048576 || rows.has(number))
      throw refusal('pivot-ambiguous', 'Ambiguous pivot target row');
    rows.set(number, row);
    let col = -1;
    for (const cell of elements(row)) {
      if (cell.namespaceURI !== NS || cell.localName !== 'c')
        throw refusal('pivot-malformed', 'Invalid pivot target cell');
      const ref = cell.getAttribute('r'),
        p = position(ref);
      if (p.row + 1 !== number || p.col <= col || seen.has(ref))
        throw refusal('pivot-malformed', 'Unordered/duplicate pivot target cell');
      seen.add(ref);
      col = p.col;
      entries.push({ cell, ref, p, row });
      if (entries.length > 10000) throw refusal('pivot-limit', 'Pivot neighboring data exceeds its bound');
    }
  }
  return { data, rows, entries };
}
function skeleton(raw, root) {
  const patches = ['dimension', 'sheetData'].map((tag) => ({
    span: span(raw, direct(root, tag)),
    text: '<' + tag + '/>',
  }));
  return patch(raw, patches);
}
function patch(raw, patches) {
  for (const {
    span: [a, b],
    text,
  } of patches.sort((a, b) => b.span[0] - a.span[0]))
    raw = raw.slice(0, a) + text + raw.slice(b);
  return raw;
}
export function qualifyPivotTarget(raw, canonical, output) {
  const root = parse(raw),
    expected = parse(canonical),
    b = rangeBounds(output),
    d = rangeBounds(direct(root, 'dimension').getAttribute('ref'));
  if (
    d.first.row !== 0 ||
    d.first.col !== 0 ||
    d.last.row < b.last.row ||
    d.last.col < b.last.col ||
    skeleton(raw, root) !== skeleton(canonical, expected)
  )
    throw refusal('pivot-protected', 'Pivot output display/metadata is protected');
  const all = cells(root);
  for (const entry of all.entries) {
    if (!contains(d, entry.p)) throw refusal('pivot-refused', 'Pivot target dimension omits a source cell');
    if (!contains(b, entry.p)) ordinary(entry.cell);
  }
  return { root, ...all };
}
export function replacePivotOutput(raw, canonical, oldOutput, newWorksheet, newOutput) {
  const before = qualifyPivotTarget(raw, canonical, oldOutput),
    old = rangeBounds(oldOutput),
    next = rangeBounds(newOutput),
    generated = parse(newWorksheet),
    newCells = cells(generated),
    union = {
      first: old.first,
      last: { row: Math.max(old.last.row, next.last.row), col: Math.max(old.last.col, next.last.col) },
    };
  for (const entry of before.entries)
    if (contains(next, entry.p) && !contains(old, entry.p))
      throw refusal('pivot-conflict', 'Pivot growth intersects an occupied neighboring cell');
  const byRef = new Map(
      newCells.entries.map((entry) => [entry.ref, newWorksheet.slice(...span(newWorksheet, entry.cell))])
    ),
    patches = [];
  for (const [number, row] of before.rows) {
    let rowRaw = raw.slice(...span(raw, row)),
      rowStart = span(raw, row)[0],
      local = [];
    for (const entry of before.entries.filter((e) => e.row === row)) {
      if (!contains(union, entry.p)) continue;
      const bounds = span(raw, entry.cell);
      local.push({ span: [bounds[0] - rowStart, bounds[1] - rowStart], text: byRef.get(entry.ref) ?? '' });
      byRef.delete(entry.ref);
    }
    rowRaw = patch(rowRaw, local);
    const additions = [...byRef].filter(([ref]) => position(ref).row + 1 === number);
    if (additions.length) {
      const originalOutside = before.entries.filter((e) => e.row === row && !contains(union, e.p));
      for (const [ref, xml] of additions) {
        const col = position(ref).col,
          following = originalOutside.find((e) => e.p.col > col);
        let at = following ? rowRaw.indexOf(raw.slice(...span(raw, following.cell))) : rowRaw.lastIndexOf('</');
        if (rowRaw.endsWith('/>')) rowRaw = rowRaw.slice(0, -2) + '>' + xml + '</' + row.tagName + '>';
        else rowRaw = rowRaw.slice(0, at) + xml + rowRaw.slice(at);
        byRef.delete(ref);
      }
    }
    patches.push({ span: span(raw, row), text: rowRaw });
  }
  let updated = patch(raw, patches);
  if (byRef.size) {
    const parsed = parse(updated),
      data = direct(parsed, 'sheetData'),
      rows = cells(parsed).rows;
    const additions = new Map();
    for (const [ref, xml] of byRef) {
      const number = position(ref).row + 1;
      additions.set(number, (additions.get(number) ?? '') + xml);
    }
    const more = [];
    for (const [number, content] of [...additions].sort((a, b) => a[0] - b[0])) {
      const following = [...rows].find(([row]) => row > number),
        end = span(updated, data)[1],
        at = following ? span(updated, following[1])[0] : updated.lastIndexOf('</', end - 1);
      more.push({ span: [at, at], text: '<row r="' + number + '">' + content + '</row>' });
    }
    updated = patch(updated, more);
  }
  const parsed = parse(updated),
    dim = direct(parsed, 'dimension'),
    d = rangeBounds(dim.getAttribute('ref')),
    ref = 'A1:' + columnName(Math.max(d.last.col, next.last.col)) + (Math.max(d.last.row, next.last.row) + 1);
  return patch(updated, [{ span: span(updated, dim), text: '<dimension ref="' + ref + '"/>' }]);
}
