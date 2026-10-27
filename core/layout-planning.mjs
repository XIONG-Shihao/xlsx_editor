// Pure worksheet geometry, shared by the existing package writer and the
// non-Node JavaScriptCore helper. No filesystem, process or publication ports.
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const children = (parent, name) =>
  Array.from(parent?.childNodes ?? []).filter(
    (n) => n.nodeType === 1 && n.namespaceURI === NS && (!name || n.localName === name)
  );
function create(doc, name, attributes = {}) {
  const prefix = doc.documentElement.prefix;
  const node = doc.createElementNS(NS, prefix ? `${prefix}:${name}` : name);
  for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
  return node;
}
function unique(parent, name) {
  const found = children(parent, name);
  if (found.length > 1) throw refusal('layout-ambiguous', `validation: Ambiguous worksheet ${name}`);
  return found[0];
}
export function resizeColumnXml(doc, edit) {
  if (
    !Number.isInteger(edit.index) ||
    edit.index < 0 ||
    edit.index >= 16384 ||
    !Number.isFinite(edit.width) ||
    edit.width < 8 ||
    edit.width > 100
  )
    throw refusal('internal-invalid-field', 'validation: Invalid column width');
  let cols = unique(doc.documentElement, 'cols');
  if (!cols) {
    cols = create(doc, 'cols');
    doc.documentElement.insertBefore(cols, unique(doc.documentElement, 'sheetData') ?? null);
  }
  const coordinate = edit.index + 1;
  let modified = false,
    previous = 0;
  for (const col of [...children(cols, 'col')]) {
    const first = Number(col.getAttribute('min')),
      last = Number(col.getAttribute('max'));
    if (!Number.isInteger(first) || !Number.isInteger(last) || first <= previous || last < first || last > 16384)
      throw refusal('layout-ambiguous', 'validation: Ambiguous column spans');
    previous = last;
    if (coordinate < first || coordinate > last) continue;
    if (first < coordinate) {
      const left = col.cloneNode(true);
      left.setAttribute('max', String(coordinate - 1));
      cols.insertBefore(left, col);
    }
    if (last > coordinate) {
      const right = col.cloneNode(true);
      right.setAttribute('min', String(coordinate + 1));
      cols.insertBefore(right, col.nextSibling);
    }
    col.setAttribute('min', String(coordinate));
    col.setAttribute('max', String(coordinate));
    col.setAttribute('width', String(edit.width));
    col.setAttribute('customWidth', '1');
    modified = true;
  }
  if (!modified)
    cols.insertBefore(
      create(doc, 'col', { min: coordinate, max: coordinate, width: edit.width, customWidth: 1 }),
      children(cols, 'col').find((col) => Number(col.getAttribute('min')) > coordinate) ?? null
    );
}
export function freezeXml(doc, edit) {
  if (
    !Number.isInteger(edit.rows) ||
    edit.rows < 0 ||
    edit.rows > 100 ||
    !Number.isInteger(edit.columns) ||
    edit.columns < 0 ||
    edit.columns > 30
  )
    throw refusal('layout-malformed', 'validation: Invalid freeze boundary');
  let views = unique(doc.documentElement, 'sheetViews');
  if (!views) {
    views = create(doc, 'sheetViews');
    doc.documentElement.insertBefore(
      views,
      children(doc.documentElement).find((n) => !['sheetPr', 'dimension'].includes(n.localName)) ?? null
    );
  }
  let view = children(views, 'sheetView')[0];
  if (!view) {
    view = create(doc, 'sheetView', { workbookViewId: 0 });
    views.appendChild(view);
  }
  for (const pane of children(view, 'pane')) view.removeChild(pane);
  for (const selection of children(view, 'selection')) selection.removeAttribute('pane');
  if (edit.rows || edit.columns) {
    let column = '',
      at = edit.columns + 1;
    while (at) {
      at--;
      column = String.fromCharCode(65 + (at % 26)) + column;
      at = Math.floor(at / 26);
    }
    view.insertBefore(
      create(doc, 'pane', {
        xSplit: edit.columns,
        ySplit: edit.rows,
        topLeftCell: `${column}${edit.rows + 1}`,
        state: 'frozen',
        activePane: edit.rows && edit.columns ? 'bottomRight' : edit.rows ? 'bottomLeft' : 'topRight',
      }),
      view.firstChild
    );
  }
}
export function planWorksheetLayout(worksheetXml, commands) {
  if (
    typeof worksheetXml !== 'string' ||
    worksheetXml.length > 16 * 1024 * 1024 ||
    !Array.isArray(commands) ||
    !commands.length ||
    commands.length > 10000
  )
    throw refusal('layout-limit', 'validation: Layout planning exceeds its bounds');
  if (/<!DOCTYPE|<!ENTITY/i.test(worksheetXml))
    throw refusal('layout-malformed', 'validation: Layout entities are unsupported');
  const fail = (message) => {
    throw refusal('layout-malformed', 'validation: Invalid layout XML: ' + message);
  };
  const doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
    worksheetXml,
    'application/xml'
  );
  if (doc.documentElement.namespaceURI !== NS || doc.documentElement.localName !== 'worksheet')
    throw refusal('internal-invalid-field', 'validation: Invalid layout worksheet');
  const tags = new Set();
  for (const edit of commands) {
    if (!edit || typeof edit !== 'object' || Array.isArray(edit))
      throw refusal('internal-invalid-field', 'validation: Invalid layout command');
    const fields =
      edit.operation === 'resizeColumn'
        ? ['operation', 'index', 'width']
        : edit.operation === 'freeze'
          ? ['operation', 'rows', 'columns']
          : null;
    if (!fields || Object.keys(edit).length !== fields.length || Object.keys(edit).some((key) => !fields.includes(key)))
      throw refusal('layout-unsupported', 'validation: Unqualified layout command');
    if (edit.operation === 'resizeColumn') {
      resizeColumnXml(doc, edit);
      tags.add('cols');
    } else {
      freezeXml(doc, edit);
      tags.add('sheetViews');
    }
  }
  const encoder = new XMLSerializer();
  return {
    fragments: [...tags].map((tag) => ({ tag, xml: encoder.serializeToString(unique(doc.documentElement, tag)) })),
  };
}
