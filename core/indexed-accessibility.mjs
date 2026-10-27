import { FORMAT } from './format.mjs';
import path from 'node:path';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { openZipParts } from './zip-parts.mjs';
import { indexedAccessibilityObjects } from './indexed-object-authoring.mjs';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { ACCESSIBILITY_LIMITS as LIMITS, ACCESSIBILITY_RULES, accessibilityReport } from './accessibility.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  REL = 'http://schemas.openxmlformats.org/package/2006/relationships',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const children = (n) => Array.from(n.childNodes).filter((c) => c.nodeType === 1);
function xml(bytes, root, ns = NS) {
  if (!bytes) throw refusal('internal-missing-field', 'Missing table metadata.');
  const source = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('note-malformed', 'Extended XML entities are not inspected.');
  const fail = (m) => {
      throw refusal('package-malformed', m);
    },
    doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
      source,
      'application/xml'
    );
  if (doc.documentElement?.localName !== root || doc.documentElement.namespaceURI !== ns)
    throw refusal('note-unsupported', 'Unsupported table metadata root.');
  return doc.documentElement;
}
const target = (base, to) =>
  to.startsWith('/') ? to.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(base), to));
function tableHeaders(parts, view, t) {
  const sheet = view.sheets.find((s) => s.name === t.sheet);
  if (!sheet || view.tables.filter((v) => v.path === t.path).length !== 1)
    throw refusal('internal-missing-field', 'Missing/shared table ownership.');
  const root = xml(parts[t.path], 'table'),
    b = rangeBounds(t.ref),
    allowed = ['id', 'name', 'displayName', 'ref', 'headerRowCount', 'totalsRowCount', 'totalsRowShown'];
  if (
    Array.from(root.attributes).some(
      (a) => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/' && (a.namespaceURI || !allowed.includes(a.localName))
    ) ||
    !/^\d+$/.test(root.getAttribute('id')) ||
    Number(root.getAttribute('id')) < 1 ||
    root.getAttribute('ref') !== t.ref ||
    root.getAttribute('displayName') !== t.name ||
    root.getAttribute('name') !== t.name ||
    !t.name ||
    t.name.length > 255 ||
    Number(root.hasAttribute('headerRowCount') ? root.getAttribute('headerRowCount') : 1) !== 1 ||
    b.columns > LIMITS.columns
  )
    throw refusal(
      'note-limit',
      'Only ordinary tables with one visible header row and up to 1,000 columns are inspected.'
    );
  const nodes = children(root);
  if (
    nodes.some(
      (n) => n.namespaceURI !== NS || !['autoFilter', 'tableColumns', 'tableStyleInfo'].includes(n.localName)
    ) ||
    new Set(nodes.map((n) => n.localName)).size !== nodes.length
  )
    throw refusal('note-unsupported', 'Extended table metadata is not inspected.');
  if (nodes.filter((n) => n.localName !== 'tableColumns').some((n) => children(n).length))
    throw refusal('note-unsupported', 'Extended table filter/style content is not inspected.');
  const columns = nodes.find((n) => n.localName === 'tableColumns'),
    entries = columns ? children(columns) : [];
  if (
    !columns ||
    Number(columns.getAttribute('count')) !== b.columns ||
    entries.length !== b.columns ||
    entries.some(
      (n) =>
        n.namespaceURI !== NS ||
        n.localName !== 'tableColumn' ||
        children(n).length ||
        Array.from(n.attributes).some((a) => !['id', 'name'].includes(a.localName))
    )
  )
    throw refusal('note-unsupported', 'Extended/inconsistent table columns are not inspected.');
  if (
    new Set(entries.map((n) => n.getAttribute('id'))).size !== entries.length ||
    entries.some((n) => !/^\d+$/.test(n.getAttribute('id')) || Number(n.getAttribute('id')) < 1)
  )
    throw refusal('note-ambiguous', 'Ambiguous table column identity.');
  const relPath = sheet.path.replace(/([^/]+)$/, '_rels/$1.rels'),
    rels = children(xml(parts[relPath], 'Relationships', REL)),
    edges = rels.filter(
      (n) =>
        n.namespaceURI === REL &&
        n.localName === 'Relationship' &&
        n.getAttribute('Type') === R + '/table' &&
        n.getAttribute('TargetMode') !== 'External' &&
        target(sheet.path, n.getAttribute('Target')) === t.path
    );
  if (edges.length !== 1) throw refusal('internal-missing-field', 'Missing/shared table relationship.');
  const worksheet = xml(parts[sheet.path], 'worksheet'),
    links = Array.from(worksheet.getElementsByTagNameNS(NS, 'tablePart')).filter(
      (n) => n.getAttributeNS(R, 'id') === edges[0].getAttribute('Id')
    );
  if (links.length !== 1) throw refusal('note-bad-input', 'Table is not an exclusive declared worksheet table.');
  return {
    range: columnName(b.first.col) + (b.first.row + 1) + ':' + columnName(b.last.col) + (b.first.row + 1),
    cells: Array.from({ length: b.columns }, (_, i) => columnName(b.first.col + i) + (b.first.row + 1)),
  };
}
/** Every cell read uses the same captured native state. Never materialize bodies/media or recalculate. */
export async function indexedAccessibility(metadataBytes, view, { readRange, check = () => {} }) {
  check();
  const report = {
    schema: FORMAT.SCHEMA_ACCESSIBILITY_V1,
    scope: 'workbook',
    status: 'complete',
    rules: [...ACCESSIBILITY_RULES],
    findings: [],
    coverage: {
      objects: { total: 0, checked: 0, skipped: 0 },
      tables: { total: (view.tables ?? []).length, checked: 0, skipped: 0 },
      headerCellsRead: 0,
      skipped: [],
    },
  };
  const skip = (kind, sheet, name, reason) => {
    report.status = 'partial';
    if (report.coverage.skipped.length < LIMITS.skips)
      report.coverage.skipped.push({
        kind,
        sheet,
        name: name?.slice(0, 255) || null,
        reason: String(reason).slice(0, 1024),
      });
  };
  const finding = (rule, message, sheet, cell, objectId = null, tableName = null, relatedCells = []) =>
    report.findings.push({ rule, message, sheet, cell, objectId, tableName, relatedCells });
  try {
    const objects = indexedAccessibilityObjects(metadataBytes, view);
    report.coverage.objects.total = objects.length;
    for (const o of objects) {
      check();
      if (!o.managed || o.protectedReason || !o.anchor?.cell || o.kind === 'opaque') {
        report.coverage.objects.skipped++;
        skip('objects', o.sheet, o.name, o.protectedReason ?? 'Object location/model is outside coverage.');
        continue;
      }
      report.coverage.objects.checked++;
      if (!o.altText?.trim())
        finding(
          'missing-alt-text',
          'This owned object has no alternative text. A short description helps explain its purpose.',
          o.sheet,
          o.anchor.cell,
          o.id
        );
    }
  } catch (e) {
    check();
    report.coverage.objects.total = null;
    skip('objects', null, null, 'Object inventory unavailable: ' + String(e.message ?? e));
  }
  let parts;
  try {
    parts = openZipParts(metadataBytes, 16 * 1024 ** 2);
  } catch (e) {
    check();
    report.coverage.tables.skipped = report.coverage.tables.total;
    skip('table', null, null, 'Table metadata unavailable: ' + String(e.message ?? e));
    return accessibilityReport(report);
  }
  let requestedHeaders = 0;
  for (const [index, t] of (view.tables ?? []).entries()) {
    check();
    try {
      if (index >= LIMITS.tables) throw refusal('note-unsupported', 'Only the first 64 tables are inspected.');
      const header = tableHeaders(parts, view, t);
      if (requestedHeaders + header.cells.length > LIMITS.headerCells)
        throw refusal('note-limit', 'The 16,000-header-cell inspection budget is exhausted.');
      requestedHeaders += header.cells.length;
      const result = await readRange(t.sheet, header.range);
      check();
      if (
        !result?.cells ||
        typeof result.cells !== 'object' ||
        Array.isArray(result.cells) ||
        result.nextOffset != null
      )
        throw refusal('note-unsupported', 'Header read was incomplete.');
      report.coverage.headerCellsRead += header.cells.length;
      // Formulas/special cells have separate calculation authority: never mistake a stale cache for text.
      if (
        header.cells.some((c) => {
          const v = result.cells[c];
          return (
            v &&
            (v.formula ||
              v.formulaType ||
              v.spill ||
              v.dataType === 'error' ||
              (v.value !== null && v.value !== undefined && !['string', 'number', 'boolean'].includes(typeof v.value)))
          );
        })
      )
        throw refusal('note-unsupported', 'Formula, error or special headers are outside this literal header check.');
      report.coverage.tables.checked++;
      const groups = new Map();
      for (const cell of header.cells) {
        const v = result.cells[cell]?.value,
          label = v === null || v === undefined ? '' : String(v).trim();
        if (!label) {
          finding(
            'blank-table-header',
            'This table header is blank. Give each column a meaningful, distinct label.',
            t.sheet,
            cell,
            null,
            t.name
          );
          continue;
        }
        const key = label.toLowerCase(),
          group = groups.get(key) ?? [];
        group.push(cell);
        groups.set(key, group);
      }
      for (const cells of groups.values())
        if (cells.length > 1)
          finding(
            'duplicate-table-header',
            'These table headers repeat after trimming spaces and ignoring letter case. Give each column a distinct label.',
            t.sheet,
            cells[0],
            null,
            t.name,
            cells
          );
    } catch (e) {
      check();
      report.coverage.tables.skipped++;
      skip('table', t.sheet, t.name, String(e.message ?? e));
    }
  }
  check();
  return accessibilityReport(report);
}
