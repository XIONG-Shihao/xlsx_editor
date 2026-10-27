/** Clone a qualified worksheet's drawing/chart ownership graph. Worksheet cells
 * and binary images remain in the native ZIP writer; only XML metadata is parsed. */
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { nativeFormula } from './native-formula.mjs';
import { renameSheetInText } from './indexed-objects.mjs';
import { refusal } from './refusal-codes.mjs';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships',
  CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const serialize = (node) => new XMLSerializer().serializeToString(node);
const parse = (bytes) => {
  const text = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw refusal('package-malformed', 'XML entities are unsupported');
  return new DOMParser({
    errorHandler: {
      warning() {},
      error(m) {
        throw refusal('package-malformed', m);
      },
      fatalError(m) {
        throw refusal('package-malformed', m);
      },
    },
  }).parseFromString(text, 'application/xml');
};
const relsPath = (part) => path.posix.join(path.posix.dirname(part), '_rels', path.posix.basename(part) + '.rels');
export async function indexedCopyRelationships(bytes, plan, command, native, root) {
  const parts = unzipSync(bytes),
    targetSheet = Object.keys(plan.copies).find((name) => plan.copies[name] === plan.sheet);
  if (!targetSheet) throw refusal('internal-missing-field', 'Missing copied worksheet mapping');
  // Clone local names with the copied worksheet's new localSheetId. Global
  // bindings retain their original definitions and references.
  const workbook = parse(new TextEncoder().encode(plan.replacements['xl/workbook.xml']));
  const sheets = Array.from(workbook.getElementsByTagNameNS('*', 'sheet'));
  const oldIndex = sheets.findIndex((s) => s.getAttribute('name') === command.sheet),
    newIndex = sheets.findIndex((s) => s.getAttribute('name') === command.newName);
  if (oldIndex < 0 || newIndex < 0) throw refusal('internal-missing-field', 'Missing copied worksheet name scope');
  const names = Array.from(workbook.getElementsByTagNameNS('*', 'definedName')).filter(
    (n) => n.getAttribute('localSheetId') === String(oldIndex)
  );
  if (names.length > 1024) throw refusal('name-limit', 'Copied worksheet names exceed qualification budget');
  for (let at = 0; at < names.length; at += 128) {
    const batch = names.slice(at, at + 128),
      renamed = await nativeFormula(native, root, 'rename', {
        old: command.sheet,
        new: command.newName,
        items: batch.map((n) => ({ formula: n.textContent })),
      });
    // A name the parser cannot read (print titles joined by a comma) is renamed in its text.
    const formulas = renamed.supported
      ? renamed.formulas
      : batch.map((n) => renameSheetInText(n.textContent, command.sheet, command.newName));
    batch.forEach((n, i) => {
      const copy = n.cloneNode(true);
      copy.setAttribute('localSheetId', String(newIndex));
      copy.textContent = formulas[i];
      n.parentNode.appendChild(copy);
    });
  }
  if (names.length) plan.replacements['xl/workbook.xml'] = serialize(workbook);
  const types = parse(new TextEncoder().encode(plan.replacements['[Content_Types].xml'])),
    mapped = new Map([[plan.sheet, targetSheet]]),
    visiting = new Set();
  let metadataBytes = 0;
  // Names and ids already used by tables, and names used by defined names.
  const takenNames = new Set(
    Array.from(workbook.getElementsByTagNameNS('*', 'definedName')).map((n) =>
      (n.getAttribute('name') ?? '').toLowerCase()
    )
  );
  let lastTableId = 0;
  // The number of worksheets once the copy is in (chart sheets are not
  // counted): the number Excel puts after the name.
  const links =
    plan.replacements['xl/_rels/workbook.xml.rels'] ??
    strFromU8(parts['xl/_rels/workbook.xml.rels'] ?? new Uint8Array());
  const copyNumber = (links.match(/relationships\/worksheet"/g) ?? []).length || sheets.length;
  for (const [file, data] of Object.entries(parts)) {
    if (!/^xl\/tables\/[^/]+\.xml$/.test(file)) continue;
    const head = /<(?:\w+:)?table\b[^>]*>/.exec(strFromU8(data))?.[0] ?? '';
    lastTableId = Math.max(lastTableId, Number(/\sid="(\d+)"/.exec(head)?.[1] ?? 0));
    for (const key of ['name', 'displayName'])
      takenNames.add((new RegExp(`\\s${key}="([^"]*)"`).exec(head)?.[1] ?? '').toLowerCase());
  }
  const quoted = `'${command.sheet.replaceAll("'", "''")}'!`,
    namesSheet = (text) => text.includes(quoted) || text.includes(command.sheet + '!');
  const all = new Set([...Object.keys(parts), ...Object.keys(plan.replacements), ...Object.keys(plan.copies)]);
  function allocate(source) {
    const ext = path.posix.extname(source),
      base = source.slice(0, -ext.length);
    let serial = 1,
      target;
    do {
      target = `${base}-copy${serial++}${ext}`;
    } while (all.has(target));
    all.add(target);
    return target;
  }
  function contentType(source, target) {
    const entries = Array.from(types.getElementsByTagName('Override')),
      entry = entries.find((n) => n.getAttribute('PartName') === '/' + source);
    if (entry) {
      const cloned = entry.cloneNode(false);
      cloned.setAttribute('PartName', '/' + target);
      types.documentElement.appendChild(cloned);
    }
  }
  async function clone(source, kind) {
    if (visiting.has(source)) throw refusal('object-malformed', 'Cyclic drawing relationships are not qualified');
    if (mapped.has(source)) return mapped.get(source);
    if (!Object.hasOwn(parts, source)) throw refusal('internal-missing-field', 'Missing copied relationship target');
    metadataBytes += parts[source].byteLength;
    if (metadataBytes > 16 * 1024 * 1024 || mapped.size >= 128)
      throw refusal('object-limit', 'Copied object metadata exceeds its qualification budget');
    const target = allocate(source);
    mapped.set(source, target);
    visiting.add(source);
    contentType(source, target);
    if (kind === 'chart') {
      const doc = parse(parts[source]);
      if (doc.documentElement.namespaceURI !== CHART || doc.documentElement.localName !== 'chartSpace')
        throw refusal('chart-malformed', 'Invalid chart relationship target');
      // Extension records inside a chart are copied with it, unless one of
      // them names the copied sheet: that reference would be left pointing at
      // the original.
      if (
        doc.getElementsByTagNameNS(CHART, 'externalData').length ||
        doc.getElementsByTagNameNS(CHART, 'pivotSource').length ||
        Array.from(doc.getElementsByTagNameNS('*', 'extLst')).some((extension) => namesSheet(serialize(extension)))
      )
        throw refusal('pivot-unsupported', 'External, pivot or extended chart copying is not qualified');
      const nodes = Array.from(doc.getElementsByTagNameNS(CHART, 'f'));
      for (let at = 0; at < nodes.length; at += 128) {
        const result = await nativeFormula(native, root, 'rename', {
          old: command.sheet,
          new: command.newName,
          items: nodes.slice(at, at + 128).map((n) => ({ formula: n.textContent })),
        });
        if (!result.supported)
          throw refusal('chart-unsupported', 'Chart reference copying is not qualified: ' + result.reason);
        result.formulas.forEach((f, i) => {
          nodes[at + i].textContent = f;
        });
      }
      plan.replacements[target] = serialize(doc);
    } else if (kind === 'table') {
      // A table's name and id must be unique in the workbook, so the copy gets
      // its own. Formulas on the copied sheet that name the table follow it.
      const doc = parse(parts[source]),
        table = doc.documentElement;
      if (table.localName !== 'table') throw refusal('table-malformed', 'Invalid table relationship target');
      const old = table.getAttribute('displayName') || table.getAttribute('name');
      if (!old) throw refusal('name-bad-input', 'The table has no name');
      // Excel names the copy after the original plus the number of sheets the
      // workbook then has ("Table1" copied into a third sheet becomes "Table13").
      let serial = copyNumber;
      while (takenNames.has((old + serial).toLowerCase())) serial++;
      const name = old + serial;
      takenNames.add(name.toLowerCase());
      table.setAttribute('id', String(++lastTableId));
      table.setAttribute('name', name);
      table.setAttribute('displayName', name);
      // Excel's revision ids identify one object each; the copy must not reuse them.
      for (const node of [table, ...Array.from(table.getElementsByTagNameNS('*', '*'))])
        for (const attribute of Array.from(node.attributes ?? []))
          if (/(^|:)uid$/.test(attribute.name)) node.removeAttribute(attribute.name);
      (plan.tableRenames ??= []).push({ old, new: name });
      plan.replacements[target] = serialize(doc);
    } else {
      if (kind === 'drawing') {
        const doc = parse(parts[source]);
        if (
          doc.documentElement.namespaceURI !== 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' ||
          doc.documentElement.localName !== 'wsDr'
        )
          throw refusal('object-malformed', 'Invalid drawing relationship target');
        // Extension records on a picture or shape are its own and copy with
        // it. A slicer or timeline belongs to the whole workbook and does not.
        if (/<(?:\w+:)?(?:slicer|timeslicer)\b/.test(strFromU8(parts[source])))
          throw refusal('workbook-unsupported', 'A sheet with a slicer or timeline cannot be copied');
      }
      plan.copies[target] = source;
    }
    await relationships(source, target, kind);
    visiting.delete(source);
    return target;
  }
  async function relationships(source, target, owner) {
    const oldRels = relsPath(source);
    if (!parts[oldRels]) return;
    const doc = parse(parts[oldRels]);
    for (const rel of Array.from(doc.getElementsByTagNameNS(REL, 'Relationship'))) {
      // A web link's address is the same on the copy; there is no part to copy.
      if (rel.getAttribute('TargetMode') === 'External') {
        if (owner === 'sheet' && rel.getAttribute('Type')?.endsWith('/hyperlink')) continue;
        throw refusal('package-unsupported', 'External copied relationships are not qualified');
      }
      const raw = rel.getAttribute('Target');
      if (!raw || raw.includes('\\') || raw.includes('#') || raw.includes('?') || raw.includes('%'))
        throw refusal('package-unsupported', 'Unqualified copied relationship target');
      const resolved = raw.startsWith('/')
        ? raw.slice(1)
        : path.posix.normalize(path.posix.join(path.posix.dirname(source), raw));
      if (resolved.startsWith('../') || !Object.hasOwn(parts, resolved))
        throw refusal('internal-missing-field', 'Missing copied relationship target');
      const type = rel.getAttribute('Type'),
        kind = type.slice(type.lastIndexOf('/') + 1);
      let next;
      const prefix = ['chartStyle', 'chartColorStyle'].includes(kind)
        ? 'http://schemas.microsoft.com/office/2011/relationships/'
        : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
      if (type !== prefix + kind) throw refusal('internal-unknown-reference', 'Unknown copied relationship type');
      if (owner === 'sheet' && kind === 'drawing') next = await clone(resolved, kind);
      // Notes: the list of notes and the drawing of their boxes, both copied whole.
      else if (owner === 'sheet' && ['comments', 'vmlDrawing', 'printerSettings', 'table'].includes(kind))
        next = await clone(resolved, kind);
      else if (owner === 'drawing' && kind === 'chart') next = await clone(resolved, kind);
      else if (['drawing', 'chart'].includes(owner) && kind === 'image')
        next = resolved; // Immutable resource shared by independent drawings.
      else if (owner === 'chart' && ['chartStyle', 'chartColorStyle'].includes(kind))
        next = await clone(resolved, kind);
      else throw refusal('package-unsupported', 'Unqualified copied relationship: ' + kind);
      rel.setAttribute('Target', path.posix.relative(path.posix.dirname(target), next));
    }
    plan.replacements[relsPath(target)] = serialize(doc);
  }
  await relationships(plan.sheet, targetSheet, 'sheet');
  if (plan.tableRenames?.length) {
    // The copy's own names (print area and so on) that mention a copied table.
    const book = parse(new TextEncoder().encode(plan.replacements['xl/workbook.xml']));
    for (const name of Array.from(book.getElementsByTagNameNS('*', 'definedName')))
      if (name.getAttribute('localSheetId') === String(newIndex))
        for (const { old, new: next } of plan.tableRenames)
          name.textContent = name.textContent
            .split(/("(?:[^"]|"")*")/)
            .map((piece, at) =>
              at % 2
                ? piece
                : piece.replace(
                    new RegExp(`(?<![\\p{L}\\p{N}_.])${old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=\\[)`, 'giu'),
                    next
                  )
            )
            .join('');
    plan.replacements['xl/workbook.xml'] = serialize(book);
  }
  plan.replacements['[Content_Types].xml'] = serialize(types);
  return plan;
}
