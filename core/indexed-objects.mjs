/** Small OOXML object/reference metadata. Literal worksheet cells stay in Rust. */
import path from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strFromU8, strToU8 } from 'fflate';
import { nativeFormula } from './native-formula.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  VML_EXCEL = 'urn:schemas-microsoft-com:office:excel',
  REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** Resolve the relationship targets of one part, by relationship type suffix. */
function related(parts, owner, suffix) {
  const rels = parts[path.posix.join(path.posix.dirname(owner), '_rels', path.posix.basename(owner) + '.rels')];
  if (!rels) return [];
  return Array.from(parse(rels).documentElement.childNodes)
    .filter(
      (n) => n.nodeType === 1 && n.getAttribute('Type')?.endsWith(suffix) && n.getAttribute('TargetMode') !== 'External'
    )
    .map((n) => n.getAttribute('Target'))
    .map((target) =>
      target.startsWith('/')
        ? target.slice(1)
        : path.posix.normalize(path.posix.join(path.posix.dirname(owner), target))
    );
}
const parse = (bytes) => new DOMParser().parseFromString(strFromU8(bytes), 'application/xml');
const serialize = (node) => new XMLSerializer().serializeToString(node);
const nodes = (doc, tag, namespace = NS) => Array.from(doc.getElementsByTagNameNS(namespace, tag));
const children = (node) => Array.from(node.childNodes).filter((n) => n.nodeType === 1);
/**
 * Rename a sheet inside a formula the reference parser cannot read (print
 * titles joined by a comma, a name whose target was deleted). A sheet name in
 * a formula is always followed by "!", quoted or not, so the qualifier can be
 * replaced in the text. Text inside double quotes is a string and is left alone.
 */
export function renameSheetInText(formula, oldName, newName) {
  // A name made of letters (in any script), digits, dots and underscores needs no quotes.
  const plain = (name) => /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name),
    quote = (name) => (plain(name) ? name : `'${name.replaceAll("'", "''")}'`),
    escaped = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
    quoted = `'${escaped(oldName.replaceAll("'", "''"))}'!`,
    bare = plain(oldName) ? `|(?<![\\p{L}\\p{N}_.'\\]])${escaped(oldName)}!` : '';
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((piece, at) => (at % 2 ? piece : piece.replace(new RegExp(quoted + bare, 'giu'), quote(newName) + '!')))
    .join('');
}

/** Rename reference-bearing metadata without touching worksheet cell graphs. */
export async function indexedRenameMetadata(bytes, plan, native, root) {
  const { unzipSync } = await import('fflate'),
    parts = unzipSync(bytes),
    items = [],
    setters = [],
    docs = new Map();
  const CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  for (const [file, data] of Object.entries(parts)) {
    if (/^xl\/charts\/[^/]+\.xml$/.test(file)) {
      // A chart's series name their cells as formulas, sheet name included.
      const chart = parse(data),
        references = nodes(chart, 'f', CHART).filter((node) => node.textContent.trim());
      if (!references.length) continue;
      docs.set(file, chart);
      for (const node of references) {
        items.push({ formula: node.textContent });
        setters.push((value) => {
          node.textContent = value;
        });
      }
      continue;
    }
    if (/pivotCache\/[^/]+\.xml$/.test(file) && plan.rename) {
      // A pivot table remembers the sheet its data is on by name. Only that
      // one attribute is changed; the rest of the part is left as it was.
      const xml = (name) =>
        name.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      const text = strFromU8(data),
        source = /<(?:\w+:)?worksheetSource\b[^>]*>/.exec(text);
      if (source?.[0].includes(` sheet="${xml(plan.rename.old)}"`))
        plan.replacements[file] =
          text.slice(0, source.index) +
          source[0].replace(` sheet="${xml(plan.rename.old)}"`, ` sheet="${xml(plan.rename.new)}"`) +
          text.slice(source.index + source[0].length);
      continue;
    }
    if (!plan.sheets.includes(file) && file !== 'xl/workbook.xml' && !/^xl\/tables\/.*\.xml$/.test(file)) continue;
    const doc = parse(plan.replacements[file] ? strToU8(plan.replacements[file]) : data);
    docs.set(file, doc);
    // Extension records are not rewritten; one that names the renamed sheet would be left wrong.
    if (
      plan.sheets.includes(file) &&
      nodes(doc, 'extLst').some((extension) => serialize(extension).includes(plan.rename?.old ?? '\u0000'))
    )
      throw refusal('object-unsupported', 'Extended worksheet references require rename qualification');
    for (const tag of ['definedName', 'formula', 'formula1', 'formula2', 'calculatedColumnFormula', 'totalsRowFormula'])
      for (const node of nodes(doc, tag)) {
        items.push({ formula: node.textContent });
        setters.push((value) => {
          node.textContent = value;
        });
      }
    for (const node of nodes(doc, 'hyperlink'))
      if (node.hasAttribute('location')) {
        items.push({ formula: node.getAttribute('location') });
        setters.push((value) => node.setAttribute('location', value));
      }
  }
  for (let offset = 0; offset < items.length; offset += 128) {
    const result = await nativeFormula(native, root, 'rename', {
      ...plan.rename,
      items: items.slice(offset, offset + 128),
    });
    if (result.supported) {
      result.formulas.forEach((value, i) => setters[offset + i](value));
      continue;
    }
    // One formula the parser cannot read must not stop the rename: try each on
    // its own and rename the sheet in the text of the ones it cannot read.
    for (let at = offset; at < Math.min(offset + 128, items.length); at++) {
      const single = await nativeFormula(native, root, 'rename', { ...plan.rename, items: [items[at]] });
      setters[at](
        single.supported ? single.formulas[0] : renameSheetInText(items[at].formula, plan.rename.old, plan.rename.new)
      );
    }
  }
  for (const [file, doc] of docs) {
    if (!plan.sheets.includes(file)) {
      const text = serialize(doc);
      // An unchanged chart is left exactly as it was.
      if (!/^xl\/charts\//.test(file) || text !== serialize(parse(parts[file]))) plan.replacements[file] = text;
      continue;
    }
    const fragments = (plan.fragments[file] ??= {});
    for (const tag of ['conditionalFormatting', 'dataValidations', 'hyperlinks']) {
      const current = children(doc.documentElement)
        .filter((n) => n.namespaceURI === NS && n.localName === tag)
        .map(serialize)
        .join('');
      const before = children(parse(parts[file]).documentElement)
        .filter((n) => n.namespaceURI === NS && n.localName === tag)
        .map(serialize)
        .join('');
      if (current !== before) fragments[tag] = current;
    }
  }
  return plan;
}
