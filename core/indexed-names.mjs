/** Bounded static named-range authoring. Reference rewrites remain protected. */
import { nodeStart } from './xml-position.mjs';
import { DOMParser } from '@xmldom/xmldom';
import { unzipSync, strFromU8 } from 'fflate';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import { packageReadOnlyReasons } from './package-capabilities.mjs';
import { validateDefinedName, definedNameConflict } from './defined-name.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const fold = (value) => value.replace(/[A-Z]/g, (char) => char.toLowerCase());
const children = (node) => Array.from(node.childNodes).filter((item) => item.nodeType === 1);
const sameScope = (node, scope) =>
  scope === undefined
    ? !node.hasAttribute('localSheetId')
    : /^\d+$/.test(node.getAttribute('localSheetId')) && Number(node.getAttribute('localSheetId')) === scope;
const escapeText = (value) => value.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[char]);
function rawSpan(source, node) {
  const start = nodeStart(source, node);
  if (!Number.isSafeInteger(start) || !source.startsWith('<' + node.tagName, start))
    throw refusal('name-unsupported', 'Name XML has no qualified raw boundary.');
  return { start, end: xmlElementEnd(source, start) };
}
function appendName(source, workbook, container, name, scope, formula) {
  const prefix = (container ?? workbook).prefix,
    tag = (prefix ? prefix + ':' : '') + 'definedName',
    addition = `<${tag} name="${name}"${scope === undefined ? '' : ` localSheetId="${scope}"`}>${escapeText(formula)}</${tag}>`;
  if (container) {
    const { start, end } = rawSpan(source, container),
      fragment = source.slice(start, end);
    if (/\/\s*>$/.test(fragment))
      return (
        source.slice(0, start) +
        fragment.replace(/\/(\s*)>$/, '$1>') +
        addition +
        `</${container.tagName}>` +
        source.slice(end)
      );
    const close = fragment.lastIndexOf('</' + container.tagName + '>');
    if (close < 0) throw refusal('name-unsupported', 'Name container has no qualified closing boundary.');
    return source.slice(0, start + close) + addition + source.slice(start + close);
  }
  const rank = [
    'calcPr',
    'oleSize',
    'customWorkbookViews',
    'pivotCaches',
    'smartTagPr',
    'smartTagTypes',
    'webPublishing',
    'fileRecoveryPr',
    'webPublishObjects',
    'extLst',
  ];
  const following = children(workbook).find((node) => node.namespaceURI === NS && rank.includes(node.localName)),
    at = following
      ? rawSpan(source, following).start
      : rawSpan(source, workbook).end - ('</' + workbook.tagName + '>').length;
  if (at < 0 || (!following && !source.startsWith('</' + workbook.tagName + '>', at)))
    throw refusal('name-unsupported', 'Workbook has no qualified name insertion boundary.');
  const parent = (prefix ? prefix + ':' : '') + 'definedNames';
  return source.slice(0, at) + `<${parent}>${addition}</${parent}>` + source.slice(at);
}
const packageNs = 'http://schemas.openxmlformats.org/package/2006/relationships',
  typeNs = 'http://schemas.openxmlformats.org/package/2006/content-types',
  officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const benignRelations = new Set([
  ...[
    'officeDocument',
    'worksheet',
    'styles',
    'sharedStrings',
    'theme',
    'metadata',
    'printerSettings',
    'extended-properties',
  ].map((name) => officeRel + '/' + name),
  packageNs + '/metadata/core-properties',
  packageNs + '/metadata/thumbnail',
]);
function isolatedReferences(parts, view) {
  const refuse = () => {
    throw refusal('name-unsupported', 'Changing referenced names requires qualified reference rewriting.');
  };
  if ((view.tables ?? []).length || packageReadOnlyReasons(parts).length) refuse();
  const parseGraph = (bytes, ns, root) => {
    const fail = () => refuse(),
      doc = new DOMParser({ errorHandler: { warning: fail, error: fail, fatalError: fail } }).parseFromString(
        strFromU8(bytes),
        'application/xml'
      );
    if (doc.documentElement.namespaceURI !== ns || doc.documentElement.localName !== root) refuse();
    return doc.documentElement;
  };
  if (/<(?:[A-Za-z_][\w.-]*:)?(?:extLst|externalReferences|pivotCaches)\b/.test(strFromU8(parts['xl/workbook.xml'])))
    refuse();
  const types = parseGraph(parts['[Content_Types].xml'], typeNs, 'Types');
  for (const node of children(types))
    if (
      /(?:table|drawing|chart|pivot|externallink|querytable|connections|activex|vbaproject|signature)/i.test(
        node.getAttribute('ContentType')
      )
    )
      refuse();
  for (const [part, bytes] of Object.entries(parts)) {
    if (/^(?:xl\/(?:charts|tables|pivot|drawings)|customXml\/|xl\/externalLinks\/)/.test(part)) refuse();
    if (part.endsWith('.rels'))
      for (const node of children(parseGraph(bytes, packageNs, 'Relationships'))) {
        if (
          node.namespaceURI !== packageNs ||
          node.localName !== 'Relationship' ||
          children(node).length ||
          !benignRelations.has(node.getAttribute('Type')) ||
          node.getAttribute('TargetMode') === 'External'
        )
          refuse();
      }
  }
  for (const sheet of view.sheets) {
    const bytes = parts[sheet.path];
    if (
      !bytes ||
      /<(?:[A-Za-z_][\w.-]*:)?(?:formula|formula1|formula2|dataValidation|dataValidations|conditionalFormatting|tableParts|drawing|legacyDrawing|legacyDrawingHF|oleObjects|controls|extLst|hyperlinks)\b/.test(
        strFromU8(bytes)
      )
    )
      refuse();
  }
}
export function isIndexedNameCommand(command) {
  return ['defineName', 'removeName', 'updateName'].includes(command?.operation);
}
export function indexedNamePlan(metadataBytes, view, command) {
  if (
    !isIndexedNameCommand(command) ||
    Object.keys(command).some(
      (key) => !['operation', 'sheet', 'name', 'range', 'nameScope', 'expectedNameFormula'].includes(key)
    )
  )
    throw refusal('internal-invalid-field', 'Invalid named-range command.');
  if (metadataBytes.byteLength > 16 * 1024 ** 2)
    throw refusal('name-limit', 'Named-range metadata exceeds its budget.');
  const name = validateDefinedName(command.name),
    source = view.sheets.find((sheet) => sheet.name === command.sheet);
  if (!source) throw refusal('internal-unknown-reference', 'Unknown worksheet.');
  if (command.nameScope !== undefined && !['workbook', 'worksheet'].includes(command.nameScope))
    throw refusal('internal-invalid-field', 'Invalid name scope.');
  const scope = command.nameScope === 'worksheet' ? view.sheets.indexOf(source) : undefined;
  if ((view.tables ?? []).some((table) => fold(table.name) === fold(name)))
    throw refusal('name-refused', 'A table already uses that name.');
  const parts = unzipSync(metadataBytes),
    raw = parts['xl/workbook.xml'];
  if (!raw || /<!DOCTYPE/i.test(strFromU8(raw)))
    throw refusal('name-unsupported', 'Unsupported workbook name metadata.');
  const problems = [],
    doc = new DOMParser({
      locator: {},
      errorHandler: {
        warning: (message) => problems.push(message),
        error: (message) => problems.push(message),
        fatalError: (message) => problems.push(message),
      },
    }).parseFromString(strFromU8(raw), 'application/xml');
  if (problems.length || doc.documentElement.namespaceURI !== NS || doc.documentElement.localName !== 'workbook')
    throw refusal('internal-invalid-field', 'Invalid workbook name metadata.');
  const containers = children(doc.documentElement).filter((node) => node.localName === 'definedNames');
  if (containers.length > 1 || containers.some((node) => node.namespaceURI !== NS))
    throw refusal('name-ambiguous', 'Ambiguous defined-name containers.');
  let container = containers[0];
  if (
    container &&
    (Array.from(container.attributes).some(
      (attribute) => attribute.name !== 'xmlns' && !attribute.name.startsWith('xmlns:')
    ) ||
      Array.from(container.childNodes).some((node) => [3, 4].includes(node.nodeType) && node.textContent.trim()))
  )
    throw refusal('name-unsupported', 'Extended name container requires qualification.');
  const names = container ? children(container) : [];
  if (names.length >= 1000 && command.operation === 'defineName')
    throw refusal('name-limit', 'Named ranges have reached the 1,000-name limit.');
  if (names.some((node) => node.namespaceURI !== NS || node.localName !== 'definedName' || children(node).length))
    throw refusal('name-unsupported', 'Extended name metadata requires qualification.');
  const matches = names.filter((node) => fold(node.getAttribute('name')) === fold(name) && sameScope(node, scope));
  if (matches.length > 1) throw refusal('name-ambiguous', 'Ambiguous named-range identity.');
  if (command.operation === 'defineName') {
    if (command.expectedNameFormula !== null)
      throw refusal('name-bad-input', 'Creating a name requires an explicit absent-name guard.');
    const conflict = definedNameConflict(
      name,
      { namedRanges: view.namedRanges ?? [], tables: view.tables ?? [] },
      { scopeSheet: scope === undefined ? undefined : source.name }
    );
    if (conflict) throw refusal('name-refused', conflict);
    if (matches.length) throw refusal('name-duplicate', 'A name already exists in that scope.');
    const bounds = rangeBounds(command.range);
    const formula = `'${source.name.replaceAll("'", "''")}'!$${columnName(bounds.first.col)}$${bounds.first.row + 1}:$${columnName(bounds.last.col)}$${bounds.last.row + 1}`;
    const replacement = appendName(strFromU8(raw), doc.documentElement, container, name, scope, formula);
    return {
      kind: 'defineName',
      name,
      scope: scope ?? null,
      formula,
      replacements: { 'xl/workbook.xml': replacement },
      removed: [],
      copies: {},
      fragments: {},
    };
  } else {
    if (
      matches.length !== 1 ||
      typeof command.expectedNameFormula !== 'string' ||
      matches[0].textContent !== command.expectedNameFormula
    )
      throw refusal('internal-missing-field', 'Conflict: named range changed or is missing.');
    if (matches[0].attributes.length !== (scope === undefined ? 1 : 2))
      throw refusal('name-unsupported', 'Extended named-range removal requires qualification.');
    // Several independent static names are ordinary, not a dependency graph.
    // Require every neighbor to be a unique, qualified static rectangle. Formula
    // consumers are checked independently against the accepted native export.
    const identities = new Set();
    for (const node of names) {
      const builtin = /^_xlnm\.Print_(?:Area|Titles)$/i.test(node.getAttribute('name'));
      if (!builtin) validateDefinedName(node.getAttribute('name'));
      if (
        Array.from(node.attributes).some((a) => !['name', 'localSheetId'].includes(a.name)) ||
        node.childNodes.length !== 1 ||
        node.firstChild?.nodeType !== 3
      )
        throw refusal('name-unsupported', 'Extended or dependent names require qualification.');
      const local = node.hasAttribute('localSheetId') ? node.getAttribute('localSheetId') : null;
      if (
        local !== null &&
        (!/^\d+$/.test(local) || !Number.isSafeInteger(Number(local)) || Number(local) >= view.sheets.length)
      )
        throw refusal('internal-invalid-field', 'Invalid named-range scope.');
      const identity = fold(node.getAttribute('name')) + '\0' + (local === null ? 'global' : Number(local));
      if (identities.has(identity)) throw refusal('name-ambiguous', 'Ambiguous named-range identity.');
      identities.add(identity);
      if (builtin) {
        if (local === null) throw refusal('name-bad-input', 'Print names require worksheet scope.');
        const prefix = "'" + view.sheets[Number(local)].name.replaceAll("'", "''") + "'!";
        let remaining = node.textContent,
          count = 0;
        for (;;) {
          if (!remaining.startsWith(prefix))
            throw refusal('name-bad-input', 'Print name must use its actual scoped worksheet.');
          remaining = remaining.slice(prefix.length);
          const comma = remaining.indexOf(','),
            part = comma < 0 ? remaining : remaining.slice(0, comma);
          if (/Print_Area$/i.test(node.getAttribute('name'))) {
            if (count || !/^\$[A-Z]{1,3}\$[1-9]\d*:\$[A-Z]{1,3}\$[1-9]\d*$/.test(part))
              throw refusal('name-unsupported', 'Unqualified print area.');
            rangeBounds(part.replaceAll('$', ''));
          } else {
            let m = /^\$([1-9]\d*):\$([1-9]\d*)$/.exec(part);
            if (m) rangeBounds('A' + m[1] + ':A' + m[2]);
            else {
              m = /^\$([A-Z]{1,3}):\$([A-Z]{1,3})$/.exec(part);
              if (!m) throw refusal('name-unsupported', 'Unqualified print title.');
              rangeBounds(m[1] + '1:' + m[2] + '1');
            }
          }
          if (++count > 2) throw refusal('name-unsupported', 'Excessive print-name areas.');
          if (comma < 0) break;
          remaining = remaining.slice(comma + 1);
        }
        continue;
      }
      const ref = /^'((?:[^']|'')+)'!\$([A-Z]{1,3})\$([1-9]\d*):\$([A-Z]{1,3})\$([1-9]\d*)$/.exec(node.textContent);
      if (!ref || !view.sheets.some((sheet) => sheet.name === ref[1].replaceAll("''", "'")))
        throw refusal('name-unsupported', 'Dependent names require qualified reference rewriting.');
      rangeBounds(ref[2] + ref[3] + ':' + ref[4] + ref[5]);
    }
    isolatedReferences(parts, view);
    if (command.operation === 'updateName') {
      if (matches[0].childNodes.length !== 1 || matches[0].firstChild?.nodeType !== 3)
        throw refusal('name-unsupported', 'Extended static name text requires qualification.');
      const prior = matches[0].textContent,
        old = /^'((?:[^']|'')+)'!\$([A-Z]{1,3})\$([1-9]\d*):\$([A-Z]{1,3})\$([1-9]\d*)$/.exec(prior);
      if (!old || !view.sheets.some((sheet) => sheet.name === old[1].replaceAll("''", "'")))
        throw refusal('name-unsupported', 'Only a canonical static rectangle can be updated.');
      rangeBounds(`${old[2]}${old[3]}:${old[4]}${old[5]}`);
      const bounds = rangeBounds(command.range),
        formula = `'${source.name.replaceAll("'", "''")}'!$${columnName(bounds.first.col)}$${bounds.first.row + 1}:$${columnName(bounds.last.col)}$${bounds.last.row + 1}`;
      const sourceXml = strFromU8(raw),
        start = nodeStart(sourceXml, matches[0]);
      if (!sourceXml.startsWith('<' + matches[0].tagName, start))
        throw refusal('name-unsupported', 'Name XML has no qualified raw boundary.');
      const end = xmlElementEnd(sourceXml, start),
        fragment = sourceXml.slice(start, end),
        opening = fragment.indexOf('>');
      if (opening < 0 || fragment[opening - 1] === '/')
        throw refusal('name-bad-input', 'Static name has no explicit body.');
      const close = fragment.lastIndexOf('</'),
        escape = (value) => value.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[char]);
      const replacement =
        sourceXml.slice(0, start) +
        fragment.slice(0, opening + 1) +
        escape(formula) +
        fragment.slice(close) +
        sourceXml.slice(end);
      return {
        kind: 'updateName',
        name: matches[0].getAttribute('name'),
        scope: scope ?? null,
        priorFormula: prior,
        formula,
        replacements: { 'xl/workbook.xml': replacement },
        removed: [],
        copies: {},
        fragments: {},
      };
    }
    const sourceXml = strFromU8(raw),
      span = rawSpan(sourceXml, matches[0]),
      replacement = sourceXml.slice(0, span.start) + sourceXml.slice(span.end);
    return {
      kind: 'removeName',
      name: matches[0].getAttribute('name'),
      scope: scope ?? null,
      priorFormula: matches[0].textContent,
      replacements: { 'xl/workbook.xml': replacement },
      removed: [],
      copies: {},
      fragments: {},
    };
  }
}
