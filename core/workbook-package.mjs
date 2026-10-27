import { lineStarts } from './xml-position.mjs';
import { refusal } from './refusal-codes.mjs';
import { ownedPageThemePart } from './page-theme-command.mjs';
import { pageThemeSamples, pageThemeXml } from './page-style-samples.mjs';
import { openZipParts, writeZipParts, sameZipMember, equalBytes } from './zip-parts.mjs';
import { readWorkbookTheme, resolveThemeColor, resolveThemeFont } from './workbook-theme.mjs';
/** Shared offline workbook adapter for the editor and MCP. Untouched ZIP members
 * remain byte-identical after decompression; edits touch only relevant XML parts.
 * No workbook is passed through a lossy generic import/export representation. */
import { zipSync, strFromU8, strToU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { relationshipPart, resolvePackageTarget } from './package-paths.mjs';
import { resolveWorkbookEdits, assertXmlText } from './workbook-commands.mjs';
import { applyWorkbookStyle, appendDifferentialFill, createWorkbookStyleWriter } from './workbook-styles.mjs';
import { resolveWorkbookStyles } from './style-publication.mjs';
import { resizeColumnXml, freezeXml } from './layout-planning.mjs';
import { readSharedStringsXml, readWorksheetXml, xmlElementEnd } from './worksheet-xml.mjs';
import { readConditionalRules, readValidationRules } from './rules-xml.mjs';
import { hasMacroProject, packageReadOnlyReasons } from './package-capabilities.mjs';
import { normalizeSortKeys } from './sort-keys.mjs';
import {
  inspectFilter,
  readFilterRecord,
  captureFilterBaseline,
  prepareFilterState,
  appendFilterRecord,
} from './filter-model.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
// Compatibility export: workbook validity follows ZIP/OOXML bounds. Resource
// policy belongs to the caller and must not silently truncate a workbook.
export const MAX_WORKBOOK_BYTES = Infinity;
const encoder = new XMLSerializer();
const viewContexts = new WeakMap();
function serializeXml(document) {
  return encoder.serializeToString(document, false, (item) => {
    if (item.nodeType !== 3) return item;
    assertXmlText(item.data);
    // Raw CR is normalized to LF by conforming XML readers. Character
    // references preserve the exact text through Excel/independent readback.
    return item.data.replace(/[&<>\r]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#13;' })[char]);
  });
}
const hash = (bytes) => bytesToHex(sha256(bytes));
const elements = (node, name) => Array.from(node.getElementsByTagNameNS(NS, name));
const children = (node, name) =>
  Array.from(node?.childNodes ?? []).filter(
    (n) => n.nodeType === 1 && (!name || (n.namespaceURI === NS && n.localName === name))
  );
const text = (node, name) => children(node, name)[0]?.textContent ?? '';
const sheetData = (document) => children(document.documentElement, 'sheetData')[0];
const sheetRows = (document) => children(sheetData(document), 'row');
const sheetCells = (document) => sheetRows(document).flatMap((row) => children(row, 'c'));
const appendCellData = (cell, value) => cell.insertBefore(value, children(cell, 'extLst')[0] ?? null);
export const encodeWorkbookString = (value) => String(value).replace(/_(?=x[0-9a-fA-F]{4}_)/g, '_x005F_');
const decodeWorkbookString = (value) =>
  value.replace(/_x([0-9a-fA-F]{4})_/g, (_match, code) => String.fromCharCode(parseInt(code, 16)));
// Phonetic guides annotate pronunciation; they are not displayed cell text.
const cellText = (node) =>
  (children(node, 'is')[0] ? elements(children(node, 'is')[0], 't') : [])
    .filter((t) => {
      for (let p = t.parentNode; p && p !== node; p = p.parentNode)
        if (p.namespaceURI === NS && p.localName === 'rPh') return false;
      return true;
    })
    .map((t) => decodeWorkbookString(t.textContent))
    .join('');
function xml(bytes, options = {}) {
  const source = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('package-malformed', 'XML entities are unsupported');
  return new DOMParser({
    ...options,
    errorHandler: {
      warning() {},
      error(message) {
        throw refusal('package-malformed', message);
      },
      fatalError(message) {
        throw refusal('package-malformed', message);
      },
    },
  }).parseFromString(source, 'application/xml');
}
function node(doc, name, attrs = {}, value) {
  const n = doc.createElementNS(NS, name);
  for (const [key, val] of Object.entries(attrs)) n.setAttribute(key, String(val));
  if (value !== undefined) n.appendChild(doc.createTextNode(String(value)));
  return n;
}
export function columnName(index) {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
export function cellPosition(ref) {
  const m = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(ref);
  if (!m) throw refusal('internal-invalid-field', `Invalid cell address: ${ref}`);
  const col = [...m[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(m[2]) - 1;
  if (col > 16383 || row > 1048575) throw refusal('package-limit', 'Cell is outside workbook limits');
  return { row, col };
}
function unpack(bytes) {
  const parts = openZipParts(bytes);
  if (!parts['xl/workbook.xml'] || !parts['xl/_rels/workbook.xml.rels'])
    throw refusal('package-bad-input', 'Not a supported XLSX workbook');
  return parts;
}
function inspect(parts, previous, { stream = false, sharedStrings } = {}) {
  const documents = new Map(previous?.sheets.map((sheet) => [sheet.path, sheet.document]) ?? []);
  const workbook = xml(parts['xl/workbook.xml']);
  const rels = xml(parts['xl/_rels/workbook.xml.rels']);
  // Prefix spelling has no identity. Keep the supported transitional namespace
  // profile closed and resolve only direct, uniquely identified package edges.
  const workbookRoot = workbook.documentElement,
    relationshipsRoot = rels.documentElement;
  if (workbookRoot.namespaceURI !== NS || workbookRoot.localName !== 'workbook')
    throw refusal('package-unsupported', 'Unsupported workbook namespace/root');
  if (relationshipsRoot.namespaceURI !== PACKAGE_REL_NS || relationshipsRoot.localName !== 'Relationships')
    throw refusal('package-malformed', 'Invalid workbook relationship namespace/root');
  const relationshipNodes = Array.from(rels.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship'));
  if (
    relationshipNodes.some((node) => node.parentNode !== relationshipsRoot) ||
    children(relationshipsRoot).some(
      (node) => node.namespaceURI !== PACKAGE_REL_NS || node.localName !== 'Relationship'
    ) ||
    Array.from(relationshipsRoot.childNodes).some((node) => [3, 4].includes(node.nodeType) && node.data.trim())
  )
    throw refusal('package-unsupported', 'Foreign or nested workbook relationship graph');
  const targets = new Map();
  for (const relation of relationshipNodes) {
    const id = relation.getAttributeNS(null, 'Id'),
      type = relation.getAttributeNS(null, 'Type');
    const target = relation.getAttributeNS(null, 'Target'),
      mode = relation.getAttributeNS(null, 'TargetMode');
    if (
      !id ||
      !type ||
      !target ||
      targets.has(id) ||
      children(relation).length ||
      Array.from(relation.childNodes).some((node) => [3, 4].includes(node.nodeType) && node.data.trim())
    )
      throw refusal('internal-missing-field', 'Missing, malformed or duplicate workbook relationship');
    const path =
      target.includes('..') || (mode && !['Internal', 'External'].includes(mode)) || mode === 'External'
        ? null
        : target.startsWith('/')
          ? target.slice(1)
          : `xl/${target}`;
    targets.set(id, { type, path });
  }
  const sheetContainers = children(workbookRoot).filter((node) => node.localName === 'sheets');
  if (sheetContainers.length !== 1 || sheetContainers[0].namespaceURI !== NS)
    throw refusal('internal-missing-field', 'Missing or foreign workbook sheet inventory');
  const sheetNodes = Array.from(workbook.getElementsByTagNameNS(NS, 'sheet'));
  if (
    !sheetNodes.length ||
    sheetNodes.some((node) => node.parentNode !== sheetContainers[0]) ||
    children(sheetContainers[0]).some((node) => node.namespaceURI !== NS || node.localName !== 'sheet')
  )
    throw refusal('package-unsupported', 'Foreign or nested workbook sheet inventory');
  const sheetNames = new Set(),
    sheetPaths = new Set();
  const reasons = packageReadOnlyReasons(parts),
    unverified = [],
    protectedSheets = [];
  const protections = elements(workbook, 'workbookProtection');
  // An empty record represents an unlocked workbook. Only the
  // closed inactive form is admitted; locks, passwords, extensions, malformed
  // values and ambiguous placement retain the existing conservative fence.
  const inactiveProtection = (node) =>
    node.parentNode === workbookRoot &&
    !children(node).length &&
    !Array.from(node.childNodes).some((child) => [3, 4].includes(child.nodeType) && child.data.trim()) &&
    Array.from(node.attributes).every(
      (attribute) =>
        attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/' ||
        (!attribute.namespaceURI &&
          ['lockStructure', 'lockWindows', 'lockRevision'].includes(attribute.localName) &&
          ['0', 'false'].includes(attribute.value))
    );
  if (protections.length > 1 || protections.some((node) => !inactiveProtection(node)))
    reasons.push('Workbook protection is enabled');
  const strings =
    sharedStrings ?? (parts['xl/sharedStrings.xml'] ? readSharedStringsXml(parts['xl/sharedStrings.xml']) : []);
  const themeProjection = readWorkbookTheme(parts, targets.values());
  const themeIssues = new Set(themeProjection.issues),
    theme = themeProjection.theme;
  const themeEdges = Array.from(targets.values()).filter((edge) => edge.type === OFFICE_REL_NS + '/theme');
  let pageTheme = {
    revision: null,
    sample: null,
    canApply: themeEdges.length === 0,
    protectedReason: themeEdges.length === 0 ? null : 'Imported, ambiguous or modified themes stay protected.',
  };
  if (theme) {
    const bytes = parts[theme.part],
      sample =
        theme.part === ownedPageThemePart
          ? pageThemeSamples.find((sample) => strFromU8(bytes) === pageThemeXml(sample.id))
          : undefined;
    pageTheme = {
      revision: hash(bytes),
      sample: sample?.id ?? null,
      canApply: !!sample,
      protectedReason: sample ? null : 'Imported or modified themes stay protected.',
    };
  }

  const styles = parts['xl/styles.xml'] ? xml(parts['xl/styles.xml']) : null;
  const formats = styles
    ? Object.fromEntries(
        elements(styles, 'numFmt').map((n) => [n.getAttribute('numFmtId'), n.getAttribute('formatCode')])
      )
    : {};
  const xfs = styles
    ? children(elements(styles, 'cellXfs')[0] ?? styles, 'xf').map((n) => ({
        numFmtId: Number(n.getAttribute('numFmtId')),
        fontId: Number(n.getAttribute('fontId')),
        fillId: Number(n.getAttribute('fillId')),
        borderId: Number(n.getAttribute('borderId')),
        alignment: children(n, 'alignment')[0],
      }))
    : [];
  const fonts = styles
    ? elements(styles, 'font').map((f) => ({
        bold: elements(f, 'b').some((flag) => !['0', 'false'].includes(flag.getAttribute('val'))),
        italic: elements(f, 'i').some((flag) => !['0', 'false'].includes(flag.getAttribute('val'))),
        underline: elements(f, 'u').length > 0,
        fontFamily: resolveThemeFont(
          elements(f, 'name')[0]?.getAttribute('val'),
          elements(f, 'scheme')[0]?.getAttribute('val'),
          theme,
          themeIssues
        ),
        fontSize: Number(elements(f, 'sz')[0]?.getAttribute('val')) || undefined,
        fontColor: resolveThemeColor(elements(f, 'color')[0], theme, themeIssues),
      }))
    : [];
  const normalStyle =
    styles &&
    elements(styles, 'cellStyle').find(
      (style) => style.getAttribute('builtinId') === '0' || style.getAttribute('name') === 'Normal'
    );
  const styleBases = styles ? children(elements(styles, 'cellStyleXfs')[0], 'xf') : [];
  const normalBase = styleBases[Number(normalStyle?.getAttribute('xfId') ?? 0)];
  const normalSourceFont = fonts[Number(normalBase?.getAttribute('fontId') ?? 0)] ?? fonts[0];
  const normalGeometryIssues = [],
    normalFont = {
      fontFamily: normalSourceFont?.fontFamily ?? 'Calibri',
      fontSize: normalSourceFont?.fontSize ?? 11,
      bold: !!normalSourceFont?.bold,
      italic: !!normalSourceFont?.italic,
    };
  if (
    !normalFont.fontFamily.trim() ||
    normalFont.fontFamily.length > 256 ||
    /[\x00-\x1f\x7f]/u.test(normalFont.fontFamily)
  ) {
    normalGeometryIssues.push('Normal font family is unresolved; its display uses Calibri');
    normalFont.fontFamily = 'Calibri';
  }
  if (!Number.isFinite(normalFont.fontSize) || normalFont.fontSize <= 0 || normalFont.fontSize > 409.5) {
    normalGeometryIssues.push('Normal font size is unresolved; its display uses 11 points');
    normalFont.fontSize = 11;
  }
  const fills = styles
    ? elements(styles, 'fill').map((f) => resolveThemeColor(elements(f, 'fgColor')[0], theme, themeIssues))
    : [];
  const borders = styles
    ? elements(styles, 'border').map((b) =>
        ['left', 'right', 'top', 'bottom'].some((side) => elements(b, side)[0]?.hasAttribute('style')) ? 'all' : 'none'
      )
    : [];
  // Resolve each style once, before scanning cells. Undefined optional fields
  // carry no workbook information but otherwise occupy slots in every cell and
  // are serialized across the desktop worker/renderer boundary millions of times.
  const cellStyles = xfs.map((style) =>
    Object.fromEntries(
      Object.entries({
        format: formats[style.numFmtId] ?? style.numFmtId,
        ...fonts[style.fontId],
        fillColor: fills[style.fillId] ?? undefined,
        border: borders[style.borderId] ?? undefined,
        horizontalAlign: style.alignment?.getAttribute('horizontal') ?? undefined,
        verticalAlign: style.alignment?.getAttribute('vertical') ?? undefined,
        wrapText: ['1', 'true'].includes(style.alignment?.getAttribute('wrapText')),
      }).filter(([, value]) => value !== undefined && value !== '')
    )
  );
  const differentialFills = styles
    ? children(elements(styles, 'dxfs')[0], 'dxf').map((dxf) =>
        resolveThemeColor(elements(dxf, 'fgColor')[0] ?? elements(dxf, 'bgColor')[0], theme, themeIssues)
      )
    : [];
  unverified.push(...themeIssues);
  const date1904 = ['1', 'true'].includes(elements(workbook, 'workbookPr')[0]?.getAttribute('date1904'));
  const sheets = sheetNodes.map((sheet) => {
    const identities = Array.from(sheet.attributes).filter(
      (attribute) => attribute.namespaceURI === OFFICE_REL_NS && attribute.localName === 'id'
    );
    const relation = identities.length === 1 ? targets.get(sheet.getAttributeNS(OFFICE_REL_NS, 'id')) : undefined;
    const path = relation?.type === OFFICE_REL_NS + '/worksheet' ? relation.path : null;
    if (!path || !parts[path]) throw refusal('package-refused', 'Workbook contains a missing or external sheet');
    const name = sheet.getAttributeNS(null, 'name');
    if (!name || sheetNames.has(name.toLowerCase()) || sheetPaths.has(path))
      throw refusal('package-refused', 'Workbook contains duplicate or missing worksheet identities');
    sheetNames.add(name.toLowerCase());
    sheetPaths.add(path);
    let document, cells, scan;
    let rows = 20,
      columns = 8;
    const parseCell = (record) => {
      const { attrs, valueText: raw, formulaAttrs } = record,
        ref = attrs.r;
      const position = cellPosition(ref);
      rows = Math.max(rows, position.row + 1);
      columns = Math.max(columns, position.col + 1);
      const type = attrs.t || 'n';
      const value =
        type === 's'
          ? (strings[Number(raw)] ?? '')
          : type === 'inlineStr'
            ? record.inlineText
            : type === 'b'
              ? raw === '1'
              : type === 'n' && raw !== ''
                ? Number(raw)
                : type === 'n'
                  ? null
                  : type === 'str'
                    ? decodeWorkbookString(raw)
                    : raw;
      const styleId = Number(attrs.s || 0);
      const formula = record.formula || undefined;
      const dataType = ['s', 'str', 'inlineStr'].includes(type)
        ? 'text'
        : type === 'b'
          ? 'boolean'
          : type === 'e'
            ? 'error'
            : raw === ''
              ? 'blank'
              : 'number';
      return {
        value,
        ...(formula ? { formula } : {}),
        ...(formulaAttrs?.t !== undefined
          ? {
              formulaType: formulaAttrs.t,
              sharedFormulaId: formulaAttrs.si || undefined,
              formulaReference: formulaAttrs.ref || undefined,
            }
          : {}),
        dataType,
        rawValue: ['s', 'inlineStr'].includes(type) ? null : record.valuePresent ? raw : null,
        cachedValue: formulaAttrs && record.valuePresent ? value : null,
        ...(type === 'e' ? { error: raw } : {}),
        styleId,
        ...(cellStyles[styleId] ?? { wrapText: false }),
      };
    };
    if (stream) {
      scan = readWorksheetXml(parts[path], parseCell);
      document = scan.document;
      cells = scan.cells;
    } else {
      document = documents.get(path) ?? xml(parts[path]);
      cells = {};
      for (const c of sheetCells(document)) {
        const formula = children(c, 'f')[0];
        cells[c.getAttribute('r')] = parseCell({
          attrs: Object.fromEntries(Array.from(c.attributes).map((attr) => [attr.name, attr.value])),
          valueText: text(c, 'v'),
          valuePresent: children(c, 'v').length > 0,
          inlineText: cellText(c),
          formula: formula?.textContent,
          formulaAttrs: formula
            ? Object.fromEntries(Array.from(formula.attributes).map((attr) => [attr.name, attr.value]))
            : undefined,
        });
      }
    }
    // Protection belongs to the sheet that has it, as in Excel: that sheet
    // refuses changes and the rest of the workbook stays editable.
    if (elements(document, 'sheetProtection').length) protectedSheets.push(sheet.getAttribute('name'));
    if (scan?.specialFormulas || elements(document, 'f').some((f) => f.hasAttribute('t')))
      unverified.push(
        'Shared, array, or data-table formulas are preserved but unresolved; their cells cannot be overwritten'
      );
    if (scan?.extended || elements(document, 'extLst').length)
      unverified.push('Extended worksheet features are preserved but not interpreted');
    const format = elements(document, 'sheetFormatPr')[0],
      geometryIssues = [...normalGeometryIssues];
    const dimensionDefault = (attribute, maximum, integer = false) => {
      if (!format?.hasAttribute(attribute)) return undefined;
      const value = Number(format.getAttribute(attribute)),
        text = format.getAttribute(attribute).trim();
      if (
        !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) ||
        !Number.isFinite(value) ||
        value < 0 ||
        value > maximum ||
        (integer && !Number.isInteger(value))
      ) {
        geometryIssues.push(`Invalid worksheet ${attribute}; its display default is unresolved`);
        return undefined;
      }
      return value;
    };
    const defaults = {
      defaultRowHeight: dimensionDefault('defaultRowHeight', 409.5),
      defaultColumnWidth: dimensionDefault('defaultColWidth', 255),
      baseColumnWidth: dimensionDefault('baseColWidth', 255, true),
    };
    const hiddenDefault = format?.getAttribute('zeroHeight');
    if (hiddenDefault && !['0', '1', 'false', 'true'].includes(hiddenDefault))
      geometryIssues.push('Invalid default row visibility; its display default is unresolved');
    return {
      name: sheet.getAttribute('name'),
      date1904,
      state: sheet.getAttribute('state') || 'visible',
      path,
      cells,
      rows,
      columns,
      sheetFormat: {
        ...Object.fromEntries(Object.entries(defaults).filter(([, value]) => value !== undefined)),
        rowsHiddenByDefault: ['1', 'true'].includes(hiddenDefault),
      },
      normalFont: { ...normalFont },
      sourceView: sourceViewDefaults(document),
      ...(geometryIssues.length ? { geometryIssues } : {}),
      merges: elements(document, 'mergeCell').map((m) => m.getAttribute('ref')),
      conditionalRules: (() => {
        try {
          return readConditionalRules(document, parts['xl/styles.xml'] ? xml(parts['xl/styles.xml']) : null).map(
            ({ kind, ...rule }, priority) => ({
              ...rule,
              rule: kind,
              threshold: rule.values?.[0],
              priority: priority + 1,
            })
          );
        } catch {
          return [];
        }
      })(),
      validationRules: (() => {
        try {
          return readValidationRules(document);
        } catch {
          return [];
        }
      })(),
      hiddenRows:
        scan?.hiddenRows ??
        sheetRows(document)
          .filter((row) => ['1', 'true'].includes(row.getAttribute('hidden')))
          .map((row) => Number(row.getAttribute('r')) - 1),
      visibleRows:
        scan?.visibleRows ??
        sheetRows(document)
          .filter((row) => ['0', 'false'].includes(row.getAttribute('hidden')))
          .map((row) => Number(row.getAttribute('r')) - 1),
      rowHeights:
        scan?.rowHeights ??
        sheetRows(document)
          .filter((row) => row.hasAttribute('ht'))
          .map((row) => ({ row: Number(row.getAttribute('r')) - 1, height: Number(row.getAttribute('ht')) })),
      freeze: elements(document, 'pane').map((pane) => ({
        columns: Number(pane.getAttribute('xSplit') || 0),
        rows: Number(pane.getAttribute('ySplit') || 0),
        state: pane.getAttribute('state'),
      }))[0],
      widths: elements(document, 'col').map((c) => ({
        min: Number(c.getAttribute('min')),
        max: Number(c.getAttribute('max')),
        width: Number(c.getAttribute('width')),
        hidden: c.getAttribute('hidden') === '1',
      })),
      document,
      scan: scan
        ? { formulaCount: scan.formulaCount, specialFormulas: scan.specialFormulas, extended: scan.extended }
        : undefined,
    };
  });
  const namedRanges = elements(workbook, 'definedName').map((name) => {
    const result = { name: name.getAttribute('name'), formula: name.textContent };
    if (!name.hasAttribute('localSheetId')) return result;
    const raw = name.getAttribute('localSheetId'),
      index = Number(raw);
    return {
      ...result,
      ...(/^[0-9]+$/.test(raw) && Number.isSafeInteger(index) && index <= 4294967295 && sheets[index]
        ? { sheet: sheets[index].name }
        : { scopeInvalid: true }),
    };
  });
  const tables = [];
  for (const sheet of sheets) {
    const relPath = sheet.path.replace(/([^/]+)$/, '_rels/$1.rels');
    if (!parts[relPath]) continue;
    const relations = xml(parts[relPath]);
    for (const relation of Array.from(relations.getElementsByTagName('Relationship'))) {
      if (!relation.getAttribute('Type').endsWith('/table') || relation.getAttribute('TargetMode') === 'External')
        continue;
      const segments = [...sheet.path.split('/').slice(0, -1)];
      const target = relation.getAttribute('Target');
      if (target.startsWith('/')) segments.length = 0;
      for (const segment of target.split('/')) {
        if (segment === '..') {
          if (!segments.length) throw refusal('package-malformed', 'Invalid table relationship');
          segments.pop();
        } else if (segment && segment !== '.') segments.push(segment);
      }
      const tablePath = segments.join('/');
      if (!parts[tablePath]) throw refusal('package-refused', 'Workbook contains a missing table');
      const document = xml(parts[tablePath]),
        table = document.documentElement;
      tables.push({
        name: table.getAttribute('displayName') || table.getAttribute('name'),
        sheet: sheet.name,
        ref: table.getAttribute('ref'),
        path: tablePath,
        headerRowCount: Number(table.getAttribute('headerRowCount') || 1),
        totalsRowCount: Number(table.getAttribute('totalsRowCount') || 0),
        columns: elements(document, 'tableColumn').map((column) => ({ name: column.getAttribute('name') })),
      });
    }
  }
  for (const sheet of sheets)
    sheet.tableRanges = tables
      .filter((table) => table.sheet === sheet.name)
      .map((table) => ({ name: table.name, ref: table.ref }));
  for (const sheet of sheets) {
    const filters = [
      ...children(sheet.document.documentElement, 'autoFilter'),
      ...tables
        .filter((t) => t.sheet === sheet.name)
        .flatMap((t) => children(xml(parts[t.path]).documentElement, 'autoFilter')),
    ];
    const inspected = filters.map((filter) => inspectFilter(filter)),
      owned = inspected.filter((f) => f.record);
    const selected =
      owned.length === 1
        ? owned[0]
        : (inspected.find((f) => f.definition?.reason !== 'No active owned filter criteria') ?? inspected[0]);
    if (owned.length > 1) {
      sheet.filter = {
        range: '',
        owned: false,
        revision: null,
        canReapply: false,
        canClear: false,
        reason: 'Multiple owned filters on one worksheet are protected',
      };
    } else if (selected) {
      sheet.filter = selected.definition;
      sheet.filterRecord = selected.record;
    } else if (
      children(sheet.document.documentElement, 'sheetPr').some((n) =>
        ['1', 'true'].includes(n.getAttribute('filterMode'))
      )
    ) {
      sheet.filter = {
        range: '',
        owned: false,
        revision: null,
        canReapply: false,
        canClear: false,
        reason: 'Imported filtering visibility has no owned provenance',
      };
    }
  }
  const names = Object.keys(parts);
  const facts = packageFacts(parts);
  for (const sheet of sheets) sheet.facts = sheetFacts(sheet, parts);
  return {
    workbook,
    pageTheme,
    sheets,
    styles,
    cellStyles,
    strings,
    namedRanges,
    tables,
    features: names.filter((name) => /charts\/|media\/|pivot|embeddings\/|customXml\//.test(name)),
    package: facts,
    linkedRowReferences: facts.linkedRowReferences,
    structuralReason: structuralReasonOf(sheets, namedRanges.length),
    reasons: [...new Set(reasons)],
    unverified: [...new Set(unverified)],
    protectedSheets,
    macroProject: hasMacroProject(parts),
  };
}
// What the package says about linked data, whatever produced the answer: the
// engine (indexed workbooks) or a read of the parts (the in-memory path).
function packageFacts(parts) {
  const names = Object.keys(parts);
  return {
    linkedRowReferences: names.some((name) => /externalLinks|pivotTables|connections\.xml|slicer/.test(name)),
    linkedData: names.some((name) => /connections\.xml|slicer|queryTables\//.test(name)),
    pivotTables: names.some((name) => /pivotTables\//.test(name)),
    pivotConsolidation: () => pivotJoinsRanges(parts),
  };
}
/** The few facts about one sheet that decide whether rows and columns can move. */
const COUNTED = [
  'f',
  'mergeCell',
  'tableParts',
  'drawing',
  'conditionalFormatting',
  'dataValidations',
  'hyperlink',
  'legacyDrawingHF',
  'oleObjects',
  'controls',
  'autoFilter',
  'rowBreaks',
  'colBreaks',
  'legacyDrawing',
];
function sheetFacts(sheet, parts) {
  const { document, scan } = sheet,
    counts = {},
    extensions = elements(document, 'extLst');
  for (const name of COUNTED) {
    const count = elements(document, name).length;
    if (count) counts[name] = count;
  }
  const slash = sheet.path.lastIndexOf('/'),
    links = parts[`${sheet.path.slice(0, slash)}/_rels/${sheet.path.slice(slash + 1)}.rels`];
  return {
    elements: counts,
    formulaCount: scan?.formulaCount ?? 0,
    scanSpecial: !!scan?.specialFormulas,
    scanExtended: !!scan?.extended,
    specialFormulas: elements(document, 'f').filter((formula) => formula.hasAttribute('t')).length,
    extLstCount: extensions.length,
    extLstFilled: extensions.some((extension) => extension.childNodes.length > 0),
    pivotLink: !!links && /\/pivotTable"/.test(strFromU8(links)),
  };
}
function structuralReasonOf(sheets, definedNames) {
  return sheets.some((sheet) => COUNTED.some((name) => sheet.facts.elements[name]) || sheet.facts.formulaCount) ||
    definedNames
    ? 'Row/column changes are unavailable for workbooks with formulas, merges, tables, drawings, validation, or named ranges. Cell editing remains available.'
    : undefined;
}
/** Known primary worksheet display flags only. Unknown/multiple views remain
 * preserved in XML and do not supply default authority for the local renderer. */
function sourceViewDefaults(document) {
  const containers = children(document.documentElement, 'sheetViews');
  if (containers.length !== 1) return {};
  const views = children(containers[0], 'sheetView').filter((view) => view.getAttribute('workbookViewId') === '0');
  if (views.length !== 1) return {};
  const view = views[0],
    result = {};
  if (view.hasAttribute('zoomScale')) {
    const raw = view.getAttribute('zoomScale');
    if (/^\d+$/.test(raw) && Number(raw) >= 10 && Number(raw) <= 400) result.zoom = Number(raw);
  }
  for (const [attribute, key] of [
    ['showGridLines', 'gridlines'],
    ['showRowColHeaders', 'headings'],
  ]) {
    const raw = view.getAttribute(attribute);
    if (['0', '1', 'false', 'true'].includes(raw)) result[key] = raw === '1' || raw === 'true';
  }
  return result;
}

export function readWorkbook(bytes) {
  return viewWorkbook(inspect(unpack(bytes), undefined, { stream: true }), hash(bytes));
}

/** Metadata packages supplied by the Rust index omit worksheet cell bodies and
 * shared strings. Resolve styles once for decorating requested ranges only. */
export function readWorkbookMetadata(bytes) {
  const state = inspect(unpack(bytes), undefined, { stream: true });
  return { view: viewWorkbook(state, hash(bytes)), styles: state.cellStyles };
}

/** The engine's answer about a cell-free package (`workbook-view`) as the state
 * that `viewWorkbook` and `rowInsertionReason` work from. Whether the workbook is
 * read-only, and what rows may do, are still decided here, from the answer. */
function stateFromEngine(answer) {
  const sheets = answer.sheets.map(({ elements, specialFormulas, extLstCount, extLstFilled, pivotLink, ...sheet }) => ({
    cells: {},
    rows: 20,
    columns: 8,
    hiddenRows: [],
    visibleRows: [],
    rowHeights: [],
    ...sheet,
    facts: {
      elements,
      formulaCount: 0,
      scanSpecial: false,
      scanExtended: extLstCount > 0,
      specialFormulas,
      extLstCount,
      extLstFilled,
      pivotLink,
    },
  }));
  const { pivotConsolidation, ...rest } = answer.package;
  return {
    sheets,
    cellStyles: answer.cellStyles,
    strings: [],
    namedRanges: answer.namedRanges,
    tables: answer.tables,
    features: answer.features,
    pageTheme: answer.pageTheme,
    package: { ...rest, pivotConsolidation: () => pivotConsolidation },
    linkedRowReferences: rest.linkedRowReferences,
    structuralReason: structuralReasonOf(sheets, answer.namedRanges.length),
    // The read-only decision for the package itself.
    reasons: answer.reasons,
    macroProject: answer.macroProject,
    unverified: answer.unverified,
    protectedSheets: answer.protectedSheets,
  };
}

/** The workbook view and the per-style display table for a cell-free package,
 * from the engine's `workbook-view` answer. */
export function viewFromEngine(answer, revision) {
  const state = stateFromEngine(answer);
  return { view: viewWorkbook(state, revision), styles: state.cellStyles };
}

/** Whether rows and columns can be inserted or deleted, by sheet name (null when
 * they can), from the engine's `workbook-view` answer. */
export function structuralReasonsFromEngine(answer, options = {}) {
  const state = stateFromEngine(answer);
  return new Map(
    state.sheets.map((sheet) => [
      sheet.name,
      rowInsertionReason(state, sheet, options.objects === true, options.scoped === true),
    ])
  );
}

/** Persist only results calculated for this projection, never stale input caches. */
export function persistCalculation(bytes, view, context) {
  if (context && context.bytes !== bytes)
    throw refusal('package-unsupported', 'Calculation package context belongs to different bytes');
  const formulaSheets = view.sheets.filter((sheet) => {
    for (const ref in sheet.cells) if (sheet.cells[ref].formula || sheet.cells[ref].formulaType) return true;
    return false;
  });
  const parts = formulaSheets.length ? (context?.parts ?? unpack(bytes)) : unpack(bytes);
  let incomplete = false;
  for (const sheet of formulaSheets) {
    const selection = readWorksheetXml(parts[sheet.path], () => null, {
      includeCell: () => false,
      includeFormulas: true,
      capture: true,
    });
    const doc = selection.document;
    let changed = false;
    for (const cell of sheetCells(doc)) {
      if (!children(cell, 'f').length) continue;
      const result = sheet.cells[cell.getAttribute('r')]?.calculation;
      const desiredType =
        result?.status === 'error'
          ? 'e'
          : result?.status === 'calculated'
            ? typeof result.value === 'boolean'
              ? 'b'
              : typeof result.value === 'string'
                ? 'str'
                : ''
            : '';
      const desiredValue =
        result?.status === 'error'
          ? result.error || '#VALUE!'
          : result?.status === 'calculated'
            ? typeof result.value === 'string'
              ? encodeWorkbookString(result.value)
              : String(typeof result.value === 'boolean' ? Number(result.value) : (result.value ?? ''))
            : null;
      if (desiredValue === null) incomplete = true;
      const values = children(cell, 'v');
      if (
        (cell.getAttribute('t') || '') === desiredType &&
        (desiredValue === null ? values.length === 0 : values.length === 1 && values[0].textContent === desiredValue)
      )
        continue;
      for (const old of children(cell, 'v')) cell.removeChild(old);
      cell.removeAttribute('t');
      changed = true;
      if (result?.status === 'calculated') {
        if (typeof result.value === 'boolean') cell.setAttribute('t', 'b');
        else if (typeof result.value === 'string') cell.setAttribute('t', 'str');
        appendCellData(cell, node(doc, 'v', {}, desiredValue));
      } else if (result?.status === 'error') {
        cell.setAttribute('t', 'e');
        appendCellData(cell, node(doc, 'v', {}, result.error || '#VALUE!'));
      } else incomplete = true;
    }
    if (changed) parts[sheet.path] = mergeSelectedWorksheet(selection, doc);
  }
  const workbook = context?.workbook ?? xml(parts['xl/workbook.xml']);
  let calc = elements(workbook, 'calcPr')[0],
    changed = !calc;
  if (!calc) {
    calc = node(workbook, 'calcPr');
    workbook.documentElement.appendChild(calc);
  }
  if (
    calc.getAttribute('fullCalcOnLoad') !== (incomplete ? '1' : '0') ||
    calc.getAttribute('forceFullCalc') !== (incomplete ? '1' : '0') ||
    calc.getAttribute('calcMode') !== 'auto'
  )
    changed = true;
  calc.setAttribute('fullCalcOnLoad', incomplete ? '1' : '0');
  calc.setAttribute('forceFullCalc', incomplete ? '1' : '0');
  calc.setAttribute('calcMode', 'auto');
  if (changed) parts['xl/workbook.xml'] = strToU8(serializeXml(workbook));
  return writeZipParts(parts);
}
function viewWorkbook(state, revision) {
  const view = {
    revision,
    readOnly: state.reasons.length > 0,
    limitations: [...state.reasons, ...state.unverified],
    protectedSheets: state.protectedSheets,
    macroProject: state.macroProject,
    structuralReason: state.structuralReason,
    calculation:
      'Stored formula caches are unverified. Calculation evidence is supplied separately for the exact accepted revision.',
    namedRanges: state.namedRanges,
    tables: state.tables,
    features: state.features,
    pageTheme: state.pageTheme,
    sheets: state.sheets.map(({ document, scan, filterRecord, facts, ...sheet }) => ({
      ...sheet,
      capabilities: {
        insertRows: !rowInsertionReason(state, { ...sheet, facts }),
        insertRowsReason: rowInsertionReason(state, { ...sheet, facts }),
      },
    })),
  };
  viewContexts.set(view, { strings: state.strings });
  return view;
}
export function editWorkbook(bytes, edits, expectedRevision) {
  return editWorkbookResult(bytes, edits, expectedRevision, false).bytes;
}
function packageResult(parts, state, bytes, includeView, retainContext) {
  // Keep decoded members and small workbook metadata operation-scoped, but
  // release worksheet DOMs before the calculation engine is allocated. They
  // never enter an accepted view or receipt.
  return {
    bytes,
    view: includeView ? viewWorkbook(inspect(parts, state, { sharedStrings: state.strings }), hash(bytes)) : undefined,
    ...(retainContext ? { context: { bytes, parts, workbook: state.workbook } } : {}),
  };
}

function applyXmlPatches(source, patches) {
  let cursor = 0;
  const chunks = [];
  for (const patch of patches.sort((a, b) => a.start - b.start || a.end - b.end)) {
    if (patch.start < cursor) throw refusal('package-conflict', 'Overlapping worksheet edits');
    chunks.push(source.slice(cursor, patch.start), patch.text);
    cursor = patch.end;
  }
  chunks.push(source.slice(cursor));
  return chunks.join('');
}

function mergeSelectedWorksheet(selection, document) {
  const { source } = selection,
    patches = [],
    addedByRow = new Map();
  let added = false;
  for (const cell of sheetCells(document)) {
    const ref = cell.getAttribute('r'),
      location = selection.cellLocations.get(ref),
      text = serializeXml(cell);
    if (location) patches.push({ ...location, text });
    else {
      added = true;
      const position = cellPosition(ref),
        row = position.row + 1;
      let cells = addedByRow.get(row);
      if (!cells) addedByRow.set(row, (cells = []));
      cells.push({ column: position.col, text });
    }
  }
  for (const [row, cells] of addedByRow) {
    cells.sort((a, b) => a.column - b.column);
    const location = selection.rowLocations.get(row);
    if (location) {
      const original = source.slice(location.start, location.end);
      if (/\/\s*>$/.test(original))
        patches.push({
          ...location,
          text: original.replace(/\/\s*>$/, '>') + cells.map((cell) => cell.text).join('') + `</${location.name}>`,
        });
      else
        for (const cell of cells) {
          const next = selection.rowCells.get(row)?.find((item) => cellPosition(item.ref).col > cell.column);
          const start = next?.start ?? source.lastIndexOf('</', location.end - 1);
          patches.push({ start, end: start, text: cell.text });
        }
    } else {
      const created = sheetRows(document).find((item) => Number(item.getAttribute('r')) === row);
      if (!created) throw refusal('package-bad-input', 'Projected worksheet cell has no containing row');
      let start;
      for (let index = 0; index < selection.rowStarts.length; index += 2)
        if (selection.rowStarts[index] > row) {
          start = selection.rowStarts[index + 1];
          break;
        }
      if (start === undefined && selection.sheetDataStart !== undefined) {
        const end = xmlElementEnd(source, selection.sheetDataStart);
        if (/\/\s*>$/.test(source.slice(selection.sheetDataStart, end))) {
          // All new rows are materialized together for an initially empty sheet.
          const data = sheetData(document);
          patches.push({ start: selection.sheetDataStart, end, text: serializeXml(data) });
          break;
        }
        start = source.lastIndexOf('</', end - 1);
      }
      if (start === undefined) {
        const end = xmlElementEnd(source, selection.rootStart);
        start = source.lastIndexOf('</', end - 1);
        patches.push({ start, end: start, text: serializeXml(sheetData(document)) });
        break;
      }
      patches.push({ start, end: start, text: serializeXml(created) });
    }
  }
  if (added && selection.dimension) patches.push({ ...selection.dimension, text: '' });
  return patches.length ? strToU8(applyXmlPatches(source, patches)) : undefined;
}

/** Ordinary edits retain full source XML and build DOMs only for addressed cells
 * and formulas. All unrelated cell records remain shared immutable input data. */
function editWorkbookCells(bytes, edits, includeView, retainContext, before) {
  const parts = unpack(bytes),
    projectedParts = unpack(bytes),
    selections = new Map();
  const targets = new Map(before.sheets.map((sheet) => [sheet.name, new Set()]));
  for (const edit of edits) {
    const refs = targets.get(edit.sheet);
    if (!refs) throw refusal('package-not-found', `Sheet not found: ${edit.sheet}`);
    refs.add(String(edit.cell).replaceAll('$', '').toUpperCase());
  }
  for (const sheet of before.sheets) {
    const refs = targets.get(sheet.name),
      targetRows = new Set([...refs].map((ref) => cellPosition(ref).row + 1));
    const selection = readWorksheetXml(parts[sheet.path], () => null, {
      includeCell: (ref) => refs.has(ref),
      includeFormulas: true,
      capture: true,
      targetRows,
    });
    selections.set(sheet.path, selection);
    projectedParts[sheet.path] = strToU8(serializeXml(selection.document));
  }
  const result = editWorkbookResult(writeZipParts(projectedParts), edits, undefined, true, undefined, false, {
    dom: true,
    sharedStrings: viewContexts.get(before)?.strings,
  });
  const changedParts = unpack(result.bytes);
  const worksheetPaths = new Set(before.sheets.map((sheet) => sheet.path));
  const changedNames = new Set(Object.keys(changedParts));
  for (const name of Object.keys(parts)) if (!changedNames.has(name)) delete parts[name];
  for (const name of Object.keys(changedParts)) {
    if (worksheetPaths.has(name)) continue;
    if (!sameZipMember(parts, changedParts, name)) parts[name] = changedParts[name];
  }
  for (const sheet of before.sheets) {
    const changed = mergeSelectedWorksheet(selections.get(sheet.path), xml(changedParts[sheet.path]));
    if (changed) parts[sheet.path] = changed;
  }
  const output = writeZipParts(parts),
    revision = hash(output);
  const view = includeView
    ? {
        ...before,
        revision,
        structuralReason: result.view.structuralReason,
        sheets: before.sheets.map((sheet) => {
          const changed = result.view.sheets.find((item) => item.name === sheet.name);
          return {
            ...sheet,
            rows: Math.max(sheet.rows, changed.rows),
            columns: Math.max(sheet.columns, changed.columns),
            cells: { ...sheet.cells, ...changed.cells },
          };
        }),
      }
    : undefined;
  if (view) viewContexts.set(view, viewContexts.get(before) ?? viewContexts.get(result.view));
  return {
    bytes: output,
    view,
    changedCells: Object.fromEntries(result.view.sheets.map((sheet) => [sheet.name, Object.keys(sheet.cells)])),
    ...(retainContext ? { context: { bytes: output, parts, workbook: xml(parts['xl/workbook.xml']) } } : {}),
  };
}
/** Rich runs and value annotations cannot be reconstructed from the plain grid value. */
function assertPreservedTextEdit(cell, parts, state) {
  const safeAttributes = (element) =>
    Array.from(element.attributes).every(
      (attribute) => attribute.name === 'xml:space' || attribute.name === 'xmlns' || attribute.name.startsWith('xmlns:')
    );
  const plain = (element) =>
    safeAttributes(element) &&
    Array.from(element.childNodes).every(
      (child) => child.nodeType === 1 || ([3, 4].includes(child.nodeType) && !child.data.trim())
    ) &&
    children(element).every(
      (child) =>
        child.namespaceURI === NS &&
        child.localName === 't' &&
        safeAttributes(child) &&
        Array.from(child.childNodes).every((node) => [3, 4].includes(node.nodeType))
    );
  if (
    children(cell, 'v').some(
      (value) => !safeAttributes(value) || Array.from(value.childNodes).some((node) => ![3, 4].includes(node.nodeType))
    ) ||
    children(cell, 'f').some(
      (formula) =>
        Array.from(formula.childNodes).some((node) => ![3, 4].includes(node.nodeType)) ||
        Array.from(formula.attributes).some(
          (attribute) =>
            !(
              formula.getAttribute('t') === 'dataTable'
                ? ['t', 'ref', 'r1', 'r2', 'dt2D', 'dtr', 'del1', 'del2', 'xml:space']
                : ['t', 'si', 'ref', 'xml:space']
            ).includes(attribute.name) &&
            attribute.name !== 'xmlns' &&
            !attribute.name.startsWith('xmlns:')
        )
    )
  )
    throw refusal(
      'package-unsupported',
      'Opaque value or formula metadata is preserved; content editing requires preservation-aware support'
    );
  let field = children(cell, 'is')[0];
  if (cell.getAttribute('t') === 's') {
    if (!state.preservedTextFields) {
      const source = parts['xl/sharedStrings.xml'];
      if (!source) throw refusal('internal-missing-field', 'Shared string preservation metadata is missing');
      state.preservedTextFields = elements(xml(source), 'si').map(plain);
    }
    const index = Number(children(cell, 'v')[0]?.textContent);
    if (!Number.isSafeInteger(index) || index < 0 || state.preservedTextFields[index] !== true)
      throw refusal(
        'package-unsupported',
        'Rich text or annotated/opaque value metadata is preserved; content editing requires preservation-aware support'
      );
  } else if (field && !plain(field))
    throw refusal(
      'package-unsupported',
      'Rich text or annotated/opaque value metadata is preserved; content editing requires preservation-aware support'
    );
}

export function editWorkbookResult(
  bytes,
  edits,
  expectedRevision,
  includeView = true,
  structuralView,
  retainContext = false,
  options = {}
) {
  if (expectedRevision && hash(bytes) !== expectedRevision)
    throw refusal('package-stale', 'Conflict: this workbook changed. Reload or save a copy.');
  if (!Array.isArray(edits)) throw refusal('internal-invalid-field', 'Invalid workbook edit batch');
  if (!options.dom && !structuralView && edits.every((edit) => !edit.operation))
    return editWorkbookCells(bytes, edits, includeView, retainContext, options.sourceView ?? readWorkbook(bytes));
  const parts = unpack(bytes);
  const state = inspect(parts, undefined, { sharedStrings: options.sharedStrings });
  if (state.reasons.length) throw refusal('package-protected', `Read-only workbook: ${state.reasons.join('; ')}`);
  if (!edits.length)
    return {
      bytes,
      view: includeView ? viewWorkbook(state, hash(bytes)) : undefined,
      ...(retainContext ? { context: { bytes, parts, workbook: state.workbook } } : {}),
    };
  edits = resolveWorkbookEdits(
    { ...viewWorkbook(state, hash(bytes)), ...(structuralView ? { structuralReason: undefined } : {}) },
    edits
  );
  const sheets = new Map(state.sheets.map((sheet) => [sheet.name, sheet]));
  const indexes = new Map();
  const changedSheets = new Set(),
    changedDimensions = new Set();
  let stylesChanged = false,
    workbookChanged = false;
  for (const edit of edits) {
    if (['addSheet', 'renameSheet', 'duplicateSheet'].includes(edit.operation)) {
      const source = sheets.get(edit.sheet);
      if (edit.operation === 'renameSheet') {
        // Only parts that name a sheet matter here. A picture, an embedded
        // file or a link to another workbook does not; a pivot table's source
        // and a chart's series do, and the indexed session rewrites both.
        const partNames = Object.keys(parts),
          namesSheet =
            partNames.some((name) => /connections\.xml|slicer|queryTables\//.test(name)) ||
            (!options.nativeRename && pivotReadsSheet(parts, edit.sheet)) ||
            pivotJoinsRanges(parts);
        if (
          (options.nativeRename ? namesSheet : state.features.length || state.linkedRowReferences) ||
          (!options.nativeRename &&
            (state.namedRanges.length ||
              state.tables.length ||
              state.sheets.some((item) => sheetCells(item.document).some((cell) => children(cell, 'f').length))))
        )
          throw refusal(
            'package-unsupported',
            'Renaming sheets with formulas, names, tables or linked features is not qualified'
          );
        const entry = elements(state.workbook, 'sheet').find((item) => item.getAttribute('name') === edit.sheet);
        if (!entry) throw refusal('internal-missing-field', 'Worksheet entry is missing');
        entry.setAttribute('name', edit.newName);
        workbookChanged = true;
        continue;
      }
      if (edit.operation === 'duplicateSheet') {
        if (!source) throw refusal('internal-missing-field', 'Source worksheet is missing');
        // A copy of a sheet carries its rules, dropdowns, links and notes with
        // it: their ranges are on the sheet itself, so they are right as they
        // stand on the copy. Its tables are copied under new names by the
        // indexed session. Extension records are copied when they are the
        // sheet's own (extended rules and dropdowns) and do not name the sheet.
        const own = [
          '{78C0D931-6437-407d-A8EE-F0AAD7539E65}', // extended highlight rules
          '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}', // extended dropdowns
        ].map((uri) => uri.toLowerCase());
        const quoted = `'${edit.sheet.replaceAll("'", "''")}'!`;
        const foreignExtension = elements(source.document, 'extLst').some(
          (list) =>
            list.parentNode === source.document.documentElement &&
            children(list).some((extension) => {
              const text = serializeXml(extension);
              return (
                !own.includes((extension.getAttribute('uri') ?? '').toLowerCase()) ||
                text.includes(quoted) ||
                text.includes(edit.sheet + '!')
              );
            })
        );
        const unsupported = options.nativeDuplicate
          ? ['f']
          : [
              'f',
              'tableParts',
              'extLst',
              'drawing',
              'mergeCell',
              'conditionalFormatting',
              'dataValidations',
              'hyperlink',
              'legacyDrawing',
            ];
        if (
          foreignExtension ||
          unsupported.some((name) => elements(source.document, name).length) ||
          (!options.nativeDuplicate && parts[source.path.replace(/([^/]+)$/, '_rels/$1.rels')])
        )
          throw refusal('package-refused', 'Duplicating this worksheet would lose linked or advanced features');
      }
      const sheetEntries = elements(state.workbook, 'sheet');
      const template = sheetEntries[0];
      if (!template) throw refusal('package-bad-input', 'Workbook has no worksheet template');
      let partNumber = 1;
      while (parts[`xl/worksheets/sheet${partNumber}.xml`]) partNumber++;
      const partPath = `xl/worksheets/sheet${partNumber}.xml`;
      const rels = xml(parts['xl/_rels/workbook.xml.rels']);
      const relationEntries = Array.from(rels.getElementsByTagName('Relationship'));
      const templateRelation = relationEntries.find(
        (item) => item.getAttribute('Id') === template.getAttribute('r:id')
      );
      if (!templateRelation) throw refusal('internal-missing-field', 'Worksheet relationship template is missing');
      let relNumber = 1;
      while (relationEntries.some((item) => item.getAttribute('Id') === `rId${relNumber}`)) relNumber++;
      const relId = `rId${relNumber}`;
      const relation = templateRelation.cloneNode(false);
      relation.setAttribute('Id', relId);
      relation.setAttribute('Target', `worksheets/sheet${partNumber}.xml`);
      rels.documentElement.appendChild(relation);
      parts['xl/_rels/workbook.xml.rels'] = strToU8(serializeXml(rels));
      const entry = template.cloneNode(false);
      entry.setAttribute('name', edit.operation === 'addSheet' ? edit.sheet : edit.newName);
      entry.setAttribute(
        'sheetId',
        String(Math.max(0, ...sheetEntries.map((item) => Number(item.getAttribute('sheetId')) || 0)) + 1)
      );
      entry.setAttribute('r:id', relId);
      entry.removeAttribute('state');
      template.parentNode.appendChild(entry);
      const types = xml(parts['[Content_Types].xml']);
      const worksheetType = Array.from(types.getElementsByTagName('Override')).find(
        (item) => item.getAttribute('PartName') === `/${source?.path ?? state.sheets[0].path}`
      );
      if (!worksheetType) throw refusal('internal-missing-field', 'Worksheet content type template is missing');
      const override = worksheetType.cloneNode(false);
      override.setAttribute('PartName', `/${partPath}`);
      types.documentElement.appendChild(override);
      parts['[Content_Types].xml'] = strToU8(serializeXml(types));
      parts[partPath] =
        edit.operation === 'duplicateSheet'
          ? parts[source.path].slice()
          : strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS}"><sheetData/></worksheet>`);
      workbookChanged = true;
      continue;
    }
    const sheet = sheets.get(edit.sheet);
    if (!sheet) throw refusal('package-not-found', `Sheet not found: ${edit.sheet}`);
    const doc = sheet.document;
    if (edit.operation === 'createTable') {
      const [firstRef, lastRef] = edit.range.split(':'),
        first = cellPosition(firstRef),
        last = cellPosition(lastRef);
      if (state.tables.some((table) => table.name.toLowerCase() === edit.tableName.toLowerCase()))
        throw refusal('package-duplicate', 'Table name already exists');
      const intersects = (range) => {
        const [a, b = a] = range.split(':'),
          start = cellPosition(a),
          end = cellPosition(b);
        return first.row <= end.row && start.row <= last.row && first.col <= end.col && start.col <= last.col;
      };
      if (
        state.tables.some((table) => table.sheet === sheet.name && intersects(table.ref)) ||
        elements(doc, 'mergeCell').some((item) => intersects(item.getAttribute('ref')))
      )
        throw refusal('package-conflict', 'Table overlaps a table or merged range');
      const covered = sheetCells(doc).filter((cell) => intersects(cell.getAttribute('r')));
      if (covered.some((cell) => children(cell, 'f').length))
        throw refusal('package-unsupported', 'Table creation over formulas is not qualified');
      const sourceSheet = viewWorkbook(state, hash(bytes)).sheets.find((item) => item.name === sheet.name);
      const headers = [],
        seen = new Set();
      for (let col = first.col; col <= last.col; col++) {
        const value = sourceSheet.cells[`${columnName(col)}${first.row + 1}`]?.value;
        if (typeof value !== 'string' || !value.trim() || seen.has(value.toLowerCase()))
          throw refusal('package-bad-input', 'Table headers must be unique nonempty text');
        seen.add(value.toLowerCase());
        headers.push(value);
      }
      let tableNumber = 1;
      while (parts[`xl/tables/table${tableNumber}.xml`]) tableNumber++;
      const tablePath = `xl/tables/table${tableNumber}.xml`;
      const tableId =
        Math.max(
          0,
          ...state.tables.map((table) => Number(xml(parts[table.path]).documentElement.getAttribute('id')) || 0)
        ) + 1;
      const tableDoc = xml(strToU8(`<table xmlns="${NS}"/>`));
      const table = tableDoc.documentElement;
      for (const [key, value] of Object.entries({
        id: tableId,
        name: edit.tableName,
        displayName: edit.tableName,
        ref: edit.range,
        headerRowCount: 1,
        totalsRowCount: 0,
      }))
        table.setAttribute(key, String(value));
      table.appendChild(node(tableDoc, 'autoFilter', { ref: edit.range }));
      const columns = node(tableDoc, 'tableColumns', { count: headers.length });
      headers.forEach((name, index) => columns.appendChild(node(tableDoc, 'tableColumn', { id: index + 1, name })));
      table.appendChild(columns);
      table.appendChild(
        node(tableDoc, 'tableStyleInfo', {
          name: 'TableStyleMedium2',
          showFirstColumn: 0,
          showLastColumn: 0,
          showRowStripes: 1,
          showColumnStripes: 0,
        })
      );
      parts[tablePath] = strToU8(serializeXml(tableDoc));
      const relPath = sheet.path.replace(/([^/]+)$/, '_rels/$1.rels');
      const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
      const relations = parts[relPath] ? xml(parts[relPath]) : xml(strToU8(`<Relationships xmlns="${REL_NS}"/>`));
      const relationEntries = Array.from(relations.getElementsByTagName('Relationship'));
      let relationNumber = 1;
      while (relationEntries.some((item) => item.getAttribute('Id') === `rId${relationNumber}`)) relationNumber++;
      const relationId = `rId${relationNumber}`;
      const relation = relations.createElementNS(REL_NS, 'Relationship');
      for (const [key, value] of Object.entries({
        Id: relationId,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
        Target: `../tables/table${tableNumber}.xml`,
      }))
        relation.setAttribute(key, value);
      relations.documentElement.appendChild(relation);
      parts[relPath] = strToU8(serializeXml(relations));
      doc.documentElement.setAttribute(
        'xmlns:r',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
      );
      const tableParts = children(doc.documentElement, 'tableParts')[0] ?? node(doc, 'tableParts', { count: 0 });
      if (!tableParts.parentNode)
        doc.documentElement.insertBefore(tableParts, children(doc.documentElement, 'extLst')[0] ?? null);
      const tablePart = node(doc, 'tablePart');
      tablePart.setAttribute('r:id', relationId);
      tableParts.appendChild(tablePart);
      tableParts.setAttribute('count', String(children(tableParts, 'tablePart').length));
      const types = xml(parts['[Content_Types].xml']);
      const type = types.createElementNS(types.documentElement.namespaceURI, 'Override');
      type.setAttribute('PartName', `/${tablePath}`);
      type.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
      types.documentElement.appendChild(type);
      parts['[Content_Types].xml'] = strToU8(serializeXml(types));
      changedSheets.add(sheet.name);
      continue;
    }
    if (edit.operation === 'addConditionalRule') {
      if (!state.styles) throw refusal('package-bad-input', 'This workbook has no editable style table');
      const dxfId = appendDifferentialFill(state.styles, edit.fillColor);
      const priority =
        Math.max(0, ...elements(doc, 'cfRule').map((rule) => Number(rule.getAttribute('priority')) || 0)) + 1;
      const container = node(doc, 'conditionalFormatting', { sqref: edit.range });
      const attributes =
        edit.rule === 'duplicateValues'
          ? { type: 'duplicateValues', dxfId, priority }
          : { type: 'cellIs', operator: edit.rule, dxfId, priority };
      const rule = node(doc, 'cfRule', attributes);
      if (edit.rule !== 'duplicateValues') rule.appendChild(node(doc, 'formula', {}, edit.threshold));
      container.appendChild(rule);
      const after = ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'drawing', 'tableParts', 'extLst']
        .map((name) => children(doc.documentElement, name)[0])
        .find(Boolean);
      doc.documentElement.insertBefore(container, after ?? null);
      changedSheets.add(sheet.name);
      stylesChanged = true;
      continue;
    }
    if (['mergeCells', 'unmergeCells'].includes(edit.operation)) {
      const unsupported = [
        'tableParts',
        'drawing',
        'conditionalFormatting',
        'dataValidations',
        'hyperlink',
        'legacyDrawing',
        'extLst',
      ];
      const mergeContainer = children(doc.documentElement, 'mergeCells')[0];
      const entries = children(mergeContainer, 'mergeCell');
      const overlap = (a, b) => {
        const [af, al = af] = a.split(':'),
          [bf, bl = bf] = b.split(':');
        const first = cellPosition(af),
          last = cellPosition(al),
          otherFirst = cellPosition(bf),
          otherLast = cellPosition(bl);
        return (
          first.row <= otherLast.row &&
          otherFirst.row <= last.row &&
          first.col <= otherLast.col &&
          otherFirst.col <= last.col
        );
      };
      // Merging empty cells moves nothing. Excel refuses a merge inside a
      // table and allows one beside it; so does this.
      if (state.tables.some((table) => table.sheet === sheet.name && overlap(table.ref, edit.range)))
        throw refusal('table-or-protection', 'Cells inside a table cannot be merged (Excel does not allow it either)');
      if (edit.operation === 'mergeCells') {
        if (entries.some((item) => overlap(item.getAttribute('ref'), edit.range)))
          throw refusal('package-conflict', 'Merge overlaps an existing merged range');
        const [from] = edit.range.split(':'),
          first = cellPosition(from);
        for (const cell of sheetCells(doc)) {
          const ref = cell.getAttribute('r'),
            position = cellPosition(ref);
          if (position.row === first.row && position.col === first.col) continue;
          if (!overlap(ref, edit.range)) continue;
          if (children(cell, 'f').length || children(cell, 'v').length || children(cell, 'is').length)
            throw refusal('package-conflict', `Merge would hide a nonempty cell: ${ref}`);
        }
        const container = mergeContainer ?? node(doc, 'mergeCells', { count: 0 });
        if (!mergeContainer) {
          // mergeCells follows the filter and view records and precedes these.
          const later = [
            'phoneticPr',
            'conditionalFormatting',
            'dataValidations',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'pageSetup',
            'headerFooter',
            'rowBreaks',
            'colBreaks',
            'customProperties',
            'cellWatches',
            'ignoredErrors',
            'smartTags',
            'drawing',
            'legacyDrawing',
            'legacyDrawingHF',
            'drawingHF',
            'picture',
            'oleObjects',
            'controls',
            'webPublishItems',
            'tableParts',
            'extLst',
          ];
          const next = Array.from(doc.documentElement.childNodes).find(
            (child) => child.nodeType === 1 && later.includes(child.localName)
          );
          doc.documentElement.insertBefore(container, next ?? null);
        }
        container.appendChild(node(doc, 'mergeCell', { ref: edit.range }));
        container.setAttribute('count', String(entries.length + 1));
      } else {
        const entry = entries.find((item) => item.getAttribute('ref') === edit.range);
        if (!entry) throw refusal('package-bad-input', 'Select the exact merged range to unmerge');
        entry.parentNode.removeChild(entry);
        if (entries.length === 1) mergeContainer.parentNode.removeChild(mergeContainer);
        else mergeContainer.setAttribute('count', String(entries.length - 1));
      }
      changedSheets.add(sheet.name);
      continue;
    }
    if (
      edit.operation === 'insertRow' ||
      edit.operation === 'deleteRow' ||
      edit.operation === 'insertColumn' ||
      edit.operation === 'deleteColumn'
    ) {
      if (state.structuralReason && !structuralView) throw refusal('package-refused', state.structuralReason);
      changedSheets.add(sheet.name);
      changedDimensions.add(sheet.name);
      const isRow = edit.operation.endsWith('Row'),
        remove = edit.operation.startsWith('delete'),
        index = edit.index,
        count = edit.count ?? 1;
      if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= (isRow ? 1048576 : 16384) ||
        !Number.isSafeInteger(count) ||
        count < 1 ||
        index + count > (isRow ? 1048576 : 16384)
      )
        throw refusal('internal-invalid-field', 'Invalid row or column');
      for (const c of [...sheetCells(doc)]) {
        const p = cellPosition(c.getAttribute('r')),
          coordinate = isRow ? p.row : p.col;
        if (remove && coordinate >= index && coordinate < index + count) {
          c.parentNode.removeChild(c);
          continue;
        }
        if (coordinate >= index) {
          if (isRow) p.row += remove ? -count : count;
          else p.col += remove ? -count : count;
          if (p.row >= 1048576 || p.col >= 16384)
            throw refusal('package-limit', 'Structural edit exceeds the XLSX boundary');
          c.setAttribute('r', `${columnName(p.col)}${p.row + 1}`);
        }
      }
      if (isRow)
        for (const row of [...sheetRows(doc)]) {
          const n = Number(row.getAttribute('r')) - 1;
          if (remove && n >= index && n < index + count) row.parentNode.removeChild(row);
          else if (n >= index) row.setAttribute('r', String(n + 1 + (remove ? -count : count)));
        }
      else
        for (const col of elements(doc, 'col')) {
          let low = Number(col.getAttribute('min')) - 1,
            high = Number(col.getAttribute('max')) - 1;
          if (remove && low >= index && high < index + count) {
            col.parentNode.removeChild(col);
            continue;
          }
          low = low >= index ? (remove ? Math.max(index, low - count) : low + count) : low;
          high = high >= index ? (remove ? Math.max(index - 1, high - count) : high + count) : high;
          // Formatting that runs to the last column (very common) stops at the
          // edge when a column is inserted, as in Excel; a record pushed wholly
          // off the sheet is dropped.
          if (low >= 16384) {
            col.parentNode.removeChild(col);
            continue;
          }
          high = Math.min(high, 16383);
          col.setAttribute('min', String(low + 1));
          col.setAttribute('max', String(high + 1));
        }
      indexes.delete(sheet.name);
      continue;
    }
    const ref = String(edit.cell).toUpperCase(),
      position = cellPosition(ref);
    let index = indexes.get(sheet.name);
    if (!index) {
      index = {
        data: sheetData(doc),
        rows: new Map(sheetRows(doc).map((row) => [Number(row.getAttribute('r')), row])),
        cells: new Map(sheetCells(doc).map((cell) => [cell.getAttribute('r'), cell])),
      };
      indexes.set(sheet.name, index);
    }
    let data = index.data;
    if (!data) {
      data = node(doc, 'sheetData');
      doc.documentElement.appendChild(data);
      index.data = data;
    }
    let row = index.rows.get(position.row + 1);
    if (!row) {
      row = node(doc, 'row', { r: position.row + 1 });
      index.rows.set(position.row + 1, row);
      data.insertBefore(row, children(data, 'row').find((r) => Number(r.getAttribute('r')) > position.row + 1) ?? null);
    }
    let cell = index.cells.get(ref);
    if (!cell) {
      changedDimensions.add(sheet.name);
      cell = node(doc, 'c', { r: ref });
      index.cells.set(ref, cell);
      row.insertBefore(
        cell,
        children(row, 'c').find((c) => cellPosition(c.getAttribute('r')).col > position.col) ?? null
      );
    }
    if (edit.value !== undefined || edit.formula !== undefined) {
      assertPreservedTextEdit(cell, parts, state);
      changedSheets.add(sheet.name);
      if (children(cell, 'f').some((formula) => formula.hasAttribute('t')))
        throw refusal('package-unsupported', 'Shared, array or data-table formula cells cannot be overwritten');
      for (const child of children(cell))
        if (child.namespaceURI === NS && ['f', 'v', 'is'].includes(child.localName)) cell.removeChild(child);
      cell.removeAttribute('t');
      if (edit.formula !== undefined) {
        const formula = String(edit.formula).replace(/^=/, '');
        if (!formula || formula.length > 8192 || /https?:|file:|DDE\(|WEBSERVICE\(/i.test(formula))
          throw refusal('package-limit', 'External or oversized formulas are unsupported');
        appendCellData(cell, node(doc, 'f', {}, formula));
      } else if (edit.valueType === 'error') {
        cell.setAttribute('t', 'e');
        appendCellData(cell, node(doc, 'v', {}, edit.value));
      } else if (typeof edit.value === 'number' && Number.isFinite(edit.value))
        appendCellData(cell, node(doc, 'v', {}, edit.value));
      else if (typeof edit.value === 'boolean') {
        cell.setAttribute('t', 'b');
        appendCellData(cell, node(doc, 'v', {}, edit.value ? '1' : '0'));
      } else if (edit.value !== null) {
        cell.setAttribute('t', 'inlineStr');
        const inline = node(doc, 'is');
        const t = node(doc, 't', {}, encodeWorkbookString(edit.value ?? ''));
        t.setAttribute('xml:space', 'preserve');
        inline.appendChild(t);
        appendCellData(cell, inline);
      }
    }
    if (edit.styleId !== undefined) {
      if (!state.styles || edit.styleId >= children(elements(state.styles, 'cellXfs')[0], 'xf').length)
        throw refusal('internal-invalid-field', 'Invalid source style');
      cell.setAttribute('s', String(edit.styleId));
      changedSheets.add(sheet.name);
    }
    if (
      [
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
      ].some((field) => edit[field] !== undefined)
    ) {
      changedSheets.add(sheet.name);
      stylesChanged = true;
      if (!state.styles) throw refusal('package-bad-input', 'This workbook has no editable style table');
      applyWorkbookStyle(state.styles, cell, edit);
    }
  }
  if (structuralView) {
    for (const sheet of state.sheets) {
      const transformed = structuralView.sheets.find((item) => item.name === sheet.name);
      if (!transformed) throw refusal('internal-mismatch', 'Structural transform lost a worksheet');
      // The shared transformation fills inserted-row formulas and inherits styles.
      // Existing cells keep their XML; only newly introduced cells are materialized.
      const stored = new Set(sheetCells(sheet.document).map((cell) => cell.getAttribute('r')));
      const data = sheetData(sheet.document);
      for (const [ref, cell] of Object.entries(transformed.cells)) {
        if (stored.has(ref)) continue;
        const position = cellPosition(ref);
        let row = children(data, 'row').find((item) => Number(item.getAttribute('r')) === position.row + 1);
        if (!row) {
          row = node(sheet.document, 'row', { r: position.row + 1 });
          data.insertBefore(
            row,
            children(data, 'row').find((item) => Number(item.getAttribute('r')) > position.row + 1) ?? null
          );
        }
        const inserted = node(sheet.document, 'c', { r: ref, ...(cell.styleId ? { s: cell.styleId } : {}) });
        if (cell.formula) inserted.appendChild(node(sheet.document, 'f', {}, cell.formula));
        row.insertBefore(
          inserted,
          children(row, 'c').find((item) => cellPosition(item.getAttribute('r')).col > position.col) ?? null
        );
        changedSheets.add(sheet.name);
      }
      for (const cell of sheetCells(sheet.document)) {
        const expected = transformed.cells[cell.getAttribute('r')];
        const formulas = children(cell, 'f');
        if (!!expected?.formula !== !!formulas.length)
          throw refusal('package-bad-input', 'Structural formula mapping does not match worksheet cells');
        if (formulas.length && formulas[0].textContent !== expected.formula) {
          formulas[0].textContent = expected.formula;
          changedSheets.add(sheet.name);
        }
      }
    }
    const names = elements(state.workbook, 'definedName');
    for (const name of names) {
      const local = name.hasAttribute('localSheetId')
        ? state.sheets[Number(name.getAttribute('localSheetId'))]?.name
        : undefined;
      const transformed = structuralView.namedRanges.find(
        (item) => item.name === name.getAttribute('name') && item.sheet === local
      );
      if (!transformed) throw refusal('internal-mismatch', 'Structural transform lost a defined name');
      if (name.textContent !== transformed.formula) {
        name.textContent = transformed.formula;
        workbookChanged = true;
      }
    }
    for (const table of state.tables) {
      const transformed = structuralView.tables.find((item) => item.name === table.name && item.sheet === table.sheet);
      if (!transformed) throw refusal('internal-mismatch', 'Structural transform lost a table');
      if (transformed.deleted) {
        const owner = state.sheets.find((s) => s.name === table.sheet),
          relPath = relationshipPart(owner.path);
        const tableRels = relationshipPart(table.path);
        const tableDoc = xml(parts[table.path]);
        if (parts[tableRels] || elements(tableDoc, 'extLst').length || elements(tableDoc, 'xmlColumnPr').length)
          throw refusal('package-unsupported', 'Deleting a linked or extended table is not qualified');
        const rels = xml(parts[relPath]),
          container = elements(owner.document, 'tableParts')[0];
        let removed = 0;
        for (const rel of Array.from(rels.getElementsByTagName('Relationship'))) {
          const target = rel.getAttribute('Target'),
            resolved = resolvePackageTarget(owner.path, target);
          if (resolved !== table.path) continue;
          if (!rel.getAttribute('Type').endsWith('/table') || rel.getAttribute('TargetMode') === 'External')
            throw refusal('package-malformed', 'Invalid table relationship');
          const linked = children(container, 'tablePart').filter(
            (n) => n.getAttribute('r:id') === rel.getAttribute('Id')
          );
          if (linked.length !== 1) throw refusal('package-malformed', 'Invalid table part reference');
          container.removeChild(linked[0]);
          rel.parentNode.removeChild(rel);
          removed++;
        }
        if (removed !== 1) throw refusal('internal-missing-field', 'Missing table relationship');
        const count = children(container, 'tablePart').length;
        if (count) container.setAttribute('count', String(count));
        else container.parentNode.removeChild(container);
        const types = xml(parts['[Content_Types].xml']);
        for (const type of Array.from(types.getElementsByTagName('Override')))
          if (type.getAttribute('PartName') === `/${table.path}`) type.parentNode.removeChild(type);
        parts['[Content_Types].xml'] = strToU8(serializeXml(types));
        parts[relPath] = strToU8(serializeXml(rels));
        delete parts[table.path];
        changedSheets.add(owner.name);
        continue;
      }
      if (transformed.ref === table.ref && !transformed.columnEdit && !transformed.emptyBody) continue;
      const doc = xml(parts[table.path]);
      doc.documentElement.setAttribute('ref', transformed.ref);
      if (transformed.emptyBody) doc.documentElement.setAttribute('insertRow', '1');
      for (const filter of elements(doc, 'autoFilter')) filter.setAttribute('ref', transformed.ref);
      if (transformed.columnEdit) {
        const edit = transformed.columnEdit,
          container = elements(doc, 'tableColumns')[0],
          columns = children(container, 'tableColumn');
        let id = Math.max(0, ...columns.map((c) => Number(c.getAttribute('id'))));
        const anchor = columns[edit.index] ?? null;
        for (const column of edit.insert)
          container.insertBefore(node(doc, 'tableColumn', { id: ++id, name: column.name }), anchor);
        for (const column of columns.slice(edit.index, edit.index + edit.remove)) container.removeChild(column);
        container.setAttribute('count', String(children(container, 'tableColumn').length));
        for (const column of elements(doc, 'filterColumn')) {
          const col = Number(column.getAttribute('colId'));
          if (col >= edit.index && col < edit.index + edit.remove) column.parentNode.removeChild(column);
          else if (col >= edit.index) column.setAttribute('colId', String(col + edit.insert.length - edit.remove));
        }
      }
      parts[table.path] = strToU8(serializeXml(doc));
    }
  }
  // Do not display stale calculated results after inputs change. Excel
  // can calculate every formula, while the app explicitly supports a small subset.
  for (const sheet of state.sheets) {
    for (const cell of sheetCells(sheet.document))
      if (children(cell, 'f').length)
        for (const v of children(cell, 'v')) {
          cell.removeChild(v);
          changedSheets.add(sheet.name);
        }
    if (changedDimensions.has(sheet.name))
      for (const dim of elements(sheet.document, 'dimension')) dim.parentNode.removeChild(dim);
    if (changedSheets.has(sheet.name)) parts[sheet.path] = strToU8(serializeXml(sheet.document));
  }
  let calc = elements(state.workbook, 'calcPr')[0];
  if (!calc) {
    calc = node(state.workbook, 'calcPr');
    state.workbook.documentElement.appendChild(calc);
    workbookChanged = true;
  }
  if (
    calc.getAttribute('fullCalcOnLoad') !== '1' ||
    calc.getAttribute('forceFullCalc') !== '1' ||
    calc.getAttribute('calcMode') !== 'auto'
  )
    workbookChanged = true;
  calc.setAttribute('fullCalcOnLoad', '1');
  calc.setAttribute('forceFullCalc', '1');
  calc.setAttribute('calcMode', 'auto');
  // Cached dependency order is invalid after a formula edit. Remove the chain
  // together with its package references; the workbook requests full recalc.
  const relationships = xml(parts['xl/_rels/workbook.xml.rels']);
  const contentTypes = xml(parts['[Content_Types].xml']);
  let chainRemoved = false;
  for (const relation of Array.from(relationships.getElementsByTagName('Relationship'))) {
    if (!relation.getAttribute('Type').endsWith('/calcChain')) continue;
    chainRemoved = true;
    const target = relation.getAttribute('Target');
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    delete parts[path];
    relation.parentNode.removeChild(relation);
    for (const type of Array.from(contentTypes.getElementsByTagName('Override'))) {
      if (type.getAttribute('PartName') === `/${path}`) type.parentNode.removeChild(type);
    }
  }
  if (chainRemoved) {
    parts['xl/_rels/workbook.xml.rels'] = strToU8(serializeXml(relationships));
    parts['[Content_Types].xml'] = strToU8(serializeXml(contentTypes));
  }
  if (workbookChanged) parts['xl/workbook.xml'] = strToU8(serializeXml(state.workbook));
  if (state.styles && stylesChanged) parts['xl/styles.xml'] = strToU8(serializeXml(state.styles));
  const output = writeZipParts(parts);
  return packageResult(parts, state, output, includeView, retainContext);
}

/**
 * Whether rows and columns can be inserted or deleted, for every sheet at once.
 * Returns the reason it cannot, by sheet name (null when it can). The package is
 * read once: a real workbook's styles alone can be several megabytes, and
 * reading it again for each sheet made opening take a minute.
 */
export function structuralEditReasons(bytes, options = {}) {
  const parts = unpack(bytes),
    state = inspect(parts),
    reasons = new Map();
  for (const sheet of state.sheets)
    reasons.set(sheet.name, rowInsertionReason(state, sheet, options.objects === true, options.scoped === true));
  return reasons;
}

export function assertStructuralEdit(bytes, edit, options = {}) {
  const parts = unpack(bytes),
    state = inspect(parts);
  if (!['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'].includes(edit.operation))
    throw refusal('package-unsupported', 'Only qualified row insertion/deletion is available for referenced workbooks');
  const count = edit.count ?? 1,
    limit = edit.operation.endsWith('Column') ? 16384 : 1048576;
  if (!Number.isSafeInteger(count) || count < 1 || edit.index + count > limit)
    throw refusal('internal-invalid-field', 'Invalid structural count');
  const sheet = state.sheets.find((item) => item.name === edit.sheet);
  if (!sheet) throw refusal('package-not-found', `Sheet not found: ${edit.sheet}`);
  if (
    !Number.isInteger(edit.index) ||
    edit.index < 0 ||
    edit.index >= (edit.operation.endsWith('Column') ? 16384 : edit.operation === 'deleteRow' ? 1048576 : 1048575)
  )
    throw refusal('package-malformed', 'Invalid row index');
  const reason = rowInsertionReason(state, sheet, options.objects === true, options.scoped === true);
  if (reason) throw refusal('package-refused', reason);
}

/** Whether any pivot table summarises this sheet. One whose source cannot be read counts. */
function pivotReadsSheet(parts, sheetName) {
  const escaped = sheetName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  return Object.keys(parts)
    .filter((name) => /pivotCache\/[^/]+\.xml$/.test(name))
    .some((name) => {
      const text = strFromU8(parts[name]);
      if (!/<(?:\w+:)?pivotCacheDefinition[\s>]/.test(text)) return false;
      const source = /<(?:\w+:)?worksheetSource\b[^>]*>/.exec(text)?.[0];
      return !source || !/\bsheet="/.test(source) || source.includes(`sheet="${escaped}"`);
    });
}

/**
 * Whether linked data depends on where this sheet's rows and columns are.
 * A link to another workbook does not: its cells live in that workbook. A pivot
 * table does when it sits on this sheet or summarises it; imported data ranges
 * and slicers are bound to cells this writer does not move.
 */
function linkedToSheet(facts, sheet) {
  if (facts.linkedData) return true;
  if (!facts.pivotTables) return false;
  if (!sheet.path) return true;
  if (sheet.facts.pivotLink) return true;
  // A pivot table elsewhere that summarises this sheet keeps its source: the
  // engine moves the remembered range with the rows. One built from several
  // ranges at once is not moved.
  return facts.pivotConsolidation();
}

/** Whether a pivot table is built from several ranges at once; those are not followed. */
function pivotJoinsRanges(parts) {
  return Object.keys(parts)
    .filter((name) => /pivotCache\/[^/]+\.xml$/.test(name))
    .some((name) => /<(?:\w+:)?consolidation[\s>]/.test(strFromU8(parts[name])));
}

function rowInsertionReason(state, sheet, objects = false, scoped = false) {
  if (state.reasons.length) return `Read-only workbook: ${state.reasons.join('; ')}`;
  if (
    (scoped ? linkedToSheet(state.package, sheet) : state.linkedRowReferences) ||
    (!objects && state.features.some((feature) => /comment/i.test(String(feature))))
  )
    return 'Row insertion cannot preserve linked-data or pivot references';
  // The indexed session has already written shared formulas out in full and
  // checked that other sheets' extension records do not mention this sheet.
  if (
    (scoped ? [sheet] : state.sheets).some(
      ({ facts }) =>
        (!scoped && (facts.scanSpecial || facts.scanExtended)) ||
        facts.specialFormulas > 0 ||
        (facts.extLstCount > 0 && (!scoped || facts.extLstFilled))
    )
  )
    return 'Row insertion cannot preserve shared formulas or extended worksheet references';
  const has = (names) => names.some((name) => sheet.facts.elements[name]);
  if (!objects && has(['mergeCell', 'drawing', 'conditionalFormatting', 'dataValidations', 'hyperlink']))
    return 'Row insertion cannot preserve affected worksheet objects in the compatibility writer';
  // With the object rewrite, a filter, cell notes and manual page breaks follow
  // their rows and columns. Header pictures and embedded controls do not yet.
  const unmoved = objects
    ? ['legacyDrawingHF', 'oleObjects', 'controls']
    : ['autoFilter', 'rowBreaks', 'colBreaks', 'legacyDrawing', 'legacyDrawingHF', 'oleObjects', 'controls'];
  if (has(unmoved))
    return objects
      ? 'Rows and columns cannot be inserted or deleted on a sheet with header pictures or embedded controls'
      : 'Row insertion cannot preserve affected merges, drawings, validation, filters or layout references';
  if (state.tables.some((table) => table.totalsRowCount || table.headerRowCount !== 1))
    return 'Row insertion with table totals or nonstandard headers is unsupported';
  return null;
}

/** Persist a qualified display/row operation without regenerating the workbook. */
export function applyProjectedOperation(bytes, edit, projected, retainContext = false) {
  const parts = unpack(bytes),
    state = inspect(parts);
  if (state.reasons.length) throw refusal('package-protected', `Read-only workbook: ${state.reasons.join('; ')}`);
  const sheet = state.sheets.find((item) => item.name === edit.sheet);
  const target = projected.sheets.find((item) => item.name === edit.sheet);
  if (!sheet || !target) throw refusal('package-not-found', 'Sheet not found');
  const doc = sheet.document,
    data = sheetData(doc);
  if (edit.operation === 'sortRows') {
    normalizeSortKeys(edit);
    if (
      [
        'mergeCell',
        'drawing',
        'conditionalFormatting',
        'dataValidations',
        'hyperlink',
        'tableParts',
        'autoFilter',
        'extLst',
      ].some((name) => elements(doc, name).length) ||
      Object.keys(parts).some((name) => /comments|threadedComments|pivotTables|externalLinks/.test(name))
    )
      throw refusal('package-unsupported', 'Sorting cannot preserve the affected workbook references or objects');
    const order = target.rowOrder;
    const bounds = edit.range.split(':').map(cellPosition),
      first = bounds[0].row,
      last = (bounds[1] ?? bounds[0]).row;
    if (
      !Array.isArray(order) ||
      order.length !== last - first + 1 ||
      order.some(
        (item) =>
          !Number.isInteger(item.from) ||
          !Number.isInteger(item.to) ||
          item.from < first ||
          item.from > last ||
          item.to < first ||
          item.to > last
      ) ||
      new Set(order.map((item) => item.from)).size !== order.length ||
      new Set(order.map((item) => item.to)).size !== order.length
    )
      throw refusal('internal-invalid-field', 'Invalid row permutation');
    const rows = new Map(children(data, 'row').map((row) => [Number(row.getAttribute('r')) - 1, row]));
    const cloned = order.map(({ from, to }) => {
      const row = rows.get(from)?.cloneNode(true);
      if (!row) return null;
      row.setAttribute('r', String(to + 1));
      for (const cell of children(row, 'c')) {
        const col = cellPosition(cell.getAttribute('r')).col,
          ref = `${columnName(col)}${to + 1}`;
        cell.setAttribute('r', ref);
        for (const formula of children(cell, 'f')) {
          if (formula.hasAttribute('t') || !target.cells[ref]?.formula)
            throw refusal('package-unsupported', 'Unqualified sorted formula');
          formula.textContent = target.cells[ref].formula;
        }
      }
      return row;
    });
    for (const { from } of order) {
      const row = rows.get(from);
      if (row) data.removeChild(row);
    }
    for (const row of cloned.filter(Boolean))
      data.insertBefore(
        row,
        children(data, 'row').find((existing) => Number(existing.getAttribute('r')) > Number(row.getAttribute('r'))) ??
          null
      );
  } else if (['filterRows', 'reapplyFilter', 'clearFilter', 'removeFilter'].includes(edit.operation)) {
    const prior = sheet.filterRecord ?? null;
    if (sheet.filter && !prior && sheet.filter.reason !== 'No active owned filter criteria')
      throw refusal('package-refused', sheet.filter.reason);
    const record = prepareFilterState(edit, prior, captureFilterBaseline(doc, edit.range ?? prior?.model.range));
    const active = record.model.values.length > 0,
      range = record.model.range,
      values = record.model.values.map((v) => (v && typeof v === 'object' && v.$workbookNumber === '-0' ? -0 : v));
    const [first, last = first] = range.split(':');
    const start = cellPosition(first),
      end = cellPosition(last),
      hidden = new Set(target.hiddenRows);
    const baselineRows = new Map(record.model.baseline.rows);
    const table = state.tables.find((item) => item.sheet === sheet.name && item.ref === range);
    if (state.tables.some((item) => item.sheet === sheet.name) && !table)
      throw refusal('package-bad-input', 'Filter must match the complete table range');
    const filterDoc = table ? xml(parts[table.path]) : doc;
    const existing = elements(filterDoc, 'autoFilter');
    if (existing.length > 1 || existing.some((filter) => filter.getAttribute('ref') !== range))
      throw refusal('package-unsupported', 'Replacing unrelated filters is unsupported');
    for (const old of existing) old.parentNode.removeChild(old);
    if (edit.operation !== 'removeFilter') {
      const filter = node(filterDoc, 'autoFilter', { ref: range });
      const column = node(filterDoc, 'filterColumn', {
        colId: cellPosition(`${record.model.column}1`).col - start.col,
      });
      const values = node(filterDoc, 'filters');
      for (const value of record.model.values.map((v) =>
        v && typeof v === 'object' && v.$workbookNumber === '-0' ? -0 : v
      )) {
        if (value === null || value === '') values.setAttribute('blank', '1');
        else
          values.appendChild(
            node(filterDoc, 'filter', {
              val: typeof value === 'boolean' ? Number(value) : Object.is(value, -0) ? '-0' : value,
            })
          );
      }
      if (active) {
        column.appendChild(values);
        filter.appendChild(column);
      }
      appendFilterRecord(filterDoc, filter, record);
      if (table) filterDoc.documentElement.insertBefore(filter, elements(filterDoc, 'tableColumns')[0] ?? null);
      else {
        const following = children(doc.documentElement).find((item) =>
          [
            'sortState',
            'dataConsolidate',
            'customSheetViews',
            'mergeCells',
            'phoneticPr',
            'conditionalFormatting',
            'dataValidations',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'pageSetup',
            'headerFooter',
            'drawing',
            'tableParts',
            'extLst',
          ].includes(item.localName)
        );
        doc.documentElement.insertBefore(filter, following ?? null);
      }
    }
    for (const row of children(data, 'row')) {
      const index = Number(row.getAttribute('r')) - 1;
      if (index <= start.row || index > end.row) continue;
      if (
        !active &&
        !baselineRows.has(index) &&
        children(row).length === 0 &&
        Array.from(row.attributes).every((a) => ['r', 'hidden'].includes(a.name) || a.name.startsWith('xmlns')) &&
        Array.from(row.childNodes).every((n) => n.nodeType === 3 && !n.data.trim())
      ) {
        data.removeChild(row);
        continue;
      }
      const lexical = baselineRows.get(index) ?? null;
      if (
        active &&
        hidden.has(index) &&
        lexical !== '1' &&
        lexical !== 'true' &&
        !(lexical === null && ['1', 'true'].includes(record.model.baseline.format.zeroHeight))
      )
        row.setAttribute('hidden', '1');
      else if (lexical === null) row.removeAttribute('hidden');
      else row.setAttribute('hidden', lexical);
    }
    if (active) {
      const existingRows = children(data, 'row'),
        present = new Set(existingRows.map((r) => Number(r.getAttribute('r')) - 1));
      let cursor = 0;
      for (const index of [...hidden].sort((a, b) => a - b)) {
        if (index <= start.row || index > end.row || present.has(index)) continue;
        const lexical = baselineRows.get(index) ?? null;
        if (
          lexical === '1' ||
          lexical === 'true' ||
          (lexical === null && ['1', 'true'].includes(record.model.baseline.format.zeroHeight))
        )
          continue;
        const row = node(doc, 'row', { r: index + 1, hidden: '1' });
        while (cursor < existingRows.length && Number(existingRows[cursor].getAttribute('r')) <= index + 1) cursor++;
        data.insertBefore(row, existingRows[cursor] ?? null);
      }
    }
    if (table) parts[table.path] = strToU8(serializeXml(filterDoc));
  } else if (edit.operation === 'unhideAll') {
    if (['autoFilter', 'tableParts', 'sheetProtection'].some((tag) => elements(doc, tag).length))
      throw refusal('package-protected', 'Unhide requires an unprotected worksheet without linked filters/tables');
    const fmt = elements(doc, 'sheetFormatPr')[0],
      rh = Number(fmt?.getAttribute('defaultRowHeight')) > 0 ? Number(fmt.getAttribute('defaultRowHeight')) : 15,
      cw = Number(fmt?.getAttribute('defaultColWidth')) > 0 ? Number(fmt.getAttribute('defaultColWidth')) : 8.43;
    if (fmt) {
      if (['1', 'true'].includes(fmt.getAttribute('zeroHeight'))) fmt.setAttribute('zeroHeight', '0');
      if (fmt.hasAttribute('defaultRowHeight') && Number(fmt.getAttribute('defaultRowHeight')) === 0)
        fmt.setAttribute('defaultRowHeight', '15');
      if (fmt.hasAttribute('defaultColWidth') && Number(fmt.getAttribute('defaultColWidth')) === 0)
        fmt.setAttribute('defaultColWidth', '8.43');
    }
    for (const row of children(data, 'row')) {
      if (['1', 'true'].includes(row.getAttribute('hidden'))) row.setAttribute('hidden', '0');
      if (row.hasAttribute('ht') && Number(row.getAttribute('ht')) === 0) row.setAttribute('ht', String(rh));
    }
    for (const cols of elements(doc, 'cols'))
      for (const col of children(cols, 'col')) {
        if (['1', 'true'].includes(col.getAttribute('hidden'))) col.setAttribute('hidden', '0');
        if (col.hasAttribute('width') && Number(col.getAttribute('width')) === 0) col.setAttribute('width', String(cw));
      }
  } else if (edit.operation === 'resizeColumn') {
    resizeColumnXml(doc, edit);
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
    const coordinate = edit.index + 1;
    let row = children(data, 'row').find((item) => Number(item.getAttribute('r')) === coordinate);
    if (!row) {
      row = node(doc, 'row', { r: coordinate });
      data.insertBefore(row, children(data, 'row').find((item) => Number(item.getAttribute('r')) > coordinate) ?? null);
    }
    row.setAttribute('ht', String(edit.height));
    row.setAttribute('customHeight', '1');
  } else if (edit.operation === 'freeze') {
    freezeXml(doc, edit);
  } else throw refusal('package-unsupported', 'Unsupported projected operation');
  parts[sheet.path] = strToU8(serializeXml(doc));
  const output = writeZipParts(parts);
  return packageResult(parts, state, output, true, retainContext);
}
/** Internal portable-owner input; baseline is deliberately absent from views. */
export function readWorkbookFilterStates(bytes) {
  const state = inspect(unpack(bytes), undefined, { stream: true });
  return Object.fromEntries(state.sheets.filter((s) => s.filterRecord).map((s) => [s.name, s.filterRecord]));
}
export function createWorkbook() {
  return zipSync(
    Object.fromEntries(
      Object.entries({
        '[Content_Types].xml':
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        '_rels/.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/_rels/workbook.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': `<worksheet xmlns="${NS}"><sheetData/></worksheet>`,
        'xl/styles.xml': `<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
      }).map(([name, value]) => [name, strToU8(value)])
    )
  );
}

/** Prepare only small package metadata for the streaming native writer. The
 * input is Rust's cell-free metadata ZIP, never the source worksheet payload. */
export function indexedCalculationMetadata(metadataBytes, incomplete, policy) {
  const parts = unpack(metadataBytes),
    raw = strFromU8(parts['xl/workbook.xml']),
    workbook = xml(parts['xl/workbook.xml'], { locator: {}, normalizeLineEndings: (value) => value });
  const existing = children(workbook.documentElement, 'calcPr');
  if (existing.length > 1) throw refusal('package-ambiguous', 'Ambiguous calculation policy metadata');
  let calc = existing[0];
  const span = (node) => {
    const lines = lineStarts(raw);
    const start = lines[node.lineNumber - 1] + node.columnNumber - 1;
    if (!Number.isSafeInteger(start) || !raw.startsWith('<' + node.tagName, start))
      throw refusal('internal-missing-field', 'Missing calculation metadata source boundary');
    return [start, xmlElementEnd(raw, start)];
  };
  const previous = calc ? span(calc) : null;
  if (policy?.preserveSource) {
    const flags = { fullCalcOnLoad: '1', forceFullCalc: '1' };
    const close = raw.lastIndexOf('</' + workbook.documentElement.tagName + '>');
    let replacement;
    if (calc) {
      const original = raw.slice(previous[0], previous[1]);
      let quote = null,
        end = 0;
      for (; end < original.length; end++) {
        const char = original[end];
        if (quote) {
          if (char === quote) quote = null;
        } else if (char === '"' || char === "'") quote = char;
        else if (char === '>') break;
      }
      if (end === original.length) throw refusal('package-malformed', 'Invalid calculation opening boundary');
      let opening = original.slice(0, end + 1);
      for (const [key, value] of Object.entries(flags)) {
        const pattern = new RegExp('\\s' + key + '\\s*=\\s*(?:"[^"<>]*"|\'[^\'<>]*\')');
        opening = pattern.test(opening)
          ? opening.replace(pattern, ' ' + key + '="' + value + '"')
          : opening.replace(/\/?>(?=$)/, (ending) => ' ' + key + '="' + value + '"' + ending);
      }
      replacement = raw.slice(0, previous[0]) + opening + original.slice(end + 1) + raw.slice(previous[1]);
    } else {
      if (close < 0) throw refusal('internal-missing-field', 'Missing workbook closing boundary');
      const prefix = workbook.documentElement.prefix ? workbook.documentElement.prefix + ':' : '';
      replacement =
        raw.slice(0, close) + '<' + prefix + 'calcPr fullCalcOnLoad="1" forceFullCalc="1"/>' + raw.slice(close);
    }
    return { replacements: { 'xl/workbook.xml': replacement }, removed: [] };
  }
  const desired = {
    fullCalcOnLoad: incomplete ? '1' : '0',
    forceFullCalc: incomplete ? '1' : '0',
    calcMode: policy?.mode ?? 'auto',
    ...(policy ? { calcOnSave: policy.recalculateBeforeSave ? '1' : '0' } : {}),
  };
  const defaults =
    !calc && !incomplete && desired.calcMode === 'auto' && (!policy || policy.recalculateBeforeSave === true);
  const unchanged = calc && Object.entries(desired).every(([key, value]) => calc.getAttribute(key) === value);
  if (!calc) {
    calc = node(workbook, 'calcPr');
    workbook.documentElement.appendChild(calc);
  }
  for (const [key, value] of Object.entries(desired)) calc.setAttribute(key, value);
  const serialized = unchanged ? raw.slice(previous[0], previous[1]) : serializeXml(calc),
    close = raw.lastIndexOf('</' + workbook.documentElement.tagName + '>');
  if (!previous && close < 0) throw refusal('internal-missing-field', 'Missing workbook closing boundary');
  const replacements = {
      'xl/workbook.xml': defaults
        ? raw
        : previous
          ? raw.slice(0, previous[0]) + serialized + raw.slice(previous[1])
          : raw.slice(0, close) + serialized + raw.slice(close),
    },
    removed = [];
  const relationships = xml(parts['xl/_rels/workbook.xml.rels']),
    contentTypes = xml(parts['[Content_Types].xml']);
  for (const rel of Array.from(relationships.getElementsByTagName('Relationship'))) {
    if (!rel.getAttribute('Type').endsWith('/calcChain')) continue;
    const target = rel.getAttribute('Target'),
      name = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    removed.push(name);
    rel.parentNode.removeChild(rel);
    for (const type of Array.from(contentTypes.getElementsByTagName('Override')))
      if (type.getAttribute('PartName') === `/${name}`) type.parentNode.removeChild(type);
  }
  if (removed.length) {
    replacements['xl/_rels/workbook.xml.rels'] = serializeXml(relationships);
    replacements['[Content_Types].xml'] = serializeXml(contentTypes);
  }
  return { replacements, removed };
}

/** Style table edits use the same qualified writer with synthetic cell nodes;
 * worksheet bodies stay in the native index. */
export function indexedStyleMetadata(metadataBytes, edits, baseStyles) {
  const parts = unpack(metadataBytes),
    styles = parts['xl/styles.xml'] ? xml(parts['xl/styles.xml']) : null;
  const cells = new Map(),
    cache = new Map(),
    applyStyle = styles ? createWorkbookStyleWriter(styles) : null;
  let styleCount = styles ? children(elements(styles, 'cellXfs')[0], 'xf').length : 0;
  const fields = [
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
  ];
  for (const edit of edits) {
    if (edit.styleId === undefined && !fields.some((key) => edit[key] !== undefined)) continue;
    if (!styles) throw refusal('package-bad-input', 'This workbook has no editable style table');
    const key = `${edit.sheet}\0${edit.cell}`;
    let cell = cells.get(key);
    if (!cell) {
      cell = node(styles, 'c', { s: baseStyles[key] ?? 0 });
      cells.set(key, cell);
    }
    if (edit.styleId !== undefined) {
      if (edit.styleId >= styleCount) throw refusal('internal-invalid-field', 'Invalid source style');
      cell.setAttribute('s', String(edit.styleId));
    }
    const changes = Object.fromEntries(fields.filter((key) => edit[key] !== undefined).map((key) => [key, edit[key]]));
    if (Object.keys(changes).length) {
      const signature = cell.getAttribute('s') + JSON.stringify(changes);
      let index = cache.get(signature);
      if (index === undefined) {
        applyStyle(cell, changes);
        index = cell.getAttribute('s');
        cache.set(signature, index);
      } else cell.setAttribute('s', index);
      styleCount = Math.max(styleCount, Number(index) + 1);
    }
  }
  return {
    changes: [...cells].map(([key, cell]) => {
      const [sheet, ref] = key.split('\0');
      return { sheet, cell: ref, styleId: Number(cell.getAttribute('s')) };
    }),
    replacements: cells.size ? { 'xl/styles.xml': serializeXml(styles) } : {},
  };
}

/** The production Node adapter and non-Node native host use the same pure
 * resolver. Only unique style signatures cross its execution boundary. */
export function indexedStylePublication(metadataBytes, signatures) {
  if (!signatures.length) return { styleIds: [], replacements: {} };
  const parts = unpack(metadataBytes),
    bytes = parts['xl/styles.xml'];
  if (!bytes) throw refusal('package-bad-input', 'This workbook has no editable style table');
  const resolved = resolveWorkbookStyles(strFromU8(bytes), signatures);
  return { styleIds: resolved.styleIds, replacements: { 'xl/styles.xml': resolved.stylesXml } };
}

/** Layout editing touches one small metadata subtree; sheetData stays native. */
export function indexedLayoutFragment(metadataBytes, command, projected) {
  const result = applyProjectedOperation(metadataBytes, command, projected);
  const sheet = projected.sheets.find((s) => s.name === command.sheet);
  const parts = unpack(result.bytes),
    doc = xml(parts[sheet.path]);
  const tag = command.operation === 'resizeColumn' ? 'cols' : 'sheetViews';
  const element = elements(doc, tag)[0];
  if (!element) throw refusal('internal-missing-field', 'Missing projected layout fragment');
  return { tag, xml: serializeXml(element) };
}

/** Merge metadata is small. The native streaming writer validates all source
 * cells before accepting the fragment, including occupied cells in the range. */
export function indexedMergeFragment(metadataBytes, command) {
  const result = editWorkbookResult(metadataBytes, [command]);
  const { view } = readWorkbookMetadata(metadataBytes),
    sheet = view.sheets.find((s) => s.name === command.sheet);
  const doc = xml(unpack(result.bytes)[sheet.path]),
    element = elements(doc, 'mergeCells')[0];
  return element ? serializeXml(element) : '';
}

/** Plan package metadata against cell-free XML. Source cell guards are checked
 * by the native streaming writer; only table headers enter this small model. */
export function indexedPackageEdit(metadataBytes, command, view, headers = []) {
  const kind = command.operation,
    source = view.sheets.find((s) => s.name === command.sheet);
  let bytes = metadataBytes;
  if (kind === 'createTable' && headers.length) bytes = editWorkbookResult(bytes, headers).bytes;
  const before = unpack(metadataBytes),
    changed = editWorkbookResult(bytes, [command], undefined, false, undefined, false, {
      nativeRename: true,
      nativeDuplicate: kind === 'duplicateSheet',
    }),
    after = unpack(changed.bytes);
  // The generic edit path also drops Excel's calculation chain, because it
  // assumes formulas changed. A new table or highlight rule changes none, so
  // the chain and the two records that name it are put back as they were.
  // (Excel writes a chain into nearly every workbook that has formulas.)
  if (
    ['createTable', 'addConditionalRule'].includes(kind) &&
    before['xl/calcChain.xml'] &&
    !after['xl/calcChain.xml']
  ) {
    after['xl/calcChain.xml'] = before['xl/calcChain.xml'];
    after['xl/_rels/workbook.xml.rels'] = before['xl/_rels/workbook.xml.rels'];
    const types = xml(after['[Content_Types].xml']),
      chain = Array.from(xml(before['[Content_Types].xml']).getElementsByTagName('Override')).find(
        (item) => item.getAttribute('PartName') === '/xl/calcChain.xml'
      );
    if (chain) {
      const restored = types.createElementNS(types.documentElement.namespaceURI, 'Override');
      for (const attribute of Array.from(chain.attributes)) restored.setAttribute(attribute.name, attribute.value);
      types.documentElement.appendChild(restored);
      after['[Content_Types].xml'] =
        kind === 'addConditionalRule' ? before['[Content_Types].xml'] : strToU8(serializeXml(types));
    }
  }
  if (
    ['createTable', 'addConditionalRule'].includes(kind) &&
    !sameZipMember(before, after, 'xl/workbook.xml') &&
    !equalBytes(before['xl/workbook.xml'], after['xl/workbook.xml'])
  ) {
    // Table headers and conditional rules enter only the metadata simulator. Its generic edit
    // path requests full Automatic recalculation, which is not an authored
    // metadata change. Remove that delta only after proving every other workbook
    // field unchanged. Omitting the override preserves the actual private
    // export's accepted Automatic/Manual policy and readiness bytes.
    const original = xml(before['xl/workbook.xml']),
      simulated = xml(after['xl/workbook.xml']);
    const prior = children(original.documentElement, 'calcPr'),
      next = children(simulated.documentElement, 'calcPr');
    const scheduling = { fullCalcOnLoad: '1', forceFullCalc: '1', calcMode: 'auto' };
    if (
      prior.length > 1 ||
      next.length !== 1 ||
      Object.entries(scheduling).some(([key, value]) => next[0].getAttribute(key) !== value)
    )
      throw refusal('package-stale', 'Metadata simulation changed unqualified workbook scheduling metadata');
    if (prior.length) {
      for (const key of Object.keys(scheduling)) {
        if (prior[0].hasAttribute(key)) next[0].setAttribute(key, prior[0].getAttribute(key));
        else next[0].removeAttribute(key);
      }
    } else {
      if (
        Array.from(next[0].attributes).some(
          (attribute) =>
            !Object.hasOwn(scheduling, attribute.name) && !(attribute.name === 'xmlns' && attribute.value === NS)
        ) ||
        next[0].childNodes.length
      )
        throw refusal('package-stale', 'Metadata simulation changed unqualified workbook scheduling metadata');
      next[0].parentNode.removeChild(next[0]);
    }
    const signature = (node) => ({
      type: node.nodeType,
      name: node.nodeName,
      namespace: node.namespaceURI ?? null,
      data: node.data ?? node.nodeValue ?? null,
      attributes: Array.from(node.attributes ?? [])
        .map((attribute) => [attribute.name, attribute.namespaceURI ?? null, attribute.value])
        .sort((a, b) => a[0].localeCompare(b[0])),
      children: Array.from(node.childNodes ?? []).map(signature),
    });
    if (JSON.stringify(signature(original)) !== JSON.stringify(signature(simulated)))
      throw refusal('package-stale', 'Metadata simulation changed unrelated workbook metadata');
    after['xl/workbook.xml'] = before['xl/workbook.xml'];
  }
  const sheets = new Set(view.sheets.map((s) => s.path)),
    replacements = {},
    removed = [],
    copies = {},
    fragments = {};
  if (kind === 'duplicateSheet') {
    const next = readWorkbookMetadata(changed.bytes).view.sheets.find((s) => s.name === command.newName);
    copies[next.path] = source.path;
  }
  for (const name of Object.keys(before)) if (!after[name]) removed.push(name);
  for (const [name, data] of Object.entries(after)) {
    if (copies[name]) continue;
    if (kind === 'createTable' && name === 'xl/workbook.xml' && equalBytes(before[name], data)) continue;
    if (sheets.has(name)) {
      const old = xml(before[name]),
        next = xml(data),
        tags = {};
      for (const tag of ['conditionalFormatting', 'tableParts']) {
        const fragment = (d) => children(d.documentElement, tag).map(serializeXml).join('');
        if (fragment(old) !== fragment(next)) tags[tag] = fragment(next);
      }
      if (Object.keys(tags).length) fragments[name] = tags;
    } else if (!before[name] || !sameZipMember(before, after, name)) replacements[name] = strFromU8(data);
  }
  return {
    replacements,
    removed,
    copies,
    fragments,
    kind,
    sheet: source?.path,
    range: command.range,
    ...(kind === 'renameSheet' || kind === 'duplicateSheet'
      ? {
          rename: { old: command.sheet, new: command.newName },
          sheets: kind === 'duplicateSheet' ? [source.path] : view.sheets.map((s) => s.path),
          copyOnly: kind === 'duplicateSheet',
        }
      : {}),
  };
}

/** Append only new standard entries; existing workbook/OPC metadata stays byte-exact. */
export function indexedPreservedSheetAppend(metadataBytes, command, view) {
  if (command.operation !== 'addSheet')
    throw refusal('package-unsupported', 'Preserved structural editing only admits a new worksheet at the end');
  const plan = indexedPackageEdit(metadataBytes, command, view),
    parts = unpack(metadataBytes);
  for (const [part, container, element, key] of [
    ['xl/workbook.xml', 'sheets', 'sheet', 'name'],
    ['xl/_rels/workbook.xml.rels', null, 'Relationship', 'Id'],
    ['[Content_Types].xml', null, 'Override', 'PartName'],
  ]) {
    const raw = strFromU8(parts[part]),
      original = xml(parts[part], { locator: {}, normalizeLineEndings: (value) => value }),
      changed = xml(strToU8(plan.replacements[part]));
    const parent = container ? children(original.documentElement, container)[0] : original.documentElement,
      afterParent = container ? children(changed.documentElement, container)[0] : changed.documentElement;
    if (!parent || !afterParent) throw refusal('internal-missing-field', 'Missing append-only package container');
    const direct = (node, name) =>
      Array.from(node.childNodes).filter(
        (child) => child.nodeType === 1 && child.localName === name && child.namespaceURI === node.namespaceURI
      );
    const before = direct(parent, element),
      after = direct(afterParent, element),
      keys = new Set(before.map((node) => node.getAttribute(key))),
      added = after.filter((node) => !keys.has(node.getAttribute(key)));
    if (added.length !== 1 || after.length !== before.length + 1)
      throw refusal('package-stale', 'Worksheet append changed existing package identities');
    let entry = added[0];
    if (part === 'xl/workbook.xml') {
      const clean = changed.createElementNS(NS, entry.tagName);
      for (const name of ['name', 'sheetId']) clean.setAttribute(name, entry.getAttribute(name));
      clean.setAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'r:id',
        entry.getAttribute('r:id')
      );
      entry = clean;
    }
    const lines = lineStarts(raw);
    const start = lines[parent.lineNumber - 1] + parent.columnNumber - 1,
      end = xmlElementEnd(raw, start),
      close = raw.lastIndexOf('</' + parent.tagName + '>', end);
    if (close < start || close >= end)
      throw refusal('internal-missing-field', 'Missing append-only container boundary');
    plan.replacements[part] = raw.slice(0, close) + serializeXml(entry) + raw.slice(close);
  }
  return plan;
}
