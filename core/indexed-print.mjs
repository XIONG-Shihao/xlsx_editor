import { formatHashSalt, FORMAT } from './format.mjs';
/** Qualified small print definitions. Cell bodies stay in the accepted owner. */
import { lineStarts } from './xml-position.mjs';
import { createHash } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { openZipParts } from './zip-parts.mjs';
import { rangeBounds, columnName, cellPosition, assertXmlText } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import { defaultPrintSettings, validatePrintSettings } from './print-model.mjs';
import { createWorksheetDimensions } from './worksheet-geometry.mjs';
import { emptyPrintHeaderFooter, decodePrintSections, encodePrintSections } from './print-header-footer.mjs';
import { refusal } from './refusal-codes.mjs';
export { validatePrintSettings } from './print-model.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const children = (node) => Array.from(node.childNodes).filter((item) => item.nodeType === 1);
const defaults = defaultPrintSettings;
const sha = (text) => createHash('sha256').update(text).digest('hex');
const printTags = [
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'legacyDrawingHF',
  'customSheetViews',
];
const reserved = (node) =>
  node.namespaceURI === NS &&
  node.localName === 'definedName' &&
  /^_xlnm\.Print_(?:Area|Titles)$/i.test(node.getAttribute('name'));
const qualified = (node) => node?.namespaceURI === NS;
const xmlText = (value) => {
  assertXmlText(String(value));
  return String(value).replace(
    /[&<>"\r]/g,
    (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;' })[char]
  );
};
function exact(value, keys, label) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function finite(value, min, max, label) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function integer(value, min, max, label) {
  finite(value, min, max, label);
  if (!Number.isSafeInteger(value)) throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function parse(bytes, name) {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > 16 * 1024 ** 2)
    throw refusal('print-limit', 'Print metadata exceeds its bound');
  const source = strFromU8(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('print-malformed', 'Print XML entities are protected');
  const problems = [],
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (value) => value,
      errorHandler: {
        warning: (value) => problems.push(value),
        error: (value) => problems.push(value),
        fatalError: (value) => problems.push(value),
      },
    }).parseFromString(source, 'application/xml');
  if (problems.length || doc.documentElement?.namespaceURI !== NS || doc.documentElement?.localName !== name)
    throw refusal('print-malformed', 'Invalid print metadata XML');
  const lines = lineStarts(source);
  const raw = (node) => {
    const start = lines[node.lineNumber - 1] + node.columnNumber - 1;
    if (!Number.isSafeInteger(start) || !source.startsWith('<' + node.tagName, start))
      throw refusal('print-unsupported', 'Print XML has no qualified raw boundary');
    const end = xmlElementEnd(source, start);
    return { start, end, raw: source.slice(start, end), node };
  };
  return { source, doc, raw };
}
function attrs(node, keys) {
  for (const item of Array.from(node.attributes))
    if (item.namespaceURI !== 'http://www.w3.org/2000/xmlns/' && (item.namespaceURI || !keys.includes(item.localName)))
      throw refusal('print-protected', 'Extended ' + node.localName + ' print settings are protected');
}
function only(node, keys) {
  attrs(node, keys);
  if (children(node).length || node.textContent.trim())
    throw refusal('print-protected', 'Nested print settings are protected');
}
function bit(text, fallback = false) {
  if (text === null || text === '') return fallback;
  if (['1', 'true'].includes(text)) return true;
  if (['0', 'false'].includes(text)) return false;
  throw refusal('internal-invalid-field', 'Invalid print boolean');
}
function decimal(node, key, fallback, min = 0, max = 1584) {
  const raw = node.getAttribute(key);
  return raw ? finite(Number(raw), min, max, key) : fallback;
}
function sheetReference(text, sheet) {
  const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text);
  if (!match || (match[1]?.replaceAll("''", "'") ?? match[2]) !== sheet)
    throw refusal('print-protected', 'External or mismatched print reference is protected');
  return match[3];
}
function titleRanges(text, sheet) {
  // Commas inside quoted worksheet names are not range separators.
  const fields = [];
  let quote = false,
    start = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'") {
      if (quote && text[i + 1] === "'") {
        i++;
        continue;
      }
      quote = !quote;
    } else if (text[i] === ',' && !quote) {
      fields.push(text.slice(start, i));
      start = i + 1;
    }
  }
  fields.push(text.slice(start));
  if (quote || fields.length > 2) throw refusal('print-protected', 'Disconnected print titles are protected');
  let rows = null,
    columns = null;
  for (const field of fields) {
    const ref = sheetReference(field, sheet);
    let match = /^\$([1-9]\d*):\$([1-9]\d*)$/.exec(ref);
    if (match) {
      if (rows) throw refusal('print-duplicate', 'Duplicate print title rows');
      rows = { first: Number(match[1]), last: Number(match[2]) };
      continue;
    }
    match = /^\$([A-Z]{1,3}):\$([A-Z]{1,3})$/.exec(ref);
    if (!match || columns) throw refusal('print-unsupported', 'Unsupported print title range');
    columns = { first: match[1], last: match[2] };
  }
  return { repeatRows: rows, repeatColumns: columns };
}

function context(metadataBytes, view, sheetName, workbookXml) {
  if (!(metadataBytes instanceof Uint8Array) || metadataBytes.byteLength > 16 * 1024 ** 2)
    throw refusal('print-limit', 'Print metadata exceeds its bound');
  const sheet = view.sheets.find((item) => item.name === sheetName);
  if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
  const scope = view.sheets.indexOf(sheet),
    parts = openZipParts(metadataBytes, 16 * 1024 ** 2),
    worksheet = parse(parts[sheet.path], 'worksheet'),
    workbook = parse(
      workbookXml === undefined ? parts['xl/workbook.xml'] : new TextEncoder().encode(workbookXml),
      'workbook'
    );
  const direct = children(worksheet.doc.documentElement),
    find = (name) => {
      const nodes = direct.filter((node) => node.localName === name);
      if (nodes.length > 1 || nodes.some((node) => !qualified(node)))
        throw refusal('print-ambiguous', 'Ambiguous/foreign ' + name + ' print metadata');
      return nodes[0];
    },
    sheetPr = find('sheetPr'),
    setupChildren = sheetPr ? children(sheetPr).filter((node) => node.localName === 'pageSetUpPr') : [];
  if (setupChildren.length > 1 || setupChildren.some((node) => !qualified(node)))
    throw refusal('print-ambiguous', 'Ambiguous pageSetUpPr');
  const setupPr = setupChildren[0],
    records = printTags.map((name) => find(name));
  const containers = children(workbook.doc.documentElement).filter((node) => node.localName === 'definedNames');
  if (containers.length > 1 || containers.some((node) => !qualified(node)))
    throw refusal('print-ambiguous', 'Ambiguous print name container');
  const container = containers[0],
    names = container
      ? children(container)
          .filter(reserved)
          .filter((node) => !node.hasAttribute('localSheetId') || node.getAttribute('localSheetId') === String(scope))
      : [];
  const selected = [setupPr, ...records, ...names].filter(Boolean),
    raw = selected.map((node) => (node.ownerDocument === worksheet.doc ? worksheet : workbook).raw(node).raw),
    revision = selected.length
      ? sha([formatHashSalt('PRINT_DEFINITION_V1'), sheet.path, sheet.name, String(scope), ...raw].join('\0'))
      : null;
  return { parts, sheet, scope, worksheet, workbook, sheetPr, setupPr, records, container, names, revision };
}
function readDefinition(ctx) {
  const settings = defaults(),
    reasons = [];
  try {
    const { setupPr, records, names } = ctx,
      [options, margins, setup, header, rows, columns, legacy, custom] = records;
    if (rows || columns || legacy || custom)
      throw refusal('print-unsupported', 'Manual breaks or custom print views require qualification');
    if (header) {
      attrs(header, ['differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins']);
      if (
        bit(header.getAttribute('differentOddEven')) ||
        bit(header.getAttribute('differentFirst')) ||
        !bit(header.getAttribute('alignWithMargins'), true) ||
        Array.from(header.childNodes).some(
          (node) => (node.nodeType === 3 || node.nodeType === 4) && node.textContent.trim()
        )
      )
        throw refusal('print-protected', 'Advanced headers/footers are protected');
      const value = emptyPrintHeaderFooter(),
        seen = new Set();
      value.scaleWithDoc = bit(header.getAttribute('scaleWithDoc'), true);
      for (const node of children(header)) {
        if (!qualified(node) || !['oddHeader', 'oddFooter'].includes(node.localName) || seen.has(node.localName))
          throw refusal('print-protected', 'Advanced headers/footers are protected');
        attrs(node, []);
        if (children(node).length) throw refusal('print-protected', 'Nested headers/footers are protected');
        seen.add(node.localName);
        value[node.localName === 'oddHeader' ? 'header' : 'footer'] = decodePrintSections(node.textContent);
      }
      settings.headerFooter = value;
    }
    let fit = false;
    if (setupPr) {
      only(setupPr, ['fitToPage', 'autoPageBreaks']);
      fit = bit(setupPr.getAttribute('fitToPage'));
      if (!bit(setupPr.getAttribute('autoPageBreaks'), true))
        throw refusal('print-protected', 'Manual page setup is protected');
    }
    if (options) {
      only(options, ['gridLines', 'gridLinesSet', 'headings', 'horizontalCentered', 'verticalCentered']);
      if (['headings', 'horizontalCentered', 'verticalCentered'].some((key) => bit(options.getAttribute(key))))
        throw refusal('print-unsupported', 'Printed headings/centering require qualification');
      settings.gridlines = bit(options.getAttribute('gridLines'));
      bit(options.getAttribute('gridLinesSet'));
    }
    if (margins) {
      only(margins, ['top', 'right', 'bottom', 'left', 'header', 'footer']);
      for (const key of Object.keys(settings.marginsPt))
        settings.marginsPt[key] = decimal(margins, key, settings.marginsPt[key] / 72, 0, 2) * 72;
    }
    if (setup) {
      only(setup, ['paperSize', 'orientation', 'scale', 'fitToWidth', 'fitToHeight']);
      const paper = setup.getAttribute('paperSize');
      if (paper && paper !== '9' && paper !== '1')
        throw refusal('print-unsupported', 'Imported paper size requires qualification');
      settings.paper = paper === '1' ? 'Letter' : 'A4';
      settings.orientation = setup.getAttribute('orientation') || 'portrait';
      settings.scaling = fit
        ? {
            mode: 'fit',
            width: decimal(setup, 'fitToWidth', 1, 1, 10),
            height: decimal(setup, 'fitToHeight', 0, 0, 10),
          }
        : { mode: 'scale', percent: decimal(setup, 'scale', 100, 10, 400) };
    }
    if (fit && !setup) throw refusal('print-unsupported', 'Fit settings are incomplete');
    const seen = new Set();
    for (const node of names) {
      attrs(node, ['name', 'localSheetId']);
      if (node.getAttribute('localSheetId') !== String(ctx.scope) || children(node).length)
        throw refusal('print-protected', 'Global/extended print names are protected');
      const key = node.getAttribute('name').toLowerCase();
      if (seen.has(key)) throw refusal('print-duplicate', 'Duplicate reserved print definition');
      seen.add(key);
      if (key.endsWith('print_area')) {
        const range = sheetReference(node.textContent, ctx.sheet.name);
        if (!/^\$[A-Z]{1,3}\$[1-9]\d*:\$[A-Z]{1,3}\$[1-9]\d*$/.test(range))
          throw refusal('print-protected', 'Disconnected/alternative print areas are protected');
        settings.area = range.replaceAll('$', '');
      } else Object.assign(settings, titleRanges(node.textContent, ctx.sheet.name));
    }
    return { settings: validatePrintSettings(settings, { allowAbsentArea: true }), protectedReasons: reasons };
  } catch (error) {
    reasons.push(String(error.message ?? error));
    return { settings: null, protectedReasons: reasons };
  }
}
export function indexedPrintDefinition(metadataBytes, view, sheet) {
  const ctx = context(metadataBytes, view, sheet),
    result = readDefinition(ctx);
  return { sheet, revision: ctx.revision, canEdit: !result.protectedReasons.length, ...result };
}
export function isIndexedPrintCommand(command) {
  return command?.operation === 'setPrintSettings';
}
function applyRaw(source, patches) {
  let result = source;
  for (const patch of patches.sort((a, b) => b.start - a.start))
    result = result.slice(0, patch.start) + patch.text + result.slice(patch.end);
  return result;
}
function qname(node, name) {
  return (node.prefix ? node.prefix + ':' : '') + name;
}
function element(name, attrs = {}) {
  return (
    '<' +
    name +
    Object.entries(attrs)
      .map(([key, value]) => ' ' + key + '="' + xmlText(value) + '"')
      .join('') +
    '/>'
  );
}
export function indexedPrintPlan(metadataBytes, view, command, { workbookXml } = {}) {
  exact(command, ['operation', 'sheet', 'expectedPrintRevision', 'settings'], 'print command');
  if (!isIndexedPrintCommand(command)) throw refusal('internal-invalid-field', 'Invalid print command');
  const original = context(metadataBytes, view, command.sheet),
    ctx = workbookXml === undefined ? original : context(metadataBytes, view, command.sheet, workbookXml),
    previous = readDefinition(original);
  if (previous.protectedReasons.length)
    throw refusal('print-protected', 'Print settings are protected: ' + previous.protectedReasons.join('; '));
  if (command.expectedPrintRevision !== original.revision)
    throw refusal('print-stale', 'Conflict: print definition changed');
  const settings = validatePrintSettings(command.settings, { allowAbsentArea: true }),
    root = ctx.worksheet.doc.documentElement,
    fit = settings.scaling.mode === 'fit';
  const setup = element(qname(root, 'pageSetUpPr'), { fitToPage: fit ? '1' : '0', autoPageBreaks: '1' });
  let sheetPr;
  if (ctx.sheetPr) {
    const record = ctx.worksheet.raw(ctx.sheetPr);
    if (ctx.setupPr) {
      const child = ctx.worksheet.raw(ctx.setupPr);
      sheetPr = applyRaw(record.raw, [
        { start: child.start - record.start, end: child.end - record.start, text: setup },
      ]);
    } else {
      const raw = record.raw;
      if (/\/\s*>$/.test(raw)) {
        sheetPr = raw.replace(/\/\s*>$/, '>') + setup + '</' + ctx.sheetPr.tagName + '>';
      } else {
        const at = raw.lastIndexOf('</');
        sheetPr = raw.slice(0, at) + setup + raw.slice(at);
      }
    }
  } else sheetPr = '<' + qname(root, 'sheetPr') + '>' + setup + '</' + qname(root, 'sheetPr') + '>';
  const fragments = {
    sheetPr,
    printOptions: element(qname(root, 'printOptions'), {
      gridLines: settings.gridlines ? '1' : '0',
      gridLinesSet: '1',
    }),
    pageMargins: element(
      qname(root, 'pageMargins'),
      Object.fromEntries(Object.entries(settings.marginsPt).map(([key, value]) => [key, String(value / 72)]))
    ),
    pageSetup: element(qname(root, 'pageSetup'), {
      paperSize: settings.paper === 'A4' ? '9' : '1',
      orientation: settings.orientation,
      ...(fit
        ? { fitToWidth: settings.scaling.width, fitToHeight: settings.scaling.height }
        : { scale: settings.scaling.percent }),
    }),
  };
  if (settings.headerFooter) {
    const hf = settings.headerFooter,
      tag = qname(root, 'headerFooter');
    fragments.headerFooter =
      '<' +
      tag +
      ' differentOddEven="0" differentFirst="0" scaleWithDoc="' +
      (hf.scaleWithDoc ? '1' : '0') +
      '" alignWithMargins="1">' +
      ['header', 'footer']
        .map((side) => {
          const child = qname(root, side === 'header' ? 'oddHeader' : 'oddFooter');
          return '<' + child + '>' + xmlText(encodePrintSections(hf[side])) + '</' + child + '>';
        })
        .join('') +
      '</' +
      tag +
      '>';
  }
  const b = settings.area === null ? null : rangeBounds(settings.area),
    sheetQuote = "'" + ctx.sheet.name.replaceAll("'", "''") + "'!",
    names = b
      ? [
          [
            '_xlnm.Print_Area',
            sheetQuote +
              '$' +
              columnName(b.first.col) +
              '$' +
              (b.first.row + 1) +
              ':$' +
              columnName(b.last.col) +
              '$' +
              (b.last.row + 1),
          ],
        ]
      : [];
  const titles = [];
  if (settings.repeatRows) titles.push(sheetQuote + '$' + settings.repeatRows.first + ':$' + settings.repeatRows.last);
  if (settings.repeatColumns)
    titles.push(sheetQuote + '$' + settings.repeatColumns.first + ':$' + settings.repeatColumns.last);
  if (titles.length) names.push(['_xlnm.Print_Titles', titles.join(',')]);
  const workbookRoot = ctx.workbook.doc.documentElement,
    nameTag = qname(workbookRoot, 'definedName'),
    nameText = names
      .map(
        ([name, formula]) =>
          '<' +
          nameTag +
          ' name="' +
          name +
          '" localSheetId="' +
          ctx.scope +
          '">' +
          xmlText(formula) +
          '</' +
          nameTag +
          '>'
      )
      .join('');
  let workbook;
  if (ctx.container) {
    const record = ctx.workbook.raw(ctx.container);
    if (/\/\s*>$/.test(record.raw)) {
      workbook = applyRaw(ctx.workbook.source, [
        {
          start: record.start,
          end: record.end,
          text: record.raw.replace(/\/\s*>$/, '>') + nameText + '</' + ctx.container.tagName + '>',
        },
      ]);
    } else {
      const patches = ctx.names.map((node) => {
          const item = ctx.workbook.raw(node);
          return { start: item.start, end: item.end, text: '' };
        }),
        at = record.start + record.raw.lastIndexOf('</');
      patches.push({ start: at, end: at, text: nameText });
      workbook = applyRaw(ctx.workbook.source, patches);
    }
  } else if (!nameText) {
    workbook = ctx.workbook.source;
  } else {
    const tag = qname(workbookRoot, 'definedNames'),
      following = children(workbookRoot).find((node) =>
        [
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
        ].includes(node.localName)
      ),
      at = following
        ? ctx.workbook.raw(following).start
        : ctx.workbook.raw(workbookRoot).start + ctx.workbook.raw(workbookRoot).raw.lastIndexOf('</');
    workbook = applyRaw(ctx.workbook.source, [
      { start: at, end: at, text: '<' + tag + '>' + nameText + '</' + tag + '>' },
    ]);
  }
  const plan = {
    kind: 'setPrintSettings',
    sheetPart: ctx.sheet.path,
    sheetName: ctx.sheet.name,
    scope: ctx.scope,
    priorPrintRevision: original.revision,
    settings,
    replacements: { 'xl/workbook.xml': workbook },
    removed: [],
    copies: {},
    fragments: { [ctx.sheet.path]: fragments },
  };
  if (Buffer.byteLength(JSON.stringify(plan)) > 2 * 1024 ** 2)
    throw refusal('print-limit', 'Print metadata plan exceeds its transport bound');
  return {
    plan,
    command: {
      operation: 'setPrintSettings',
      sheet: ctx.sheet.name,
      expectedPrintRevision: original.revision,
      settings,
    },
  };
}
/** Metadata-derived physical defaults and area geometry; no cell projection. */
export function indexedPrintGeometry(metadataBytes, view, sheetName, settings, geometryProfile) {
  const ctx = context(metadataBytes, view, sheetName),
    bounds = rangeBounds(settings.area);
  if (
    (ctx.sheet.rowHeights?.length ?? 0) > 100000 ||
    (ctx.sheet.hiddenRows?.length ?? 0) > 100000 ||
    (ctx.sheet.visibleRows?.length ?? 0) > 100000 ||
    (ctx.sheet.widths?.length ?? 0) > 10000 ||
    (ctx.sheet.merges?.length ?? 0) > 10000
  )
    throw Object.assign(Error('Print geometry exceeds this bounded slice'), { code: 'PRINT_GEOMETRY_BOUND' });
  if (ctx.sheet.geometryIssues?.length)
    throw Object.assign(Error('Worksheet dimensions are unresolved'), {
      code: 'PRINT_UNRESOLVED_GEOMETRY',
      reasons: ctx.sheet.geometryIssues,
    });
  const model = createWorksheetDimensions(ctx.sheet, geometryProfile),
    rows = [],
    columns = [];
  for (let row = bounds.first.row; row <= bounds.last.row; row++) {
    const heightPt = model.rowHeightPt(row);
    rows.push({ index: row, heightPt, hidden: model.rowHidden(row) || heightPt === 0 });
  }
  for (let col = bounds.first.col; col <= bounds.last.col; col++) {
    const pixels = model.columnPixels(col);
    columns.push({ index: col, widthPt: pixels * 0.75, hidden: pixels === 0 });
  }
  const distance = (first, last, dimension) => {
    if (dimension === 'row') return model.rowDistancePt(first, last);
    const sign = first <= last ? 1 : -1;
    let sum = 0;
    for (let col = Math.min(first, last); col < Math.max(first, last); col++) sum += model.columnPixels(col) * 0.75;
    return sign * sum;
  };
  const objectBox = (anchor) => {
    const at = cellPosition(anchor.cell);
    if (model.rowHeightPt(at.row) === 0 || model.columnPixels(at.col) === 0) return null;
    return [
      distance(bounds.first.col, at.col, 'column') + anchor.dxEmu / 12700,
      distance(bounds.first.row, at.row, 'row') + anchor.dyEmu / 12700,
      anchor.widthEmu / 12700,
      anchor.heightEmu / 12700,
    ];
  };
  return {
    geometryProfile: model.profile,
    rows,
    columns,
    objectBox,
    merges: (ctx.sheet.merges ?? []).filter((range) => {
      const b = rangeBounds(range);
      return (
        b.first.row <= bounds.last.row &&
        b.last.row >= bounds.first.row &&
        b.first.col <= bounds.last.col &&
        b.last.col >= bounds.first.col
      );
    }),
  };
}

export function indexedPrintReferenceGuard() {
  // A print area and repeated titles are defined names. They move with inserted
  // and deleted rows and columns, follow a sheet rename and are copied with a
  // duplicated sheet, like any other name; nothing about them needs a refusal.
  return null;
}
