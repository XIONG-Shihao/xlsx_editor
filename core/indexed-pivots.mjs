import { FORMAT, matchesFormat } from './format.mjs';
/** Real, manually refreshed worksheet PivotTables with a closed package/cache profile. */
import { lineStarts } from './xml-position.mjs';
import path from 'node:path';
import { pivotPresentation, pivotLayoutSample, pivotOutputData } from './pivot-layout-samples.mjs';
import { createHash } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { openZipParts } from './zip-parts.mjs';
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import {
  createPivotModel,
  pivotParts,
  pivotIdentifier,
  pivotEscape,
  PIVOT_NS as S,
  PIVOT_REL as R,
  PIVOT_TYPES,
} from './pivot-model.mjs';
import { qualifyPivotTarget, replacePivotOutput } from './pivot-target.mjs';
import { refusal } from './refusal-codes.mjs';
const P = 'http://schemas.openxmlformats.org/package/2006/relationships',
  T = 'http://schemas.openxmlformats.org/package/2006/content-types';
const children = (n) => Array.from(n?.childNodes ?? []).filter((x) => x.nodeType === 1),
  hash = (b) => createHash('sha256').update(b).digest('hex');
function parse(bytes, ns, root) {
  const raw = typeof bytes === 'string' ? bytes : strFromU8(bytes ?? new Uint8Array());
  if (Buffer.byteLength(raw) > 16 * 1024 ** 2 || /<!DOCTYPE|<!ENTITY/i.test(raw))
    throw refusal('internal-invalid-field', 'Invalid bounded pivot metadata');
  const problems = [],
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (v) => v,
      errorHandler: {
        warning: (e) => problems.push(e),
        error: (e) => problems.push(e),
        fatalError: (e) => problems.push(e),
      },
    }).parseFromString(raw, 'application/xml');
  if (problems.length || doc.documentElement.namespaceURI !== ns || doc.documentElement.localName !== root)
    throw refusal('pivot-malformed', 'Invalid pivot package XML');
  return { raw, doc };
}
function span(raw, n) {
  const starts = lineStarts(raw);
  const a = starts[n.lineNumber - 1] + n.columnNumber - 1;
  if (!Number.isSafeInteger(a) || !raw.startsWith('<' + n.tagName, a))
    throw refusal('internal-missing-field', 'Missing raw pivot metadata boundary');
  return [a, xmlElementEnd(raw, a)];
}
function append(raw, node, addition) {
  const [a, b] = span(raw, node),
    fragment = raw.slice(a, b),
    prefix = node.prefix ? node.prefix + ':' : '';
  if (fragment.endsWith('/>'))
    return (
      raw.slice(0, a) + fragment.slice(0, -2) + '>' + addition + '</' + prefix + node.localName + '>' + raw.slice(b)
    );
  const close = fragment.lastIndexOf('</');
  if (close < 0) throw refusal('pivot-malformed', 'Invalid pivot metadata container');
  return raw.slice(0, a) + fragment.slice(0, close) + addition + fragment.slice(close) + raw.slice(b);
}
const relationship = (id, type, target) =>
  '<Relationship xmlns="' +
  P +
  '" Id="' +
  pivotEscape(id) +
  '" Type="' +
  R +
  '/' +
  type +
  '" Target="' +
  pivotEscape(target) +
  '"/>';
const rels = (content) => '<Relationships xmlns="' + P + '">' + content + '</Relationships>';
const relPath = (file) => path.posix.join(path.posix.dirname(file), '_rels', path.posix.basename(file) + '.rels');
const sharedValue = (n) =>
  n.localName === 'm'
    ? null
    : n.localName === 's'
      ? n.getAttribute('v')
      : n.localName === 'b'
        ? n.getAttribute('v') === '1'
        : n.localName === 'n'
          ? Number(n.getAttribute('v'))
          : undefined;
export const isIndexedPivotCommand = (c) =>
  ['createPivot', 'refreshPivot', 'updatePivot', 'setPivotLayout'].includes(c?.operation);
function cacheModel(parts, definitionFile) {
  const { doc } = parse(parts[definitionFile], S, 'pivotCacheDefinition'),
    root = doc.documentElement,
    source = children(root).find((n) => n.localName === 'cacheSource'),
    ws = children(source ?? {}).find((n) => n.localName === 'worksheetSource'),
    container = children(root).find((n) => n.localName === 'cacheFields');
  if (!ws || !container) throw refusal('pivot-bad-input', 'Pivot has no worksheet source/cache fields');
  const fields = children(container).map((f) => ({
    name: f.getAttribute('name'),
    items: children(children(f).find((n) => n.localName === 'sharedItems') ?? {}).map(sharedValue),
  }));
  const rr = parse(parts[relPath(definitionFile)], P, 'Relationships'),
    links = children(rr.doc.documentElement);
  if (
    links.length !== 1 ||
    links[0].getAttribute('Id') !== 'rIdRecords' ||
    links[0].getAttribute('Type') !== R + '/pivotCacheRecords' ||
    links[0].hasAttribute('TargetMode')
  )
    throw refusal('pivot-unsupported', 'Unqualified pivot record relationship');
  const recordFile = path.posix.normalize(
      path.posix.join(path.posix.dirname(definitionFile), links[0].getAttribute('Target'))
    ),
    records = children(parse(parts[recordFile], S, 'pivotCacheRecords').doc.documentElement).map((r) =>
      children(r).map((x) => Number(x.getAttribute('v')))
    );
  const b = rangeBounds(ws.getAttribute('ref'));
  if (records.length !== b.rows - 1 || fields.length !== b.columns || records.some((r) => r.length !== fields.length))
    throw refusal('pivot-stale', 'Pivot cache geometry changed');
  const cells = {};
  fields.forEach((f, i) => {
    cells[columnName(b.first.col + i) + (b.first.row + 1)] = { value: f.name };
  });
  records.forEach((r, i) =>
    r.forEach((x, c) => {
      if (!Number.isSafeInteger(x) || x < 0 || x >= fields[c].items.length || fields[c].items[x] === undefined)
        throw refusal('internal-invalid-field', 'Invalid pivot shared item');
      cells[columnName(b.first.col + c) + (b.first.row + i + 2)] = { value: fields[c].items[x] };
    })
  );
  return { sheet: ws.getAttribute('sheet'), range: ws.getAttribute('ref'), cells, recordFile };
}
export function indexedPivotInventory(metadataBytes, view) {
  const parts = openZipParts(metadataBytes),
    found = [];
  const types = parse(parts['[Content_Types].xml'], T, 'Types'),
    files = children(types.doc.documentElement)
      .filter((n) => n.getAttribute('ContentType') === PIVOT_TYPES.table)
      .map((n) => n.getAttribute('PartName').replace(/^\//, ''));
  if (files.length > 32)
    return [
      {
        name: 'Pivot inventory limit',
        sheet: null,
        revision: null,
        canEdit: false,
        reason: 'More than 32 pivot definitions require separate qualification',
      },
    ];
  for (const file of files) {
    let targetSheet;
    try {
      const targetSheets = view.sheets.filter((sheet) => {
        const part = parts[relPath(sheet.path)];
        if (!part) return false;
        return children(parse(part, P, 'Relationships').doc.documentElement).some(
          (n) =>
            n.getAttribute('Type') === R + '/pivotTable' &&
            !n.hasAttribute('TargetMode') &&
            // A target may be written from the package root ("/xl/pivotTables/...") or relative to the sheet.
            (n.getAttribute('Target').startsWith('/')
              ? n.getAttribute('Target').slice(1)
              : path.posix.normalize(path.posix.join(path.posix.dirname(sheet.path), n.getAttribute('Target')))) ===
              file
        );
      });
      if (targetSheets.length === 1) targetSheet = targetSheets[0];
      if (!matchesFormat(file, 'XL_PIVOTTABLES_PIVOT1_XML'))
        throw refusal('pivot-protected', 'Imported pivot definition is preserved and protected');
      const raw = strFromU8(parts[file]),
        root = parse(raw, S, 'pivotTableDefinition').doc.documentElement,
        name = root.getAttribute('name'),
        cacheId = Number(root.getAttribute('cacheId'));
      const links = children(parse(parts[relPath(file)], P, 'Relationships').doc.documentElement);
      if (
        links.length !== 1 ||
        links[0].getAttribute('Type') !== R + '/pivotCacheDefinition' ||
        links[0].hasAttribute('TargetMode')
      )
        throw refusal('pivot-unsupported', 'Unqualified pivot cache link');
      const definitionFile = path.posix.normalize(
          path.posix.join(path.posix.dirname(file), links[0].getAttribute('Target'))
        ),
        cached = cacheModel(parts, definitionFile),
        row = Number(children(children(root).find((n) => n.localName === 'rowFields') ?? {})[0]?.getAttribute('x')),
        data = children(children(root).find((n) => n.localName === 'dataFields') ?? {})[0];
      if (!data) throw refusal('pivot-bad-input', 'Pivot has no aggregate');
      const model = createPivotModel({
          ...cached,
          rowField: row,
          valueField: Number(data.getAttribute('fld')),
          aggregate: data.getAttribute('subtotal'),
        }),
        presentation = pivotPresentation(root),
        expected = pivotParts(model, { name, cacheId, ...presentation });
      if (
        !targetSheet ||
        raw !== expected.table ||
        strFromU8(parts[definitionFile]) !== expected.definition ||
        strFromU8(parts[cached.recordFile]) !== expected.records
      )
        throw refusal('pivot-stale', 'Imported or changed pivot profile is protected');
      const target = parse(parts[targetSheet.path], S, 'worksheet').doc.documentElement;
      if (
        children(target).some(
          (n) => !['dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData'].includes(n.localName)
        )
      )
        throw refusal('pivot-unsupported', 'Pivot output has additional metadata that must be preserved');
      const rawTarget = strFromU8(parts[targetSheet.path]),
        skeleton = (raw, root) => {
          const sd = children(root).find((n) => n.localName === 'sheetData');
          if (!sd || Array.from(sd.attributes ?? []).some((a) => !a.name.startsWith('xmlns')))
            throw refusal('pivot-protected', 'Extended pivot output data attributes are protected');
          const [a, b] = span(raw, sd);
          return raw.slice(0, a) + '<sheetData/>' + raw.slice(b);
        };
      qualifyPivotTarget(rawTarget, expected.worksheet, expected.output);
      if ((targetSheet.hiddenRows ?? []).length || (targetSheet.rowHeights ?? []).length)
        throw refusal('pivot-stale', 'Pivot output display settings changed; preserved and protected');
      const revision = hash(
        JSON.stringify([raw, strFromU8(parts[definitionFile]), strFromU8(parts[cached.recordFile])])
      );
      found.push({
        name,
        cacheId,
        sheet: targetSheet.name,
        targetPath: targetSheet.path,
        sourceSheet: model.sheet,
        sourceRange: model.range,
        rowField: model.rowField,
        valueField: model.valueField,
        aggregate: model.aggregate,
        ...presentation,
        revision,
        outputRange: expected.output,
        parts: { table: file, definition: definitionFile, records: cached.recordFile },
        model,
        canEdit: true,
      });
    } catch (error) {
      // Preserve a safely bound location for protected-pivot inspection. This
      // supplies no edit capability and cannot invent an unknown output owner.
      let geometry = {};
      try {
        const root = parse(parts[file], S, 'pivotTableDefinition').doc.documentElement,
          locations = children(root).filter((n) => n.namespaceURI === S && n.localName === 'location');
        if (targetSheet && locations.length === 1) {
          const range = locations[0].getAttribute('ref'),
            b = rangeBounds(range);
          if (b.first.row <= b.last.row && b.first.col <= b.last.col)
            geometry = {
              outputRange: range,
              ...(pivotIdentifier(root.getAttribute('name')) ? { name: root.getAttribute('name') } : {}),
            };
        }
      } catch {}
      found.push({
        name: file,
        sheet: targetSheet?.name ?? null,
        ...geometry,
        revision: null,
        canEdit: false,
        reason: String(error.message ?? error),
      });
    }
  }
  return found;
}
export function indexedPivotMutationGuard(metadataBytes, view, command) {
  if (isIndexedPivotCommand(command)) return;
  const pivots = indexedPivotInventory(metadataBytes, view).filter(
    (pivot) => pivot.sheet === command.sheet || pivot.sheet === null
  );
  if (!pivots.length) return;
  // Excel protects the cells a pivot table occupies, not the sheet it sits on:
  // typing, formatting, merging or sorting beside the pivot is ordinary work.
  // Inserting or deleting rows and columns would have to move the pivot, which
  // is not done here, so that is still refused, as is anything on a sheet whose
  // pivot cannot be located.
  const located = pivots.every((pivot) => pivot.sheet && pivot.outputRange),
    areas = [...(command.cells ?? []).map((edit) => edit.cell), command.cell, command.range, command.source].filter(
      (area) => typeof area === 'string' && (command.sourceSheet === undefined || area !== command.source)
    ),
    clear = (area) =>
      pivots.every((pivot) => {
        const a = rangeBounds(pivot.outputRange),
          b = rangeBounds(area.includes(':') ? area : `${area}:${area}`);
        return (
          a.last.row < b.first.row || b.last.row < a.first.row || a.last.col < b.first.col || b.last.col < a.first.col
        );
      });
  if (located && ['resizeColumn', 'resizeRow', 'freeze'].includes(command.operation)) return;
  // Formatting the cells of a pivot table (bold headers, a fill, number formats)
  // is allowed in Excel; only their values belong to the pivot.
  if (located && command.type === 'formatRange') return;
  if (located && areas.length && areas.every(clear)) return;
  throw refusal(
    'pivot-refused',
    located
      ? 'This would change a pivot table or move it; edit its source data and refresh the pivot instead'
      : 'Pivot output worksheet is managed/protected; edit qualified source values and refresh the pivot'
  );
}
export function capturePivotCommands(commands) {
  if (!Array.isArray(commands)) throw refusal('pivot-bad-input', 'Pivot edits require commands');
  const copied = structuredClone(commands);
  for (const c of copied) {
    if (!isIndexedPivotCommand(c)) continue;
    const fields =
      c.operation === 'createPivot'
        ? [
            'operation',
            'sheet',
            'name',
            'range',
            'rowField',
            'valueField',
            'aggregate',
            'pivotSheet',
            'expectedPivotRevision',
          ]
        : c.operation === 'setPivotLayout'
          ? ['operation', 'sheet', 'name', 'sample', 'grandTotals', 'expectedPivotRevision']
          : c.operation === 'updatePivot'
            ? [
                'operation',
                'sheet',
                'name',
                'sourceSheet',
                'range',
                'rowField',
                'valueField',
                'aggregate',
                'expectedPivotRevision',
              ]
            : ['operation', 'sheet', 'name', 'expectedPivotRevision'];
    if (
      !c ||
      Array.isArray(c) ||
      Object.keys(c).some((k) => !fields.includes(k)) ||
      (['updatePivot', 'setPivotLayout'].includes(c.operation) &&
        (Object.keys(c).length !== fields.length || fields.some((k) => !Object.hasOwn(c, k))))
    )
      throw refusal('internal-invalid-field', 'Invalid pivot command fields');
  }
  return copied;
}
export async function indexedPivotPlan({ metadataBytes, view, command, readRange, readWorksheet }) {
  if (
    !isIndexedPivotCommand(command) ||
    Object.keys(command).some(
      (k) =>
        ![
          'operation',
          'sheet',
          'name',
          'range',
          'rowField',
          'valueField',
          'aggregate',
          'pivotSheet',
          'expectedPivotRevision',
          'sourceSheet',
          'sample',
          'grandTotals',
        ].includes(k)
    )
  )
    throw refusal('internal-invalid-field', 'Invalid pivot command');
  command = capturePivotCommands([command])[0];
  const parts = openZipParts(metadataBytes),
    inventory = indexedPivotInventory(metadataBytes, view),
    layoutOnly = command.operation === 'setPivotLayout',
    editing = command.operation === 'updatePivot',
    refresh = command.operation !== 'createPivot';
  let sample = 'tabular-light',
    grandTotals = true,
    ownedModel;
  let sourceSheet,
    sourceRange,
    rowField,
    valueField,
    aggregate,
    targetSheet,
    targetPath,
    cacheId,
    files,
    priorRevision = null,
    name = command.name;
  if (refresh) {
    const current = inventory.find((p) => p.name === name && p.sheet === command.sheet);
    if (!current?.canEdit || current.revision !== command.expectedPivotRevision)
      throw refusal('internal-missing-field', 'Conflict: pivot is protected, changed or missing');
    ({ sourceSheet, sourceRange, rowField, valueField, aggregate, cacheId, sample, grandTotals } = current);
    ownedModel = current.model;
    targetSheet = current.sheet;
    targetPath = current.targetPath;
    files = current.parts;
    priorRevision = current.revision;
    if (editing) {
      if (inventory.length !== 1) throw refusal('pivot-bad-input', 'Field editing requires one owned pivot');
      sourceSheet = command.sourceSheet;
      sourceRange = command.range;
      rowField = command.rowField;
      valueField = command.valueField;
      aggregate = command.aggregate;
    }
    if (layoutOnly) {
      if (inventory.length !== 1) throw refusal('pivot-bad-input', 'Layout editing requires one owned pivot');
      sample = pivotLayoutSample(command.sample).id;
      if (typeof command.grandTotals !== 'boolean')
        throw refusal('pivot-bad-input', 'Pivot grand totals require an explicit Boolean');
      grandTotals = command.grandTotals;
    }
    const data = pivotOutputData(current.model, current);
    for (let r = 0; r < data.length; r += 500) {
      const page = await readRange(targetSheet, 'A' + (r + 1) + ':B' + Math.min(data.length, r + 500));
      for (let i = r; i < Math.min(data.length, r + 500); i++)
        for (let c = 0; c < 2; c++) {
          const cell = page.cells[columnName(c) + (i + 1)];
          if (
            !cell ||
            cell.formula ||
            cell.formulaType ||
            cell.richText ||
            (cell.styleId !== undefined && cell.styleId !== 0) ||
            cell.value !== data[i][c]
          )
            throw refusal(
              'pivot-stale',
              'Pivot output cells changed outside its managed cache; preserved and protected'
            );
        }
    }
  } else {
    if (
      command.expectedPivotRevision !== null ||
      inventory.length ||
      Object.keys(parts).some((p) => /^xl\/pivot(?:Tables|Cache)\//.test(p))
    )
      throw refusal('pivot-unsupported', 'Creating a pivot requires no existing pivot cache/definition');
    if (
      !pivotIdentifier(name) ||
      (view.tables ?? []).some((t) => t.name.toLowerCase() === name.toLowerCase()) ||
      (view.namedRanges ?? []).some((n) => n.name.toLowerCase() === name.toLowerCase())
    )
      throw refusal('internal-invalid-field', 'Invalid or colliding pivot name');
    sourceSheet = command.sheet;
    sourceRange = command.range;
    rowField = command.rowField;
    valueField = command.valueField;
    aggregate = command.aggregate ?? 'sum';
    targetSheet = command.pivotSheet;
    if (
      typeof targetSheet !== 'string' ||
      !targetSheet ||
      targetSheet.length > 31 ||
      /[\\/*?:\[\]\x00-\x1f]/.test(targetSheet) ||
      targetSheet.startsWith("'") ||
      targetSheet.endsWith("'") ||
      view.sheets.some((s) => s.name.toLowerCase() === targetSheet.toLowerCase())
    )
      throw refusal('pivot-bad-input', 'Choose a distinct valid pivot worksheet name');
    const used = new Set(Object.keys(parts));
    let id = 1;
    while (used.has('xl/worksheets/sheet' + id + '.xml')) id++;
    targetPath = 'xl/worksheets/sheet' + id + '.xml';
    cacheId = 1;
    files = {
      table: FORMAT.XL_PIVOTTABLES_PIVOT1_XML,
      definition: FORMAT.XL_PIVOTCACHE_CACHEDEFINITION1_XML,
      records: FORMAT.XL_PIVOTCACHE_CACHERECORDS1_XML,
    };
  }
  const source = view.sheets.find((s) => s.name === sourceSheet);
  if (!source) throw refusal('internal-missing-field', 'Pivot source worksheet is missing');
  const b = rangeBounds(sourceRange),
    cells = {};
  if (b.size > 8000) throw refusal('pivot-limit', 'Pivot source exceeds its bounded rectangle');
  for (let r = b.first.row; !layoutOnly && r <= b.last.row; ) {
    const end = Math.min(b.last.row, r + Math.max(1, Math.floor(1000 / b.columns)) - 1),
      page = await readRange(sourceSheet, columnName(b.first.col) + (r + 1) + ':' + columnName(b.last.col) + (end + 1));
    Object.assign(cells, page.cells);
    r = end + 1;
  }
  const model = layoutOnly
      ? ownedModel
      : createPivotModel({ sheet: sourceSheet, range: sourceRange, cells, rowField, valueField, aggregate }),
    xml = pivotParts(model, { name, cacheId, sample, grandTotals }),
    replacements = {
      [files.table]: xml.table,
      [files.definition]: xml.definition,
      [files.records]: xml.records,
      [targetPath]: xml.worksheet,
    };
  if (refresh) {
    const current = inventory.find((p) => p.name === name && p.sheet === command.sheet),
      prior = pivotParts(current.model, { name, cacheId, sample: current.sample, grandTotals: current.grandTotals });
    if (typeof readWorksheet !== 'function')
      throw refusal('pivot-bad-input', 'Pivot refresh requires its actual bounded source worksheet');
    const target = await readWorksheet(targetPath, { maxBytes: 16 * 1024 ** 2 });
    replacements[targetPath] = replacePivotOutput(
      typeof target === 'string' ? target : strFromU8(target),
      prior.worksheet,
      prior.output,
      xml.worksheet,
      xml.output
    );
  }
  if (!refresh) {
    // Cache export may normalize calcPr independently of the original metadata
    // skeleton. Preserve the exact workbook envelope used by this package edit.
    if (typeof readWorksheet === 'function')
      parts['xl/workbook.xml'] = await readWorksheet('xl/workbook.xml', { maxBytes: 16 * 1024 ** 2 });
    for (const file of [
      targetPath,
      ...Object.values(files),
      relPath(targetPath),
      relPath(files.table),
      relPath(files.definition),
    ])
      if (parts[file]) throw refusal('pivot-duplicate', 'Pivot destination part already exists');
    const wb = parse(parts['xl/workbook.xml'], S, 'workbook'),
      sheets = children(wb.doc.documentElement).find((n) => n.localName === 'sheets');
    if (
      !sheets ||
      children(wb.doc.documentElement).some((n) =>
        ['pivotCaches', 'extLst', 'externalReferences'].includes(n.localName)
      )
    )
      throw refusal('pivot-unsupported', 'Extended workbook/pivot metadata requires qualification');
    const wbRels = parse(parts['xl/_rels/workbook.xml.rels'], P, 'Relationships'),
      ids = new Set(children(wbRels.doc.documentElement).map((n) => n.getAttribute('Id')));
    let n = 1;
    while (ids.has(FORMAT.RIDPIVOTSHEET + n) || ids.has(FORMAT.RIDPIVOTCACHE + n)) n++;
    const sheetRel = FORMAT.RIDPIVOTSHEET + n,
      cacheRel = FORMAT.RIDPIVOTCACHE + n,
      sheetId =
        Math.max(
          0,
          ...view.sheets.map((s) => {
            const nodes = children(sheets);
            return Number(nodes.find((x) => x.getAttribute('name') === s.name)?.getAttribute('sheetId')) || 0;
          })
        ) + 1;
    let workbook = append(
      wb.raw,
      sheets,
      '<sheet xmlns="' +
        S +
        '" xmlns:r="' +
        R +
        '" name="' +
        pivotEscape(targetSheet) +
        '" sheetId="' +
        sheetId +
        '" r:id="' +
        sheetRel +
        '"/>'
    );
    const after = parse(workbook, S, 'workbook');
    const pivotCaches =
      '<pivotCaches xmlns="' +
      S +
      '" xmlns:r="' +
      R +
      '"><pivotCache cacheId="' +
      cacheId +
      '" r:id="' +
      cacheRel +
      '"/></pivotCaches>';
    const following = children(after.doc.documentElement).find((n) =>
      ['smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'].includes(
        n.localName
      )
    );
    if (following) {
      const [at] = span(workbook, following);
      workbook = workbook.slice(0, at) + pivotCaches + workbook.slice(at);
    } else workbook = append(workbook, after.doc.documentElement, pivotCaches);
    replacements['xl/workbook.xml'] = workbook;
    replacements['xl/_rels/workbook.xml.rels'] = append(
      wbRels.raw,
      wbRels.doc.documentElement,
      relationship(sheetRel, 'worksheet', targetPath.slice(3)) +
        relationship(cacheRel, 'pivotCacheDefinition', files.definition.slice(3))
    );
    const types = parse(parts['[Content_Types].xml'], T, 'Types'),
      entries = [
        [targetPath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],
        [files.table, PIVOT_TYPES.table],
        [files.definition, PIVOT_TYPES.definition],
        [files.records, PIVOT_TYPES.records],
      ];
    replacements['[Content_Types].xml'] = append(
      types.raw,
      types.doc.documentElement,
      entries
        .map(([file, type]) => '<Override xmlns="' + T + '" PartName="/' + file + '" ContentType="' + type + '"/>')
        .join('')
    );
    replacements[relPath(targetPath)] = rels(
      relationship('rIdPivot', 'pivotTable', '../pivotTables/' + path.posix.basename(files.table))
    );
    replacements[relPath(files.table)] = rels(
      relationship('rIdCache', 'pivotCacheDefinition', '../pivotCache/' + path.posix.basename(files.definition))
    );
    replacements[relPath(files.definition)] = rels(
      relationship('rIdRecords', 'pivotCacheRecords', path.posix.basename(files.records))
    );
  }
  const normalized = layoutOnly
    ? {
        operation: 'setPivotLayout',
        sheet: targetSheet,
        name,
        sample,
        grandTotals,
        expectedPivotRevision: priorRevision,
      }
    : editing
      ? {
          operation: 'updatePivot',
          sheet: targetSheet,
          name,
          sourceSheet: model.sheet,
          range: model.range,
          rowField: model.rowField,
          valueField: model.valueField,
          aggregate: model.aggregate,
          expectedPivotRevision: priorRevision,
        }
      : refresh
        ? { operation: 'refreshPivot', sheet: targetSheet, name, expectedPivotRevision: priorRevision }
        : {
            operation: 'createPivot',
            sheet: sourceSheet,
            name,
            range: model.range,
            rowField: model.rowField,
            valueField: model.valueField,
            aggregate: model.aggregate,
            pivotSheet: targetSheet,
            expectedPivotRevision: null,
          };
  return {
    command: normalized,
    plan: {
      kind: command.operation,
      name,
      sourceSheet,
      sourcePath: source.path,
      sourceRange: model.range,
      rowField: model.rowField,
      valueField: model.valueField,
      aggregate: model.aggregate,
      targetSheet,
      targetPath,
      cacheId,
      sample,
      grandTotals,
      parts: files,
      priorRevision,
      replacements,
      removed: [],
      copies: {},
      fragments: {},
    },
  };
}
