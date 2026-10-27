import { FORMAT, FORMAT_REGEX, formatTemplate, stripFormatPrefix, formatPattern } from './format.mjs';
import { lineStarts } from './xml-position.mjs';
import { normalizeInkStroke, isInkStrokeCommand, inkPen } from './ink-stroke.mjs';
import { inkStrokeAnchorXml, readInkStrokeAnchor } from './ink-stroke-xml.mjs';
import { normalizeChartModel as chartModel, readChartData as chartData } from './chart-authoring-model.mjs';
import { CHART_STYLES, chartSeriesPropertiesXml } from './chart-styles.mjs';
import {
  MODEL3D_NS,
  MODEL3D_REL,
  MODEL3D_MIME,
  model3DEdgeId,
  model3DAnchorXml,
  model3DFromAnchor,
  isModel3DCommand,
} from './model3d-authoring.mjs';
import { MODEL3D_LIMITS, MODEL3D_SAMPLES, normalizeModel3D, model3DGlb, parseModel3DGlb } from './model3d-model.mjs';
/** Bounded managed chart/PNG package plans. Cells stay in the accepted owner. */
import path from 'node:path';
import {
  SMARTART_NS,
  smartartModelFromPackage,
  normalizeSmartArt,
  smartartPackage,
  smartartAnchorXml,
  smartartEdgeIds,
  isSmartArtCommand,
} from './smartart-model.mjs';
import { normalizeShape, shapeSample, shapeAnchorXml, shapeModelFromAnchor, isShapeCommand } from './shape-model.mjs';
import { createHash } from 'node:crypto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strFromU8, unzlibSync } from 'fflate';
import { openZipParts } from './zip-parts.mjs';
import { cellPosition, columnName, rangeBounds, assertXmlText } from './workbook-commands.mjs';
import { xmlElementEnd } from './worksheet-xml.mjs';
import { refusal } from './refusal-codes.mjs';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  A = 'http://schemas.openxmlformats.org/drawingml/2006/main',
  C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  P = 'http://schemas.openxmlformats.org/package/2006/relationships',
  T = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  CT_DRAW = 'application/vnd.openxmlformats-officedocument.drawing+xml';
export const INDEXED_OBJECT_LIMITS = Object.freeze({
  pngBytes: 128 * 1024,
  pngPixels: 1024 * 1024,
  sourcePositions: 1000,
  managedObjects: 32,
  charts: 16,
  inventoryObjects: 256,
  metadataBytes: 16 * 1024 * 1024,
  chartBytes: 4 * 1024 * 1024,
});
const uuid = (value) => typeof value === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(value),
  hash = (value) => createHash('sha256').update(value).digest('hex');
const elements = (node, local, ns) => Array.from(node.getElementsByTagNameNS(ns, local)),
  children = (node) => Array.from(node.childNodes).filter((item) => item.nodeType === 1);
const serialize = (node) => new XMLSerializer().serializeToString(node);
const quote = (value) => {
  assertXmlText(String(value));
  return String(value).replace(
    /[&<>"\r]/g,
    (char) =>
      ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        '\r': '&#13;',
      })[char]
  );
};
function text(value, limit, label) {
  if (typeof value !== 'string' || value.length > limit) throw refusal('internal-invalid-field', 'Invalid ' + label);
  assertXmlText(value);
  return value;
}
function integer(value, min, max, label) {
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function object(value, keys, label) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function parse(bytes, namespace, local) {
  if (typeof bytes !== 'string' && !(bytes instanceof Uint8Array))
    throw refusal('internal-missing-field', 'Object XML part is missing');
  const source = typeof bytes === 'string' ? bytes : strFromU8(bytes);
  if (Buffer.byteLength(source) > INDEXED_OBJECT_LIMITS.metadataBytes || /<!DOCTYPE|<!ENTITY/i.test(source))
    throw refusal('object-malformed', 'Object XML exceeds its bound or uses entities');
  const problems = [],
    doc = new DOMParser({
      locator: {},
      normalizeLineEndings: (value) => value,
      errorHandler: {
        warning: (message) => problems.push(message),
        error: (message) => problems.push(message),
        fatalError: (message) => problems.push(message),
      },
    }).parseFromString(source, 'application/xml');
  if (problems.length || doc.documentElement.namespaceURI !== namespace || doc.documentElement.localName !== local)
    throw refusal('object-malformed', 'Invalid object XML');
  return {
    source,
    doc,
  };
}
function rawChildren(source, doc) {
  const lines = lineStarts(source);
  return children(doc.documentElement).map((node) => {
    const start = lines[node.lineNumber - 1] + node.columnNumber - 1;
    if (!Number.isSafeInteger(start) || !source.startsWith('<' + node.tagName, start))
      throw refusal('object-unsupported', 'Object XML has no qualified raw child boundary');
    const end = xmlElementEnd(source, start);
    return {
      node,
      start,
      end,
      raw: source.slice(start, end),
    };
  });
}
function relPath(owner) {
  return path.posix.join(path.posix.dirname(owner), '_rels', path.posix.basename(owner) + '.rels');
}
function target(owner, value) {
  if (typeof value !== 'string' || !value || /[\\?#%\0]/.test(value))
    throw refusal('object-unsupported', 'Unqualified object relationship target');
  const resolved = value.startsWith('/')
    ? value.slice(1)
    : path.posix.normalize(path.posix.join(path.posix.dirname(owner), value));
  if (
    !resolved ||
    resolved.startsWith('/') ||
    resolved.split('/').some((segment) => !segment || segment === '.' || segment === '..')
  )
    throw refusal('object-unsupported', 'Object relationship escapes the package');
  return resolved;
}
function relationships(parts, owner) {
  const file = relPath(owner),
    raw = parts[file];
  if (!raw)
    return {
      file,
      doc: parse(`<Relationships xmlns="${P}"/>`, P, 'Relationships').doc,
      map: new Map(),
    };
  const { doc } = parse(raw, P, 'Relationships'),
    map = new Map();
  for (const node of children(doc.documentElement)) {
    const id = node.getAttribute('Id');
    if (node.namespaceURI !== P || node.localName !== 'Relationship' || !id || map.has(id))
      throw refusal('object-ambiguous', 'Ambiguous object relationships');
    const external = node.getAttribute('TargetMode') === 'External',
      to = external ? null : target(owner, node.getAttribute('Target'));
    if (!external && !Object.hasOwn(parts, to))
      throw refusal('internal-missing-field', 'Missing object relationship target');
    map.set(id, {
      id,
      node,
      type: node.getAttribute('Type'),
      external,
      target: to,
    });
  }
  return {
    file,
    doc,
    map,
  };
}
function one(node, local, ns) {
  const found = elements(node, local, ns);
  if (found.length !== 1) throw refusal('object-ambiguous', 'Ambiguous ' + local + ' object metadata');
  return found[0];
}
function numberText(node, local, ns, max) {
  return integer(Number(one(node, local, ns).textContent), 0, max, local);
}
function geometry(node) {
  if (node.namespaceURI !== D || node.localName !== 'oneCellAnchor')
    throw refusal('object-unsupported', 'Only ordinary one-cell authored anchors are qualified');
  const from = one(node, 'from', D),
    ext = one(node, 'ext', D);
  return {
    kind: 'oneCell',
    cell: columnName(numberText(from, 'col', D, 16383)) + (numberText(from, 'row', D, 1048575) + 1),
    dxEmu: numberText(from, 'colOff', D, 10000000),
    dyEmu: numberText(from, 'rowOff', D, 10000000),
    widthEmu: integer(Number(ext.getAttribute('cx')), 1, 50000000, 'width'),
    heightEmu: integer(Number(ext.getAttribute('cy')), 1, 50000000, 'height'),
  };
}
function anchor(value, defaults) {
  const source = {
    ...defaults,
    ...(value ?? {}),
  };
  object(source, ['kind', 'cell', 'dxEmu', 'dyEmu', 'widthEmu', 'heightEmu'], 'object anchor');
  if (source.kind !== 'oneCell') throw refusal('object-unsupported', 'Only one-cell placement is qualified');
  const position = cellPosition(text(source.cell, 12, 'anchor cell'));
  return {
    kind: 'oneCell',
    cell: columnName(position.col) + (position.row + 1),
    dxEmu: integer(source.dxEmu, 0, 10000000, 'horizontal offset'),
    dyEmu: integer(source.dyEmu, 0, 10000000, 'vertical offset'),
    widthEmu: integer(source.widthEmu, 1, 50000000, 'width'),
    heightEmu: integer(source.heightEmu, 1, 50000000, 'height'),
  };
}
const defaultAnchor = {
  kind: 'oneCell',
  cell: 'E2',
  dxEmu: 0,
  dyEmu: 0,
  widthEmu: 5715000,
  heightEmu: 3333750,
};
function formula(sheet, col, first, last = first) {
  return `'${sheet.replaceAll("'", "''")}'!$${columnName(col)}$${first + 1}${last === first ? '' : `:$${columnName(col)}$${last + 1}`}`;
}
function reference(value) {
  const match = /^'((?:[^']|'')+)'!\$([A-Z]{1,3})\$([1-9]\d*)(?::\$([A-Z]{1,3})\$([1-9]\d*))?$/.exec(value);
  if (!match) throw refusal('object-unsupported', 'Unqualified chart source reference');
  const sheet = match[1].replaceAll("''", "'"),
    a = cellPosition(match[2] + match[3]),
    b = cellPosition((match[4] ?? match[2]) + (match[5] ?? match[3]));
  return {
    sheet,
    first: a,
    last: b,
  };
}
function ordinaryAttributes(node, allowed) {
  for (const attribute of Array.from(node.attributes)) {
    if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
    const name =
      attribute.namespaceURI === R
        ? 'r:' + attribute.localName
        : attribute.namespaceURI
          ? '?' + attribute.localName
          : attribute.localName;
    if (!allowed.includes(name)) throw refusal('object-protected', 'Extended object attributes are protected');
  }
}
function managedAnchor(node) {
  const tags = new Map([
    [
      D,
      new Set([
        'oneCellAnchor',
        'from',
        'col',
        'colOff',
        'row',
        'rowOff',
        'ext',
        'graphicFrame',
        'nvGraphicFramePr',
        'cNvPr',
        'cNvGraphicFramePr',
        'xfrm',
        'pic',
        'nvPicPr',
        'cNvPicPr',
        'blipFill',
        'spPr',
        'clientData',
      ]),
    ],
    [A, new Set(['graphic', 'graphicData', 'blip', 'stretch', 'fillRect', 'prstGeom', 'avLst'])],
    [C, new Set(['chart'])],
  ]);
  for (const item of [node, ...elements(node, '*', '*')]) {
    if (!tags.get(item.namespaceURI)?.has(item.localName))
      throw refusal('object-protected', 'Extended drawing object is protected');
    ordinaryAttributes(
      item,
      item.namespaceURI === D && item.localName === 'cNvPr'
        ? ['id', 'name', 'descr']
        : item.namespaceURI === D && item.localName === 'ext'
          ? ['cx', 'cy']
          : item.namespaceURI === A && item.localName === 'graphicData'
            ? ['uri']
            : item.namespaceURI === A && item.localName === 'blip'
              ? ['r:embed']
              : item.namespaceURI === A && item.localName === 'prstGeom'
                ? ['prst']
                : item.namespaceURI === C && item.localName === 'chart'
                  ? ['r:id']
                  : []
    );
  }
}
function modelFromChart(bytes, sheet) {
  const { doc } = parse(bytes, C, 'chartSpace');
  const allowed = new Set([
    'chartSpace',
    'chart',
    'title',
    'tx',
    'rich',
    'plotArea',
    'layout',
    'barChart',
    'lineChart',
    'barDir',
    'grouping',
    'ser',
    'idx',
    'order',
    'strRef',
    'numRef',
    'f',
    'strCache',
    'numCache',
    'ptCount',
    'pt',
    'v',
    'cat',
    'val',
    'axId',
    'catAx',
    'valAx',
    'scaling',
    'orientation',
    'axPos',
    'crossAx',
    'legend',
    'legendPos',
    'plotVisOnly',
    'dispBlanksAs',
    'spPr',
  ]);
  for (const node of elements(doc, '*', '*')) {
    if (
      (node.namespaceURI === C && !allowed.has(node.localName)) ||
      (node.namespaceURI !== C &&
        (node.namespaceURI !== A ||
          !['bodyPr', 'lstStyle', 'p', 'r', 't', 'solidFill', 'srgbClr', 'noFill', 'ln', 'prstDash'].includes(
            node.localName
          )))
    )
      throw refusal('object-protected', 'Extended or opaque chart metadata is protected');
    ordinaryAttributes(
      node,
      node.namespaceURI === A && ['srgbClr', 'prstDash'].includes(node.localName)
        ? ['val']
        : node.namespaceURI === A && node.localName === 'ln'
          ? ['w']
          : node.namespaceURI === C && node.localName === 'pt'
            ? ['idx']
            : node.namespaceURI === C &&
                [
                  'barDir',
                  'grouping',
                  'idx',
                  'order',
                  'ptCount',
                  'axId',
                  'orientation',
                  'axPos',
                  'crossAx',
                  'legendPos',
                  'plotVisOnly',
                  'dispBlanksAs',
                ].includes(node.localName)
              ? ['val']
              : []
    );
  }
  const plots = elements(doc, 'barChart', C).concat(elements(doc, 'lineChart', C));
  if (plots.length !== 1) throw refusal('object-protected', 'Combo/alternative chart metadata is protected');
  const plot = plots[0],
    direction = plot.localName === 'barChart' ? one(plot, 'barDir', C).getAttribute('val') : null,
    kind = direction === 'bar' ? 'bar' : direction === 'col' ? 'column' : 'line';
  if (
    (direction !== null &&
      (!['col', 'bar'].includes(direction) || one(plot, 'grouping', C).getAttribute('val') !== 'clustered')) ||
    (direction === null && one(plot, 'grouping', C).getAttribute('val') !== 'standard')
  )
    throw refusal('object-protected', 'Alternative chart layout is protected');
  const series = elements(plot, 'ser', C);
  const appearance = elements(doc, 'spPr', C);
  let style;
  if (appearance.length) {
    const sig = (node) => [
      node.namespaceURI,
      node.localName,
      Array.from(node.attributes)
        .filter((a) => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
        .map((a) => [a.namespaceURI ?? '', a.localName, a.value])
        .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
      children(node).length ? (node.textContent.trim() === '' ? '' : '#mixed') : node.textContent,
      children(node).map(sig),
    ];
    if (appearance.length !== series.length)
      throw refusal('object-protected', 'Incomplete/foreign chart formatting is protected');
    for (const candidate of CHART_STYLES) {
      if (
        series.every((ser, index) => {
          const props = children(ser).filter((n) => n.namespaceURI === C && n.localName === 'spPr');
          if (props.length !== 1) return false;
          const expected = parse(
            `<c:spPr xmlns:c="${C}" xmlns:a="${A}">${chartSeriesPropertiesXml(candidate.id, index).replace(/^<c:spPr>|<\/c:spPr>$/g, '')}</c:spPr>`,
            C,
            'spPr'
          ).doc.documentElement;
          return JSON.stringify(sig(props[0])) === JSON.stringify(sig(expected));
        })
      ) {
        style = candidate.id;
        break;
      }
    }
    if (!style) throw refusal('object-protected', 'Foreign/extended chart style is protected');
  }

  if (!series.length || series.length > 8) throw refusal('object-unsupported', 'Chart series count is not qualified');
  const axes = elements(doc, 'catAx', C).concat(elements(doc, 'valAx', C));
  if (axes.length !== 2 || axes[0].localName !== 'catAx' || axes[1].localName !== 'valAx')
    throw refusal('object-unsupported', 'Chart axes are not qualified');
  const axisIds = axes.map((axis) => one(axis, 'axId', C).getAttribute('val'));
  if (
    axisIds[0] === axisIds[1] ||
    axisIds.some((id) => !/^[1-9]\d*$/.test(id)) ||
    elements(plot, 'axId', C)
      .map((node) => node.getAttribute('val'))
      .sort()
      .join(',') !== [...axisIds].sort().join(',')
  )
    throw refusal('internal-mismatch', 'Chart axis identities disagree');
  axes.forEach((axis, index) => {
    if (
      one(axis, 'crossAx', C).getAttribute('val') !== axisIds[1 - index] ||
      one(axis, 'orientation', C).getAttribute('val') !== 'minMax' ||
      one(axis, 'axPos', C).getAttribute('val') !== (kind === 'bar' ? (index ? 'b' : 'l') : index ? 'l' : 'b')
    )
      throw refusal('object-protected', 'Alternative chart axes are protected');
  });
  let first, last, categoryColumn;
  series.forEach((ser, index) => {
    if (
      Number(one(ser, 'idx', C).getAttribute('val')) !== index ||
      Number(one(ser, 'order', C).getAttribute('val')) !== index
    )
      throw refusal('object-unsupported', 'Chart series ordering is not qualified');
    const cat = reference(one(one(ser, 'cat', C), 'f', C).textContent),
      val = reference(one(one(ser, 'val', C), 'f', C).textContent),
      header = reference(one(one(ser, 'tx', C), 'f', C).textContent);
    if (
      cat.sheet !== sheet ||
      val.sheet !== sheet ||
      header.sheet !== sheet ||
      cat.first.col !== cat.last.col ||
      val.first.col !== val.last.col ||
      header.first.row !== header.last.row ||
      header.first.col !== header.last.col ||
      header.first.col !== val.first.col ||
      header.first.row !== val.first.row - 1 ||
      cat.first.row !== val.first.row ||
      cat.last.row !== val.last.row ||
      val.first.col !== cat.first.col + index + 1
    )
      throw refusal('object-unsupported', 'Chart source mapping is not qualified');
    if (index && (cat.first.col !== categoryColumn || val.first.row !== first || val.last.row !== last))
      throw refusal('internal-mismatch', 'Chart series ranges disagree');
    first = val.first.row;
    last = val.last.row;
    categoryColumn = cat.first.col;
  });
  const title = elements(doc, 'title', C)[0],
    legend = elements(doc, 'legendPos', C)[0];
  if (elements(doc, 'title', C).length > 1 || elements(doc, 'legend', C).length > 1)
    throw refusal('object-ambiguous', 'Ambiguous chart title/legend is protected');
  if (
    one(doc, 'plotVisOnly', C).getAttribute('val') !== '1' ||
    one(doc, 'dispBlanksAs', C).getAttribute('val') !== 'gap'
  )
    throw refusal('object-protected', 'Alternative chart display policy is protected');
  return chartModel({
    ...(style ? { style } : {}),
    kind,
    range: `${columnName(categoryColumn)}${first}:${columnName(categoryColumn + series.length)}${last + 1}`,
    title: title
      ? elements(title, 't', A)
          .map((node) => node.textContent)
          .join('')
      : '',
    legend: legend ? (legend.getAttribute('val') === 'b' ? 'bottom' : 'unsupported') : 'none',
    blanks: 'gap',
    includeHidden: false,
  });
}
function typeFor(doc, part) {
  let result;
  const seen = new Set();
  for (const node of children(doc.documentElement)) {
    if (node.namespaceURI !== T || !['Override', 'Default'].includes(node.localName))
      throw refusal('internal-unknown-reference', 'Unknown object content-type entry');
    const override = node.localName === 'Override',
      key = override ? node.getAttribute('PartName') : node.getAttribute('Extension').toLowerCase(),
      identity = node.localName + ':' + key;
    if (!key || seen.has(identity)) throw refusal('object-ambiguous', 'Ambiguous object content types');
    seen.add(identity);
    if (override && key === '/' + part) result = node.getAttribute('ContentType');
    else if (!override && key === part.split('.').at(-1).toLowerCase() && result === undefined)
      result = node.getAttribute('ContentType');
  }
  return result;
}
function context(metadataBytes, view) {
  if (
    !(metadataBytes instanceof Uint8Array) ||
    metadataBytes.byteLength > INDEXED_OBJECT_LIMITS.metadataBytes ||
    !Array.isArray(view?.sheets)
  )
    throw refusal('object-limit', 'Object metadata exceeds its bound');
  const parts = openZipParts(metadataBytes, INDEXED_OBJECT_LIMITS.metadataBytes),
    types = parse(parts['[Content_Types].xml'], T, 'Types').doc,
    canvases = new Map(),
    all = [];
  let managed = 0;
  const seen = new Set();
  for (const sheet of view.sheets) {
    const { doc, source } = parse(parts[sheet.path], S, 'worksheet'),
      drawing = elements(doc, 'drawing', S);
    if (drawing.length > 1) throw refusal('object-protected', 'Shared/ambiguous worksheet drawings are protected');
    const sheetRels = relationships(parts, sheet.path),
      canvas = {
        sheet,
        doc,
        source,
        sheetRels,
        objects: [],
      };
    canvases.set(sheet.name, canvas);
    if (!drawing.length) continue;
    const edge = sheetRels.map.get(drawing[0].getAttributeNS(R, 'id'));
    if (!edge || edge.external || edge.type !== R + '/drawing' || typeFor(types, edge.target) !== CT_DRAW)
      throw refusal('object-protected', 'Worksheet drawing ownership is protected');
    canvas.drawingPart = edge.target;
    canvas.sheetEdge = edge;
    const parsed = parse(parts[edge.target], D, 'wsDr');
    Object.assign(canvas, {
      drawingDoc: parsed.doc,
      drawingSource: parsed.source,
      drawingRels: relationships(parts, edge.target),
    });
    const nonvisualIds = new Set();
    for (const item of rawChildren(parsed.source, parsed.doc)) {
      if (all.length >= INDEXED_OBJECT_LIMITS.inventoryObjects)
        throw refusal('object-limit', 'Object inventory exceeds its 256-object bound');
      const properties = elements(item.node, 'cNvPr', D);
      if (properties.length !== 1) throw refusal('object-ambiguous', 'Ambiguous drawing nonvisual metadata');
      const props = properties[0],
        name = props.getAttribute('name'),
        id = uuid(stripFormatPrefix(name, 'FORMAT_183450'))
          ? stripFormatPrefix(name, 'FORMAT_183450')
          : 'opaque-' + hash(sheet.path + '\0' + edge.target + '\0' + props.getAttribute('id')).slice(0, 24),
        entry = {
          id,
          name: text(name, 256, 'object name'),
          canvas,
          item,
          properties: props,
          kind: 'opaque',
          protectedReason: 'Unmanaged drawing object is preserved',
          managed: uuid(id),
        };
      const nonvisualId = props.getAttribute('id');
      if (!/^\d+$/.test(nonvisualId) || nonvisualIds.has(nonvisualId))
        throw refusal('object-refused', 'Drawing nonvisual identity is duplicated/invalid');
      nonvisualIds.add(nonvisualId);
      if (seen.has(id)) throw refusal('object-refused', 'Object identity is duplicated');
      seen.add(id);
      if (entry.managed) managed++;
      const chart = elements(item.node, 'chart', C),
        pic = elements(item.node, 'pic', D),
        shape = elements(item.node, 'sp', D),
        diagram = elements(item.node, 'relIds', SMARTART_NS),
        model3d = elements(item.node, 'model3d', MODEL3D_NS);
      entry.kind = chart.length
        ? 'chart'
        : pic.length
          ? 'image'
          : shape.length
            ? 'shape'
            : diagram.length
              ? 'smartart'
              : model3d.length
                ? 'model3d'
                : 'opaque';
      try {
        entry.anchor = geometry(item.node);
        if (entry.managed && model3d.length === 1 && !diagram.length && !chart.length && !pic.length && !shape.length) {
          entry.model3d = model3DFromAnchor(item.node, entry.anchor, entry.id);
          entry.edge = canvas.drawingRels.map.get(model3DEdgeId(entry.id, FORMAT));
          if (!entry.edge || entry.edge.external || entry.edge.type !== MODEL3D_REL)
            throw refusal('internal-unknown-reference', 'Unknown/external 3D relationship');
          entry.part = entry.edge.target;
          if (
            !new RegExp(formatPattern('XL_MEDIA_MODEL3D') + entry.id + '-[0-9a-f]{64}\\.glb$').test(entry.part) ||
            typeFor(types, entry.part) !== MODEL3D_MIME ||
            Object.hasOwn(parts, relPath(entry.part))
          )
            throw refusal('object-protected', 'Foreign/dependent 3D asset is protected');
          entry.protectedReason = null;
        } else if (entry.managed && diagram.length === 1 && !chart.length && !pic.length && !shape.length) {
          const graph = smartartModelFromPackage(item.node, parts, canvas.drawingPart, entry.anchor, entry.id);
          entry.smartart = graph.model;
          entry.smartartParts = graph.parts;
          entry.protectedReason = null;
        } else if (entry.managed && shape.length === 1 && !chart.length && !pic.length) {
          try {
            const ink = readInkStrokeAnchor(serialize(item.node));
            entry.stroke = ink.stroke;
            entry.kind = 'ink';
          } catch {
            entry.shape = shapeModelFromAnchor(item.node, entry.anchor);
          }
          entry.protectedReason = null;
        } else if (entry.managed) managedAnchor(item.node);
        if (chart.length === 1 && !pic.length) {
          entry.kind = 'chart';
          entry.edge = canvas.drawingRels.map.get(chart[0].getAttributeNS(R, 'id'));
          if (!entry.edge || entry.edge.external || entry.edge.type !== R + '/chart')
            throw refusal('object-protected', 'External/unknown chart relationship is protected');
          entry.part = entry.edge.target;
          if (entry.managed) {
            if (
              stripFormatPrefix(entry.part, 'XL_CHARTS_CHART') !== entry.id + '.xml' ||
              typeFor(types, entry.part) !== CT_CHART ||
              Object.hasOwn(parts, relPath(entry.part))
            )
              throw refusal('object-protected', 'Unmanaged/dependent chart graph is protected');
            entry.chart = modelFromChart(parts[entry.part], sheet.name);
            entry.protectedReason = null;
          }
        } else if (pic.length === 1 && !chart.length) {
          entry.kind = 'image';
          const image = one(pic[0], 'blip', A);
          entry.edge = canvas.drawingRels.map.get(image.getAttributeNS(R, 'embed'));
          if (!entry.edge || entry.edge.external || entry.edge.type !== R + '/image')
            throw refusal('object-protected', 'External/unknown picture relationship is protected');
          entry.part = entry.edge.target;
          if (entry.managed) {
            if (
              !FORMAT_REGEX.IMAGE_PART.test(entry.part) ||
              typeFor(types, entry.part) !== 'image/png' ||
              elements(item.node, '*', '*').some((node) =>
                ['extLst', 'srcRect', 'effectLst', 'effectDag', 'hlinkClick', 'hlinkHover'].includes(node.localName)
              )
            )
              throw refusal('object-protected', 'Unmanaged/cropped/extended picture is protected');
            entry.protectedReason = null;
          }
        }
        if (!entry.managed) entry.protectedReason = 'Unmanaged drawing object is preserved';
      } catch (error) {
        delete entry.model3d;
        entry.protectedReason = String(error.message ?? error);
      }
      entry.revision = hash(
        item.raw +
          '\0' +
          (entry.part ?? '') +
          '\0' +
          (entry.kind === 'chart' && entry.part ? strFromU8(parts[entry.part]) : '') +
          '\0' +
          (entry.edge ? serialize(entry.edge.node) : '')
      );
      if (entry.model3d && entry.part)
        entry.revision = hash(
          item.raw +
            '\0' +
            entry.part +
            '\0' +
            entry.edge.id +
            '\0' +
            entry.edge.type +
            '\0' +
            entry.edge.node.getAttribute('Target')
        );
      if (entry.smartartParts)
        entry.revision = hash(
          item.raw +
            '\0' +
            Object.keys(entry.smartartParts)
              .sort()
              .map((path) => path + '\0' + strFromU8(parts[path]))
              .join('\0')
        );
      canvas.objects.push(entry);
      all.push(entry);
    }
  }
  if (managed > INDEXED_OBJECT_LIMITS.managedObjects)
    throw refusal('object-limit', 'Managed objects exceed their 32-object bound');
  const ctx = {
    parts,
    types,
    canvases,
    all,
    managed,
  };
  for (const entry of all)
    if (
      !entry.protectedReason &&
      (incoming(ctx, entry.canvas.drawingPart) !== 1 ||
        (['chart', 'model3d'].includes(entry.kind) && incoming(ctx, entry.part) !== 1))
    )
      entry.protectedReason = 'Shared mutable object ownership is protected';
  for (const entry of all)
    if (entry.model3d && (incoming(ctx, entry.canvas.drawingPart) !== 1 || incoming(ctx, entry.part) !== 1)) {
      delete entry.model3d;
      entry.protectedReason = 'Shared 3D asset/drawing is protected';
    }
  return ctx;
}
function summary(entry) {
  return {
    id: entry.id,
    name:
      (entry.stroke ? inkPen(entry.stroke.pen).label + ' stroke' : '') ||
      (entry.model3d ? MODEL3D_SAMPLES.find((s) => s.id === entry.model3d.sample).label : '') ||
      entry.smartart?.nodes[0]?.text ||
      entry.shape?.text.trim().split('\n')[0].slice(0, 128) ||
      (entry.shape ? shapeSample(entry.shape.sample).label : '') ||
      entry.chart?.title ||
      (entry.kind === 'image' ? entry.properties.getAttribute('descr') : '') ||
      (entry.managed ? (entry.kind === 'chart' ? 'Chart' : 'Image') : entry.name),
    kind: entry.kind,
    revision: entry.revision,
    anchor: entry.anchor ?? null,
    managed: entry.managed,
    protectedReason: entry.protectedReason,
    canEdit: ['chart', 'image', 'shape', 'smartart', 'model3d', 'ink'].includes(entry.kind) && !entry.protectedReason,
    canDelete: !entry.protectedReason,
    ...(entry.model3d ? { model3d: entry.model3d } : {}),
    ...(entry.smartart ? { smartart: entry.smartart } : {}),
    ...(entry.stroke ? { stroke: entry.stroke } : {}),
    ...(entry.shape ? { shape: entry.shape } : {}),
    ...(['chart', 'image', 'shape', 'smartart', 'model3d', 'ink'].includes(entry.kind)
      ? { altText: entry.properties.getAttribute('descr') || '' }
      : {}),
    ...(entry.chart
      ? {
          chart: entry.chart,
        }
      : {}),
  };
}
export function isIndexedObjectCommand(command) {
  return (
    isInkStrokeCommand(command) ||
    ['createChart', 'updateChart', 'insertImage', 'updateImage', 'deleteObject'].includes(command?.operation) ||
    isShapeCommand(command) ||
    isSmartArtCommand(command) ||
    isModel3DCommand(command)
  );
}
/** One bounded metadata pass for read-only workbook-wide content inspection. */
export function indexedAccessibilityObjects(metadataBytes, view) {
  return context(metadataBytes, view).all.map((entry) => ({ sheet: entry.canvas.sheet.name, ...summary(entry) }));
}
export function indexedObjectList(metadataBytes, view, sheet, { after, limit = 50 } = {}) {
  integer(limit, 1, 50, 'object page limit');
  const ctx = context(metadataBytes, view);
  if (!ctx.canvases.has(sheet)) throw refusal('internal-unknown-reference', 'Unknown worksheet');
  const entries = ctx.all.filter((item) => item.canvas.sheet.name === sheet).sort((a, b) => a.id.localeCompare(b.id));
  let offset = 0;
  if (after !== undefined) {
    offset = entries.findIndex((item) => item.id === after) + 1;
    if (!offset) throw refusal('object-stale', 'Object cursor expired; refresh the list');
  }
  const objects = entries.slice(offset, offset + limit).map(summary);
  return {
    objects,
    total: entries.length,
    nextCursor: offset + objects.length < entries.length ? objects.at(-1).id : null,
  };
}
function cacheXml(kind, formulaText, values) {
  const points = values
    .map((value, index) =>
      value === null ? '' : `<c:pt idx="${index}"><c:v>${quote(Object.is(value, -0) ? '-0' : value)}</c:v></c:pt>`
    )
    .join('');
  return `<c:${kind}Ref><c:f>${quote(formulaText)}</c:f><c:${kind}Cache><c:ptCount val="${values.length}"/>${points}</c:${kind}Cache></c:${kind}Ref>`;
}
function chartXml(sheet, model, data) {
  const bounds = rangeBounds(model.range),
    series = data.series
      .map(
        (item, index) =>
          `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx>${cacheXml('str', formula(sheet, bounds.first.col + index + 1, bounds.first.row), [item.name])}</c:tx>${chartSeriesPropertiesXml(model.style, index)}<c:cat>${cacheXml('str', formula(sheet, bounds.first.col, bounds.first.row + 1, bounds.last.row), data.categories)}</c:cat><c:val>${cacheXml('num', formula(sheet, bounds.first.col + index + 1, bounds.first.row + 1, bounds.last.row), item.values)}</c:val></c:ser>`
      )
      .join(''),
    tag = model.kind === 'line' ? 'lineChart' : 'barChart',
    options =
      model.kind === 'line'
        ? '<c:grouping val="standard"/>'
        : `<c:barDir val="${model.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/>`;
  return `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart>${model.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${quote(model.title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : ''}<c:plotArea><c:layout/><c:${tag}>${options}${series}<c:axId val="10"/><c:axId val="20"/></c:${tag}><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="${model.kind === 'bar' ? 'l' : 'b'}"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="${model.kind === 'bar' ? 'b' : 'l'}"/><c:crossAx val="10"/></c:valAx></c:plotArea>${model.legend === 'bottom' ? '<c:legend><c:legendPos val="b"/></c:legend>' : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}
const crcTable = Uint32Array.from(
  {
    length: 256,
  },
  (_, value) => {
    for (let i = 0; i < 8; i++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    return value >>> 0;
  }
);
function crc(bytes) {
  let value = 0xffffffff;
  for (const byte of bytes) value = crcTable[(value ^ byte) & 255] ^ (value >>> 8);
  return (value ^ 0xffffffff) >>> 0;
}
export function indexedPngPayload(asset) {
  object(asset, ['mime', 'base64', 'sha256', 'bytes', 'widthPx', 'heightPx'], 'PNG payload');
  if (
    asset.mime !== 'image/png' ||
    typeof asset.base64 !== 'string' ||
    asset.base64.length > Math.ceil(INDEXED_OBJECT_LIMITS.pngBytes / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(asset.base64)
  )
    throw refusal('object-bad-input', 'PNG needs canonical bounded base64');
  const bytes = Buffer.from(asset.base64, 'base64');
  if (
    !bytes.length ||
    bytes.length > INDEXED_OBJECT_LIMITS.pngBytes ||
    bytes.toString('base64') !== asset.base64 ||
    !bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
  )
    throw refusal('internal-invalid-field', 'Invalid PNG bytes');
  let at = 8,
    width,
    height,
    channels,
    ended = false;
  const idat = [];
  while (at < bytes.length) {
    if (at + 12 > bytes.length) throw refusal('object-malformed', 'Truncated PNG chunk');
    const size = bytes.readUInt32BE(at),
      type = bytes.toString('ascii', at + 4, at + 8),
      end = at + 12 + size;
    if (end > bytes.length || crc(bytes.subarray(at + 4, at + 8 + size)) !== bytes.readUInt32BE(at + 8 + size))
      throw refusal('internal-invalid-field', 'Invalid PNG chunk bounds/CRC');
    if (
      !bytes.subarray(at + 4, at + 8).every((value) => (value >= 65 && value <= 90) || (value >= 97 && value <= 122)) ||
      ['iCCP', 'zTXt', 'iTXt'].includes(type)
    )
      throw refusal('object-unsupported', 'Compressed/unknown PNG metadata is not qualified');
    const data = bytes.subarray(at + 8, at + 8 + size);
    if ((at === 8 && type !== 'IHDR') || type === 'acTL' || type === 'fcTL' || type === 'fdAT')
      throw refusal('object-unsupported', 'Only static PNG is qualified');
    if (type === 'IHDR') {
      if (width !== undefined || size !== 13) throw refusal('internal-invalid-field', 'Invalid PNG header');
      width = integer(data.readUInt32BE(0), 1, 4096, 'PNG width');
      height = integer(data.readUInt32BE(4), 1, 4096, 'PNG height');
      if (
        width * height > INDEXED_OBJECT_LIMITS.pngPixels ||
        data[8] !== 8 ||
        ![2, 6].includes(data[9]) ||
        data[10] ||
        data[11] ||
        data[12]
      )
        throw refusal('object-bad-input', 'PNG requires bounded noninterlaced 8-bit RGB/RGBA');
      channels = data[9] === 2 ? 3 : 4;
    } else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') {
      if (size || end !== bytes.length) throw refusal('internal-invalid-field', 'Invalid PNG end');
      ended = true;
    } else if (type[0] === type[0].toUpperCase() && type !== 'PLTE')
      throw refusal('internal-unknown-reference', 'Unknown critical PNG chunk');
    at = end;
  }
  if (!ended || !idat.length) throw refusal('object-bad-input', 'PNG has no complete pixel stream');
  const stream = Buffer.concat(idat);
  if (
    stream.length < 6 ||
    (stream[0] & 15) !== 8 ||
    stream[0] >> 4 > 7 ||
    ((stream[0] << 8) | stream[1]) % 31 ||
    stream[1] & 32
  )
    throw refusal('internal-invalid-field', 'Invalid PNG zlib stream');
  const expected = height * (1 + width * channels),
    decoded = unzlibSync(stream, {
      out: new Uint8Array(expected + 1),
    });
  if (decoded.length !== expected) throw refusal('internal-mismatch', 'PNG pixel stream disagrees with dimensions');
  let first = 1,
    second = 0;
  for (const value of decoded) {
    first = (first + value) % 65521;
    second = (second + first) % 65521;
  }
  if (((second << 16) | first) >>> 0 !== stream.readUInt32BE(stream.length - 4))
    throw refusal('internal-invalid-field', 'PNG pixel checksum is invalid');
  for (let row = 0; row < height; row++)
    if (decoded[row * (1 + width * channels)] > 4) throw refusal('internal-invalid-field', 'Invalid PNG row filter');
  const sha256 = hash(bytes);
  if (
    (asset.sha256 !== undefined && asset.sha256 !== sha256) ||
    (asset.bytes !== undefined && asset.bytes !== bytes.length) ||
    (asset.widthPx !== undefined && asset.widthPx !== width) ||
    (asset.heightPx !== undefined && asset.heightPx !== height)
  )
    throw refusal('internal-mismatch', 'PNG metadata/hash disagrees with its bytes');
  return {
    mime: 'image/png',
    base64: bytes.toString('base64'),
    sha256,
    bytes: bytes.length,
    widthPx: width,
    heightPx: height,
  };
}
function anchorXml(id, number, kind, relid, value, altText) {
  const p = cellPosition(value.cell),
    properties = formatTemplate(
      [FORMAT.XML_3837B8BF2F, FORMAT.NAME, FORMAT.DESCR, FORMAT.XML_5AB2E39E37],
      [number, id, quote(altText)]
    ),
    body =
      kind === 'chart'
        ? `<xdr:graphicFrame><xdr:nvGraphicFramePr>${properties}<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${C}"><c:chart r:id="${relid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
        : `<xdr:pic><xdr:nvPicPr>${properties}<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${relid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
  return `<xdr:oneCellAnchor xmlns:xdr="${D}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"><xdr:from><xdr:col>${p.col}</xdr:col><xdr:colOff>${value.dxEmu}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${value.dyEmu}</xdr:rowOff></xdr:from><xdr:ext cx="${value.widthEmu}" cy="${value.heightEmu}"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
}
function addType(doc, part, mime) {
  const entries = children(doc.documentElement),
    old = entries.filter((node) => node.localName === 'Override' && node.getAttribute('PartName') === '/' + part);
  if (old.length) throw refusal('object-conflict', 'New object content type collides');
  const node = doc.createElementNS(T, 'Override');
  node.setAttribute('PartName', '/' + part);
  node.setAttribute('ContentType', mime);
  doc.documentElement.appendChild(node);
}
function removeType(doc, part) {
  let changed = false;
  for (const node of children(doc.documentElement))
    if (node.localName === 'Override' && node.getAttribute('PartName') === '/' + part) {
      doc.documentElement.removeChild(node);
      changed = true;
    }
  return changed;
}
function addEdge(doc, id, kind, to) {
  if (children(doc.documentElement).some((node) => node.getAttribute('Id') === id))
    throw refusal('object-conflict', 'Object relationship ID collides');
  const node = doc.createElementNS(P, 'Relationship');
  for (const [key, value] of Object.entries({
    Id: id,
    Type: R + '/' + kind,
    Target: to,
  }))
    node.setAttribute(key, value);
  doc.documentElement.appendChild(node);
}
function incoming(ctx, part) {
  let count = 0;
  for (const file of Object.keys(ctx.parts))
    if (file.endsWith('.rels') && file !== '_rels/.rels') {
      const owner = file.replace('/_rels/', '/').slice(0, -5);
      for (const rel of relationships(ctx.parts, owner).map.values()) if (!rel.external && rel.target === part) count++;
    }
  return count;
}
function planBase(command) {
  return {
    kind: command.operation,
    sheet: null,
    objectId: command.objectId,
    replacements: {},
    removed: [],
    copies: {},
    fragments: {},
    binaryAdditions: {},
  };
}
function appendRaw(source, doc, value) {
  const closing = source.lastIndexOf('</' + doc.documentElement.tagName);
  if (closing < 0) {
    const opening = source.lastIndexOf('/>');
    if (opening < 0) throw refusal('internal-invalid-field', 'Invalid empty drawing');
    return (
      source.slice(0, opening) + '>' + value + '</' + doc.documentElement.tagName + '>' + source.slice(opening + 2)
    );
  }
  return source.slice(0, closing) + value + source.slice(closing);
}
export async function indexedObjectPlan({ metadataBytes, view, readRange, readMedia }, command) {
  object(
    command,
    [
      'operation',
      'sheet',
      'objectId',
      'expectedObjectRevision',
      'chart',
      'patch',
      'anchor',
      'asset',
      'altText',
      'shape',
      'smartart',
      'model3d',
      'stroke',
    ],
    'object command'
  );
  if (!isIndexedObjectCommand(command) || !uuid(command.objectId) || view.readOnly)
    throw refusal('object-unsupported', 'Unsupported object command/identity');
  const fields = {
    createChart: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'chart', 'anchor'],
    updateChart: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'patch', 'anchor'],
    insertImage: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'asset', 'anchor', 'altText'],
    updateImage: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'anchor', 'altText'],
    createModel3D: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'model3d', 'anchor', 'altText'],
    updateModel3D: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'patch', 'anchor', 'altText'],
    deleteModel3D: ['operation', 'sheet', 'objectId', 'expectedObjectRevision'],
    createSmartArt: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'smartart', 'anchor', 'altText'],
    updateSmartArt: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'patch', 'anchor', 'altText'],
    deleteSmartArt: ['operation', 'sheet', 'objectId', 'expectedObjectRevision'],
    createInkStroke: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'stroke', 'anchor', 'altText'],
    updateInkStroke: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'anchor', 'altText'],
    deleteInkStroke: ['operation', 'sheet', 'objectId', 'expectedObjectRevision'],
    createShape: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'shape', 'anchor', 'altText'],
    updateShape: ['operation', 'sheet', 'objectId', 'expectedObjectRevision', 'patch', 'anchor', 'altText'],
    deleteShape: ['operation', 'sheet', 'objectId', 'expectedObjectRevision'],
    deleteObject: ['operation', 'sheet', 'objectId', 'expectedObjectRevision'],
  };
  object(command, fields[command.operation], 'object command fields');
  const ctx = context(metadataBytes, view),
    canvas = ctx.canvases.get(command.sheet);
  if (!canvas) throw refusal('internal-unknown-reference', 'Unknown worksheet');
  if (isInkStrokeCommand(command) && (canvas.sheet.freeze?.rows || canvas.sheet.freeze?.columns))
    throw refusal('object-unsupported', 'Stroke authoring across frozen panes is outside the first profile');
  const create = [
      'createChart',
      'insertImage',
      'createShape',
      'createSmartArt',
      'createModel3D',
      'createInkStroke',
    ].includes(command.operation),
    existing = ctx.all.find((item) => item.id === command.objectId);
  if (create) {
    if (command.expectedObjectRevision !== null || existing)
      throw refusal('object-bad-input', 'Creating an object requires an absent identity guard');
    if (ctx.managed >= INDEXED_OBJECT_LIMITS.managedObjects)
      throw refusal('object-limit', 'Managed object limit reached');
  } else if (
    !existing ||
    existing.canvas !== canvas ||
    (existing.protectedReason &&
      !((isSmartArtCommand(command) && existing.smartart) || (isModel3DCommand(command) && existing.model3d))) ||
    command.expectedObjectRevision !== existing.revision
  )
    throw refusal('object-stale', 'Object changed, belongs to another sheet, or is protected');
  if (canvas.drawingPart && incoming(ctx, canvas.drawingPart) !== 1)
    throw refusal('object-protected', 'Shared mutable drawing is protected');
  if (command.operation === 'updateChart' && existing?.kind !== 'chart')
    throw refusal('object-bad-input', 'Select a managed chart before changing chart fields');
  if (command.operation === 'updateImage' && existing?.kind !== 'image')
    throw refusal('object-bad-input', 'Select a managed PNG before changing its placement');
  if (command.operation === 'deleteObject' && existing?.model3d) command = { ...command, operation: 'deleteModel3D' };
  if (isModel3DCommand(command) && !create) {
    if (!existing.model3d) throw refusal('object-bad-input', 'Choose an admitted 3D sample');
    await validatedModel3D(existing, readMedia);
  }
  if (isSmartArtCommand(command) && !create && !existing.smartart)
    throw refusal('object-bad-input', 'Choose an admitted SmartArt sample');
  if (isInkStrokeCommand(command) && !create && (!existing.stroke || existing.kind !== 'ink'))
    throw refusal('object-unsupported', 'Choose a qualified owned stroke');
  if (isShapeCommand(command) && !create && existing?.kind !== 'shape')
    throw refusal('object-bad-input', 'Choose an owned text/shape object');
  if (command.operation === 'deleteObject' && existing?.smartart) command = { ...command, operation: 'deleteSmartArt' };
  if (command.operation === 'deleteObject' && existing?.stroke) command = { ...command, operation: 'deleteInkStroke' };
  if (command.operation === 'deleteObject' && existing?.kind === 'shape')
    command = { ...command, operation: 'deleteShape' };
  const plan = planBase(command);
  plan.priorObjectRevision = command.expectedObjectRevision;
  let typesChanged = false;
  plan.sheet = canvas.sheet.path;
  let drawingPart = canvas.drawingPart,
    drawingDoc = canvas.drawingDoc,
    drawingSource = canvas.drawingSource,
    drawingRels = canvas.drawingRels;
  if (!drawingPart) {
    if (!create) throw refusal('internal-missing-field', 'Object drawing is missing');
    drawingPart = formatTemplate([FORMAT.XL_DRAWINGS_DRAWING, FORMAT.XML], [command.objectId]);
    if (Object.hasOwn(ctx.parts, drawingPart)) throw refusal('object-conflict', 'Object drawing part collides');
    const parsed = parse(`<xdr:wsDr xmlns:xdr="${D}"/>`, D, 'wsDr');
    drawingDoc = parsed.doc;
    drawingSource = parsed.source;
    drawingRels = relationships(ctx.parts, drawingPart);
    const rid = FORMAT.DRAWING + command.objectId.replaceAll('-', '');
    addEdge(
      canvas.sheetRels.doc,
      rid,
      'drawing',
      path.posix.relative(path.posix.dirname(canvas.sheet.path), drawingPart)
    );
    plan.replacements[canvas.sheetRels.file] = serialize(canvas.sheetRels.doc);
    plan.fragments[canvas.sheet.path] = {
      drawing: `<drawing xmlns="${S}" xmlns:r="${R}" r:id="${rid}"/>`,
    };
    addType(ctx.types, drawingPart, CT_DRAW);
    typesChanged = true;
  }
  plan.drawingPart = drawingPart;
  const existingFormat = FORMAT;
  let normalized;
  if (
    ['deleteObject', 'deleteShape', 'deleteSmartArt', 'deleteModel3D', 'deleteInkStroke'].includes(command.operation)
  ) {
    if (existing.kind === 'chart' && incoming(ctx, existing.part) !== 1)
      throw refusal('object-protected', 'Shared mutable chart is protected');
    drawingSource = drawingSource.slice(0, existing.item.start) + drawingSource.slice(existing.item.end);
    if (existing.smartart) {
      for (const part of Object.keys(existing.smartartParts)) {
        plan.removed.push(part);
        typesChanged = removeType(ctx.types, part) || typesChanged;
      }
      for (const rid of Object.values(smartartEdgeIds(existing.id, existingFormat))) {
        const edge = drawingRels.map.get(rid);
        drawingRels.doc.documentElement.removeChild(edge.node);
      }
    }
    if (existing.edge) drawingRels.doc.documentElement.removeChild(existing.edge.node);
    if (existing.part && incoming(ctx, existing.part) === 1) {
      plan.removed.push(existing.part);
      typesChanged = removeType(ctx.types, existing.part) || typesChanged;
    }
    if (canvas.objects.length === 1) {
      plan.removed.push(drawingPart);
      if (Object.hasOwn(ctx.parts, drawingRels.file)) plan.removed.push(drawingRels.file);
      canvas.sheetRels.doc.documentElement.removeChild(canvas.sheetEdge.node);
      plan.replacements[canvas.sheetRels.file] = serialize(canvas.sheetRels.doc);
      plan.fragments[canvas.sheet.path] = {
        drawing: '',
      };
      typesChanged = removeType(ctx.types, drawingPart) || typesChanged;
    } else {
      plan.replacements[drawingPart] = drawingSource;
      if (existing.edge || existing.smartart) plan.replacements[drawingRels.file] = serialize(drawingRels.doc);
    }
    plan.objectPart = existing.smartart
      ? smartartPackage(existing.id, existing.smartart, existing.anchor, existingFormat).names.data
      : ['shape', 'ink'].includes(existing.kind)
        ? drawingPart
        : existing.part;
    normalized = {
      operation: command.operation,
      sheet: command.sheet,
      objectId: command.objectId,
      expectedObjectRevision: command.expectedObjectRevision,
    };
  } else if (isModel3DCommand(command)) {
    if (command.patch !== undefined) object(command.patch, ['sample', 'color', 'rotation'], '3D patch');
    const model = normalizeModel3D(existing ? { ...existing.model3d, ...(command.patch ?? {}) } : command.model3d),
      placement = anchor(
        command.anchor,
        existing?.anchor ?? { ...defaultAnchor, widthEmu: 3048000, heightEmu: 2286000 }
      ),
      description = text(
        command.altText ?? existing?.properties.getAttribute('descr') ?? '',
        256,
        '3D alternative text'
      );
    const number = existing
      ? integer(Number(existing.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
      : Math.max(
          0,
          ...canvas.objects.map((item) =>
            integer(Number(item.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
          )
        ) + 1;
    integer(number, 1, 2147483647, 'drawing ID');
    const bytes = model3DGlb(model),
      sha256 = hash(bytes),
      part = formatTemplate(
        [FORMAT.XL_MEDIA_MODEL3D_630ED4, FORMAT.FORMAT_3973E0, FORMAT.GLB],
        [command.objectId, sha256]
      ),
      rid = model3DEdgeId(command.objectId);
    if (!existing || existing.part !== part) {
      if (Object.hasOwn(ctx.parts, part)) throw refusal('object-conflict', '3D asset identity collides');
      plan.binaryAdditions[part] = { base64: Buffer.from(bytes).toString('base64'), sha256 };
      addType(ctx.types, part, MODEL3D_MIME);
      typesChanged = true;
      if (existing) {
        plan.removed.push(existing.part);
        removeType(ctx.types, existing.part);
        drawingRels.doc.documentElement.removeChild(existing.edge.node);
      }
      const node = drawingRels.doc.createElementNS(P, 'Relationship');
      node.setAttribute('Id', rid);
      node.setAttribute('Type', MODEL3D_REL);
      node.setAttribute('Target', path.posix.relative(path.posix.dirname(drawingPart), part));
      drawingRels.doc.documentElement.appendChild(node);
      plan.replacements[drawingRels.file] = serialize(drawingRels.doc);
    }
    const fragment = model3DAnchorXml(command.objectId, number, placement, model, description);
    drawingSource = existing
      ? drawingSource.slice(0, existing.item.start) + fragment + drawingSource.slice(existing.item.end)
      : appendRaw(drawingSource, drawingDoc, fragment);
    plan.replacements[drawingPart] = drawingSource;
    plan.objectPart = part;
    normalized = {
      operation: command.operation,
      sheet: command.sheet,
      objectId: command.objectId,
      expectedObjectRevision: command.expectedObjectRevision,
      ...(create ? { model3d: model } : { patch: model }),
      anchor: placement,
      altText: description,
    };
  } else if (isSmartArtCommand(command)) {
    if (command.patch !== undefined) object(command.patch, ['layout', 'nodes', 'connections'], 'SmartArt patch');
    const model = normalizeSmartArt(
      existing
        ? {
            layout: command.patch?.layout ?? existing.smartart.layout,
            nodes: command.patch?.nodes ?? existing.smartart.nodes,
            ...(command.patch?.connections ? { connections: command.patch.connections } : {}),
          }
        : command.smartart
    );
    const placement = anchor(
        command.anchor,
        existing?.anchor ?? { ...defaultAnchor, widthEmu: 5715000, heightEmu: 3429000 }
      ),
      description = text(
        command.altText ?? existing?.properties.getAttribute('descr') ?? '',
        256,
        'SmartArt alternative text'
      );
    const number = existing
      ? integer(Number(existing.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
      : Math.max(
          0,
          ...canvas.objects.map((item) =>
            integer(Number(item.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
          )
        ) + 1;
    integer(number, 1, 2147483647, 'drawing ID');
    const graph = smartartPackage(command.objectId, model, placement, existingFormat),
      fragment = smartartAnchorXml(command.objectId, number, placement, description, existingFormat);
    if (create && Object.keys(graph.parts).some((part) => Object.hasOwn(ctx.parts, part)))
      throw refusal('object-conflict', 'SmartArt graph identity collides');
    Object.assign(plan.replacements, graph.parts);
    plan.objectPart = graph.names.data;
    if (create) {
      for (const [part, mime] of Object.entries(graph.types)) addType(ctx.types, part, mime);
      typesChanged = true;
      const edges = smartartEdgeIds(command.objectId, existingFormat);
      for (const [attr, role, suffix] of [
        ['dm', 'data', 'Data'],
        ['lo', 'layout', 'Layout'],
        ['qs', 'style', 'QuickStyle'],
        ['cs', 'colors', 'Colors'],
      ]) {
        const node = drawingRels.doc.createElementNS(P, 'Relationship');
        node.setAttribute('Id', edges[attr]);
        node.setAttribute('Type', R + '/diagram' + suffix);
        node.setAttribute('Target', path.posix.relative(path.posix.dirname(drawingPart), graph.names[role]));
        drawingRels.doc.documentElement.appendChild(node);
      }
    }
    drawingSource = existing
      ? drawingSource.slice(0, existing.item.start) + fragment + drawingSource.slice(existing.item.end)
      : appendRaw(drawingSource, drawingDoc, fragment);
    plan.replacements[drawingPart] = drawingSource;
    if (create) plan.replacements[drawingRels.file] = serialize(drawingRels.doc);
    normalized = {
      operation: command.operation,
      sheet: command.sheet,
      objectId: command.objectId,
      expectedObjectRevision: command.expectedObjectRevision,
      ...(create ? { smartart: model } : { patch: model }),
      anchor: placement,
      altText: description,
    };
  } else if (isInkStrokeCommand(command)) {
    if (create && !command.anchor) throw refusal('object-bad-input', 'A new stroke needs its explicit captured anchor');
    const model = normalizeInkStroke(existing?.stroke ?? command.stroke),
      placement = anchor(command.anchor, existing?.anchor),
      description = text(
        command.altText ?? existing?.properties.getAttribute('descr') ?? 'Freehand stroke',
        256,
        'stroke alternative text'
      );
    if (placement.widthEmu !== model.widthEmu || placement.heightEmu !== model.heightEmu)
      throw refusal('object-bad-input', 'Stroke placement must retain its original dimensions');
    const number = existing
      ? integer(Number(existing.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
      : Math.max(
          0,
          ...canvas.objects.map((item) =>
            integer(Number(item.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
          )
        ) + 1;
    const fragment = inkStrokeAnchorXml(command.objectId, number, model, placement, description);
    drawingSource = existing
      ? drawingSource.slice(0, existing.item.start) + fragment + drawingSource.slice(existing.item.end)
      : appendRaw(drawingSource, drawingDoc, fragment);
    plan.replacements[drawingPart] = drawingSource;
    plan.objectPart = drawingPart;
    normalized = {
      operation: command.operation,
      sheet: command.sheet,
      objectId: command.objectId,
      expectedObjectRevision: command.expectedObjectRevision,
      ...(create ? { stroke: model } : {}),
      anchor: placement,
      altText: description,
    };
  } else if (isShapeCommand(command)) {
    if (command.patch) object(command.patch, ['sample', 'text', 'color'], 'shape patch');
    const model = normalizeShape(existing ? { ...existing.shape, ...(command.patch ?? {}) } : command.shape),
      placement = anchor(
        command.anchor,
        existing?.anchor ?? {
          ...defaultAnchor,
          widthEmu:
            shapeSample(model.sample).family === 'icon'
              ? 914400
              : shapeSample(model.sample).family === 'wordart'
                ? 3048000
                : 2286000,
          heightEmu: shapeSample(model.sample).family === 'icon' ? 914400 : 1143000,
        }
      ),
      description = text(
        command.altText ?? existing?.properties.getAttribute('descr') ?? '',
        256,
        'shape alternative text'
      );
    const number = existing
      ? integer(Number(existing.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
      : Math.max(
          0,
          ...canvas.objects.map((item) =>
            integer(Number(item.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
          )
        ) + 1;
    integer(number, 1, 2147483647, 'drawing ID');
    const fragment = shapeAnchorXml(command.objectId, number, model, placement, description);
    drawingSource = existing
      ? drawingSource.slice(0, existing.item.start) + fragment + drawingSource.slice(existing.item.end)
      : appendRaw(drawingSource, drawingDoc, fragment);
    plan.replacements[drawingPart] = drawingSource;
    plan.objectPart = drawingPart;
    normalized = {
      operation: command.operation,
      sheet: command.sheet,
      objectId: command.objectId,
      expectedObjectRevision: command.expectedObjectRevision,
      ...(existing ? { patch: model } : { shape: model }),
      anchor: placement,
      altText: description,
    };
  } else {
    const image = ['insertImage', 'updateImage'].includes(command.operation),
      kind = image ? 'image' : 'chart',
      asset = image
        ? existing
          ? (await indexedObjectPreview({ metadataBytes, view, readRange, readMedia }, command.sheet, command.objectId))
              .image
          : indexedPngPayload(command.asset)
        : null,
      part = image
        ? (existing?.part ??
          formatTemplate([FORMAT.XL_MEDIA_IMAGE, FORMAT.FORMAT_3973E0, FORMAT.PNG], [command.objectId, asset.sha256]))
        : (existing?.part ?? formatTemplate([FORMAT.XL_CHARTS_CHART, FORMAT.XML], [command.objectId]));
    if (create && Object.hasOwn(ctx.parts, part)) throw refusal('object-conflict', 'Object package part collides');
    if (
      !image &&
      !existing &&
      ctx.all.filter((item) => item.managed && item.kind === 'chart').length >= INDEXED_OBJECT_LIMITS.charts
    )
      throw refusal('object-limit', 'Managed chart limit reached');
    if (existing && !image && incoming(ctx, part) !== 1)
      throw refusal('object-protected', 'Shared mutable chart is protected');
    const rid = existing?.edge.id ?? FORMAT.OBJECT + command.objectId.replaceAll('-', ''),
      number = existing
        ? integer(Number(existing.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
        : Math.max(
            0,
            ...canvas.objects.map((item) =>
              integer(Number(item.properties.getAttribute('id')), 1, 2147483647, 'drawing ID')
            )
          ) + 1;
    if (number > 2147483647) throw refusal('object-refused', 'Drawing nonvisual ID exhausted');
    let placement, description;
    if (image) {
      const width = Math.min(600, asset.widthPx) * 9525,
        height = Math.round((width * asset.heightPx) / asset.widthPx);
      placement = anchor(
        command.anchor,
        existing?.anchor ?? {
          ...defaultAnchor,
          widthEmu: width,
          heightEmu: height,
        }
      );
      if (
        Math.abs(placement.widthEmu * asset.heightPx - placement.heightEmu * asset.widthPx) >
        Math.max(asset.widthPx, asset.heightPx)
      )
        throw refusal('object-unsupported', 'PNG placement must preserve its aspect ratio');
      description = text(command.altText ?? existing?.properties.getAttribute('descr') ?? '', 256, 'image alt text');
      if (!existing)
        plan.binaryAdditions[part] = {
          base64: asset.base64,
          sha256: asset.sha256,
          widthPx: asset.widthPx,
          heightPx: asset.heightPx,
        };
      const defaults = children(ctx.types.documentElement).filter(
        (node) => node.localName === 'Default' && node.getAttribute('Extension').toLowerCase() === 'png'
      );
      if (defaults.length > 1 || (defaults.length && defaults[0].getAttribute('ContentType') !== 'image/png'))
        throw refusal('object-refused', 'PNG content type is ambiguous');
      if (!existing && !defaults.length) {
        const node = ctx.types.createElementNS(T, 'Default');
        node.setAttribute('Extension', 'png');
        node.setAttribute('ContentType', 'image/png');
        ctx.types.documentElement.appendChild(node);
        typesChanged = true;
      }
      normalized = {
        operation: command.operation,
        sheet: command.sheet,
        objectId: command.objectId,
        expectedObjectRevision: command.expectedObjectRevision,
        ...(!existing ? { asset } : {}),
        anchor: placement,
        altText: description,
      };
    } else {
      if (command.operation === 'updateChart') {
        object(
          command.patch ?? {},
          ['kind', 'range', 'title', 'legend', 'blanks', 'includeHidden', 'style'],
          'chart patch'
        );
        if (Object.hasOwn(command.patch ?? {}, 'style') && command.patch.style === undefined)
          throw refusal('object-bad-input', 'Choose a declared chart style; omit style to keep the saved preset');
      }
      const model = chartModel(
          existing
            ? {
                ...existing.chart,
                ...(command.patch ?? {}),
              }
            : command.chart
        ),
        data = await chartData(readRange, command.sheet, model, view);
      placement = anchor(command.anchor, existing?.anchor ?? defaultAnchor);
      description = model.title;
      plan.replacements[part] = chartXml(command.sheet, model, data);
      if (!existing) {
        addType(ctx.types, part, CT_CHART);
        typesChanged = true;
      }
      normalized = existing
        ? {
            operation: 'updateChart',
            sheet: command.sheet,
            objectId: command.objectId,
            expectedObjectRevision: command.expectedObjectRevision,
            patch: model,
            anchor: placement,
          }
        : {
            operation: 'createChart',
            sheet: command.sheet,
            objectId: command.objectId,
            expectedObjectRevision: null,
            chart: model,
            anchor: placement,
          };
    }
    const fragment = anchorXml(command.objectId, number, kind, rid, placement, description);
    if (existing)
      drawingSource = drawingSource.slice(0, existing.item.start) + fragment + drawingSource.slice(existing.item.end);
    else {
      drawingSource = appendRaw(drawingSource, drawingDoc, fragment);
      addEdge(drawingRels.doc, rid, kind, path.posix.relative(path.posix.dirname(drawingPart), part));
    }
    plan.replacements[drawingPart] = drawingSource;
    plan.replacements[drawingRels.file] = serialize(drawingRels.doc);
    plan.objectPart = part;
  }
  const typeText = serialize(ctx.types);
  if (typesChanged) plan.replacements['[Content_Types].xml'] = typeText;
  if (
    Buffer.byteLength(JSON.stringify(plan)) > 2 * 1024 * 1024 ||
    Buffer.byteLength(JSON.stringify(normalized)) > 2 * 1024 * 1024
  )
    throw refusal('object-limit', 'Object command/plan exceeds transport bound');
  if (['createChart', 'updateChart'].includes(command.operation)) {
    const chartParts = new Set(
      ctx.all.filter((item) => item.managed && item.kind === 'chart' && item.part).map((item) => item.part)
    );
    chartParts.add(plan.objectPart);
    let bytes = 0;
    for (const part of chartParts)
      bytes +=
        plan.replacements[part] !== undefined ? Buffer.byteLength(plan.replacements[part]) : ctx.parts[part].byteLength;
    if (bytes > INDEXED_OBJECT_LIMITS.chartBytes)
      throw refusal('object-limit', 'Managed chart metadata exceeds the 4 MiB authoring limit');
  }
  return {
    plan,
    command: normalized,
  };
}
async function validatedModel3D(entry, readMedia) {
  if (typeof readMedia !== 'function') throw refusal('object-unsupported', 'Bounded 3D media reader is unavailable');
  const bytes = await readMedia(entry.part, { maxBytes: MODEL3D_LIMITS.assetBytes }),
    asset = parseModel3DGlb(bytes);
  if (!entry.part.endsWith('-' + hash(bytes) + '.glb') || JSON.stringify(asset.model) !== JSON.stringify(entry.model3d))
    throw refusal('internal-mismatch', '3D bytes disagree with their managed model/hash');
  return asset;
}
export async function indexedObjectPreview({ metadataBytes, view, readRange, readMedia }, sheet, id) {
  const ctx = context(metadataBytes, view),
    entry = ctx.all.find((item) => item.id === id && item.canvas.sheet.name === sheet);
  if (!entry) throw refusal('object-bad-input', 'Object access is not granted');
  const result = summary(entry);
  if (entry.model3d) return { ...result, asset3d: await validatedModel3D(entry, readMedia) };
  if (entry.protectedReason) return result;
  if (['shape', 'smartart', 'ink'].includes(entry.kind)) return result;
  if (entry.kind === 'chart')
    return {
      ...result,
      data: await chartData(readRange, sheet, entry.chart, view),
    };
  if (typeof readMedia !== 'function') throw refusal('object-unsupported', 'Bounded PNG media reader is unavailable');
  const bytes = await readMedia(entry.part, {
    maxBytes: INDEXED_OBJECT_LIMITS.pngBytes,
  });
  if (!(bytes instanceof Uint8Array) || bytes.byteLength > INDEXED_OBJECT_LIMITS.pngBytes)
    throw refusal('object-limit', 'PNG media reader exceeded its byte bound');
  const asset = indexedPngPayload({
    mime: 'image/png',
    base64: Buffer.from(bytes).toString('base64'),
  });
  if (!entry.part.endsWith('-' + asset.sha256 + '.png'))
    throw refusal('internal-mismatch', 'PNG asset bytes disagree with their managed identity');
  return {
    ...result,
    image: asset,
  };
}
/** ZIP-directory only: no XML/media inflation or cell/source reads. Opaque
 * workbooks without our private chart parts keep their normal export path. */
export function indexedObjectHasManagedCharts(metadataBytes) {
  if (!(metadataBytes instanceof Uint8Array)) throw refusal('object-unsupported', 'Object metadata is unavailable');
  return Object.keys(openZipParts(metadataBytes)).some((part) => FORMAT_REGEX.CHART_PART.test(part));
}
function refreshedChartCaches(bytes, data) {
  const { source, doc } = parse(bytes, C, 'chartSpace'),
    lines = lineStarts(source);
  const patches = [];
  const replace = (node, kind, values) => {
    const start = lines[node.lineNumber - 1] + node.columnNumber - 1;
    if (!Number.isSafeInteger(start) || !source.startsWith('<' + node.tagName, start))
      throw refusal('object-unsupported', 'Chart cache has no qualified raw boundary');
    const end = xmlElementEnd(source, start),
      raw = source.slice(start, end);
    let at = 1,
      quoteChar = '';
    for (; at < raw.length; at++) {
      const char = raw[at];
      if (quoteChar) {
        if (char === quoteChar) quoteChar = '';
      } else if (char === '"' || char === "'") quoteChar = char;
      else if (char === '>') break;
    }
    const opening = raw.slice(0, at + 1).replace(/\/\s*>$/, '>'),
      prefix = node.prefix ? node.prefix + ':' : '';
    const points = values
      .map((value, index) =>
        value === null
          ? ''
          : `<${prefix}pt idx="${index}"><${prefix}v>${quote(Object.is(value, -0) ? '-0' : value)}</${prefix}v></${prefix}pt>`
      )
      .join('');
    patches.push({
      start,
      end,
      text: `${opening}<${prefix}ptCount val="${values.length}"/>${points}</${node.tagName}>`,
    });
  };
  elements(doc, 'ser', C).forEach((ser, index) => {
    replace(one(one(ser, 'tx', C), 'strCache', C), 'str', [data.series[index].name]);
    replace(one(one(ser, 'cat', C), 'strCache', C), 'str', data.categories);
    replace(one(one(ser, 'val', C), 'numCache', C), 'num', data.series[index].values);
  });
  patches.sort((left, right) => right.start - left.start);
  let next = source;
  for (const patch of patches) next = next.slice(0, patch.start) + patch.text + next.slice(patch.end);
  return next;
}
export async function indexedObjectCachePlan({ metadataBytes, view, readRange, sheet, parts }) {
  if (!indexedObjectHasManagedCharts(metadataBytes)) return { replacements: {} };
  const ctx = context(metadataBytes, view),
    replacements = {},
    managedChartCaches = {};
  let bytes = 0;
  for (const entry of ctx.all)
    if (
      entry.managed &&
      entry.kind === 'chart' &&
      (sheet === undefined || entry.canvas.sheet.name === sheet) &&
      (parts === undefined || parts.includes(entry.part))
    ) {
      if (entry.protectedReason)
        throw refusal('object-unsupported', 'Managed chart cannot be safely exported: ' + entry.protectedReason);
      const data = await chartData(readRange, entry.canvas.sheet.name, entry.chart, view);
      const xml = refreshedChartCaches(ctx.parts[entry.part], data);
      bytes += Buffer.byteLength(xml);
      if (bytes > INDEXED_OBJECT_LIMITS.chartBytes)
        throw refusal('object-limit', 'Managed chart caches exceed the 4 MiB export bound');
      replacements[entry.part] = xml;
      managedChartCaches[entry.part] = { xml, sourceSha256: hash(ctx.parts[entry.part]) };
    }
  return {
    replacements,
    managedChartCaches,
  };
}
export function indexedObjectReferenceGuard(metadataBytes, view, command) {
  const ctx = context(metadataBytes, view),
    objects = ctx.all.filter((item) => item.managed && item.canvas.sheet.name === command.sheet);
  if (command.operation === 'duplicateSheet' && objects.length)
    return 'Duplicating managed object identities requires qualified independent ownership';
  if (
    objects.some((item) => item.kind === 'ink') &&
    ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'sortRows', 'duplicateSheet'].includes(command.operation)
  )
    return 'Stroke anchor transformations require separate qualification';
  if (
    objects.some((item) => item.kind === 'model3d') &&
    ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'sortRows', 'duplicateSheet'].includes(command.operation)
  )
    return '3D anchor transformations require separate qualification';
  if (
    objects.some((item) => item.kind === 'smartart') &&
    ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'sortRows', 'duplicateSheet'].includes(command.operation)
  )
    return 'SmartArt row/column/sort/duplication transformations require separate qualification';
  if (
    objects.some((item) => item.kind === 'shape') &&
    ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'sortRows'].includes(command.operation)
  )
    return 'Shape anchor transformations require separate qualification; remove the shapes first';
  if (
    objects.some((item) => item.kind === 'chart') &&
    ['renameSheet', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'sortRows', 'filterRows'].includes(
      command.operation
    )
  )
    return 'Chart reference/anchor transformation is not qualified; remove the managed chart first';
  return null;
}
