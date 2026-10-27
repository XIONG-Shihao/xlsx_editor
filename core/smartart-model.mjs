import { FORMAT } from './format.mjs';
/** Closed semantic SmartArt models and real linked DiagramML parts.
 * The reader reconstructs every declared part; native authoring is a later phase. */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { assertXmlText, cellPosition } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const SMARTART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
export const SMARTART_CACHE_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main',
  D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  P = 'http://schemas.openxmlformats.org/package/2006/relationships',
  T = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const SMARTART_SAMPLES = Object.freeze([
  Object.freeze({
    id: 'process',
    label: 'Process',
    algorithm: 'lin',
    labels: Object.freeze(['Plan', 'Build', 'Review']),
  }),
  Object.freeze({
    id: 'hierarchy',
    label: 'Hierarchy',
    algorithm: 'hierRoot',
    labels: Object.freeze(['Lead', 'Design', 'Build']),
  }),
  Object.freeze({ id: 'cycle', label: 'Cycle', algorithm: 'cycle', labels: Object.freeze(['Plan', 'Do', 'Review']) }),
]);
export const SMARTART_LIMITS = Object.freeze({ nodes: 6, text: 128, partBytes: 256 * 1024, graphBytes: 1024 * 1024 });
const keys = (o, allowed) => {
  if (!o || typeof o !== 'object' || Array.isArray(o) || Object.keys(o).some((k) => !allowed.includes(k)))
    throw refusal('internal-invalid-field', 'Invalid SmartArt model fields');
};
const q = (s) => {
  assertXmlText(s);
  return s.replace(
    /[&<>"\n\r\t]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\n': '&#10;', '\r': '&#13;', '\t': '&#9;' })[c]
  );
};
export function normalizeSmartArt(value) {
  keys(value, ['layout', 'nodes', 'connections']);
  const sample = SMARTART_SAMPLES.find((s) => s.id === value.layout);
  if (!sample) throw refusal('object-bad-input', 'Choose Process, Hierarchy or Cycle');
  if (!Array.isArray(value.nodes) || value.nodes.length < 2 || value.nodes.length > SMARTART_LIMITS.nodes)
    throw refusal('object-refused', 'SmartArt supports two to six nodes');
  const nodes = value.nodes.map((n, i) => {
    keys(n, ['id', 'text']);
    if (
      n.id !== 'n' + (i + 1) ||
      typeof n.text !== 'string' ||
      !n.text.trim() ||
      n.text.length > SMARTART_LIMITS.text ||
      /[\r\n\t]/.test(n.text) ||
      /_x[0-9a-f]{4}_/i.test(n.text)
    )
      throw refusal(
        'object-limit',
        'SmartArt needs ordered node IDs and nonblank single-paragraph text up to 128 characters'
      );
    assertXmlText(n.text);
    return { id: n.id, text: n.text };
  });
  const connections =
    sample.id === 'hierarchy'
      ? nodes.slice(1).map((n) => ({ from: 'n1', to: n.id }))
      : nodes.slice(1).map((n, i) => ({ from: nodes[i].id, to: n.id }));
  if (sample.id === 'cycle') connections.push({ from: nodes.at(-1).id, to: 'n1' });
  if (value.connections !== undefined) {
    if (!Array.isArray(value.connections) || value.connections.length !== connections.length)
      throw refusal('object-bad-input', 'Connections must match the selected sample topology');
    value.connections.forEach((c, i) => {
      keys(c, ['from', 'to']);
      if (c.from !== connections[i].from || c.to !== connections[i].to)
        throw refusal('object-bad-input', 'Connections must match the selected sample topology');
    });
  }
  return { layout: sample.id, nodes, connections };
}
export function smartartSampleModel(layout) {
  const sample = SMARTART_SAMPLES.find((s) => s.id === layout);
  if (!sample) throw refusal('internal-unknown-reference', 'Unknown SmartArt sample');
  return normalizeSmartArt({ layout, nodes: sample.labels.map((text, i) => ({ id: 'n' + (i + 1), text })) });
}
export function smartartGeometry(value) {
  const model = normalizeSmartArt(value),
    n = model.nodes.length;
  let boxes;
  if (model.layout === 'process') {
    const step = 920 / n,
      w = Math.min(260, step - 30);
    boxes = model.nodes.map((node, i) => ({ ...node, x: 40 + i * step + (step - w) / 2, y: 220, w, h: 160 }));
  } else if (model.layout === 'hierarchy') {
    const step = 920 / (n - 1),
      w = Math.min(260, step - 24);
    boxes = model.nodes.map((node, i) => ({
      ...node,
      x: i === 0 ? 370 : 40 + (i - 1) * step + (step - w) / 2,
      y: i === 0 ? 50 : 380,
      w: i === 0 ? 260 : w,
      h: 150,
    }));
  } else {
    const radius = 180;
    boxes = model.nodes.map((node, i) => {
      const a = -Math.PI / 2 + (2 * Math.PI * i) / n;
      return { ...node, x: 500 + Math.cos(a) * 320 - 90, y: 300 + Math.sin(a) * radius - 55, w: 180, h: 110 };
    });
  }
  const byId = new Map(boxes.map((b) => [b.id, b]));
  return {
    width: 1000,
    height: 600,
    boxes,
    edges: model.connections.map((c) => {
      const a = byId.get(c.from),
        b = byId.get(c.to),
        dx = b.x + b.w / 2 - a.x - a.w / 2,
        dy = b.y + b.h / 2 - a.y - a.h / 2;
      const side = (box, x, y) => {
        const t = 1 / Math.max(Math.abs(x) / (box.w / 2), Math.abs(y) / (box.h / 2));
        return [box.x + box.w / 2 + x * t, box.y + box.h / 2 + y * t];
      };
      const start = side(a, dx, dy),
        end = side(b, -dx, -dy);
      return { ...c, x1: start[0], y1: start[1], x2: end[0], y2: end[1] };
    }),
  };
}
const uuid = (s) => typeof s === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(s);
const guid = (id, slot) =>
  '{' +
  (id.slice(0, 28) + ((Number.parseInt(id.slice(28), 16) ^ slot) >>> 0).toString(16).padStart(8, '0')).toUpperCase() +
  '}';
export function smartartPartNames(id, format = FORMAT) {
  if (!uuid(id)) throw refusal('object-malformed', 'Invalid SmartArt identity');
  return Object.fromEntries(
    ['data', 'layout', 'style', 'colors', 'drawing'].map((role) => [
      role,
      format.XL_DIAGRAMS_SMARTART + id + '-' + role + '.xml',
    ])
  );
}
export function smartartEdgeIds(id, format = FORMAT) {
  if (!uuid(id)) throw refusal('object-malformed', 'Invalid SmartArt identity');
  return Object.fromEntries(
    ['dm', 'lo', 'qs', 'cs'].map((role) => [role, format.DIAGRAM + id.replaceAll('-', '') + '_' + role])
  );
}
const fill = (c) => '<a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill>';
const textBody = (text) =>
  '<a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400">' +
  fill('FFFFFF') +
  '<a:latin typeface="Arial"/></a:rPr><a:t xml:space="preserve">' +
  q(text) +
  '</a:t></a:r></a:p>';
export function smartartAnchorXml(id, number, anchor, alt = '', format = FORMAT) {
  if (
    !uuid(id) ||
    !Number.isSafeInteger(number) ||
    number < 1 ||
    number > 2147483647 ||
    typeof alt !== 'string' ||
    alt.length > 256
  )
    throw refusal('object-malformed', 'Invalid SmartArt anchor identity');
  const p = cellPosition(anchor.cell),
    e = smartartEdgeIds(id, format);
  for (const [key, high, min] of [
    ['dxEmu', 10000000, 0],
    ['dyEmu', 10000000, 0],
    ['widthEmu', 50000000, 1],
    ['heightEmu', 50000000, 1],
  ])
    if (!Number.isSafeInteger(anchor[key]) || anchor[key] < min || anchor[key] > high)
      throw refusal('internal-invalid-field', 'Invalid SmartArt anchor bounds');
  return (
    '<xdr:oneCellAnchor xmlns:xdr="' +
    D +
    '" xmlns:a="' +
    A +
    '" xmlns:dgm="' +
    SMARTART_NS +
    '" xmlns:r="' +
    R +
    '"><xdr:from><xdr:col>' +
    p.col +
    '</xdr:col><xdr:colOff>' +
    anchor.dxEmu +
    '</xdr:colOff><xdr:row>' +
    p.row +
    '</xdr:row><xdr:rowOff>' +
    anchor.dyEmu +
    '</xdr:rowOff></xdr:from><xdr:ext cx="' +
    anchor.widthEmu +
    '" cy="' +
    anchor.heightEmu +
    '"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="' +
    number +
    format.NAME +
    id +
    '" descr="' +
    q(alt) +
    '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    anchor.widthEmu +
    '" cy="' +
    anchor.heightEmu +
    '"/></xdr:xfrm><a:graphic><a:graphicData uri="' +
    SMARTART_NS +
    '"><dgm:relIds r:dm="' +
    e.dm +
    '" r:lo="' +
    e.lo +
    '" r:qs="' +
    e.qs +
    '" r:cs="' +
    e.cs +
    '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>'
  );
}
export function smartartPackage(id, value, anchor, format = FORMAT) {
  const model = normalizeSmartArt(value),
    names = smartartPartNames(id, format),
    sample = SMARTART_SAMPLES.find((s) => s.id === model.layout),
    uid = format.URN_SMARTART_LAYOUT + model.layout,
    n = model.nodes.length,
    points = [],
    connections = [];
  const pt = (slot, type, props = '', text = '', extra = '') =>
    '<dgm:pt modelId="' +
    guid(id, slot) +
    '" type="' +
    type +
    '"' +
    extra +
    '><dgm:prSet' +
    props +
    '/><dgm:spPr/>' +
    (type === 'node' ? '<dgm:t>' + textBody(text) + '</dgm:t>' : '') +
    '</dgm:pt>';
  points.push(pt(0, 'doc'));
  model.nodes.forEach((node, i) => {
    points.push(pt(i + 1, 'node', ' phldrT="Node"', node.text));
    points.push(pt(200 + i, 'parTrans', '', '', ' cxnId="' + guid(id, 100 + i) + '"'));
    points.push(pt(300 + i, 'sibTrans', '', '', ' cxnId="' + guid(id, 100 + i) + '"'));
    const parent = model.layout === 'hierarchy' && i > 0 ? 1 : 0;
    connections.push(
      '<dgm:cxn modelId="' +
        guid(id, 100 + i) +
        '" type="parOf" srcId="' +
        guid(id, parent) +
        '" destId="' +
        guid(id, i + 1) +
        '" srcOrd="' +
        (model.layout === 'hierarchy' && i > 0 ? i - 1 : i) +
        '" destOrd="0" parTransId="' +
        guid(id, 200 + i) +
        '" sibTransId="' +
        guid(id, 300 + i) +
        '"/>'
    );
  });
  const pres = (slot, assoc, name) => {
    points.push(
      pt(slot, 'pres', ' presAssocID="' + guid(id, assoc) + '" presName="' + name + '" presStyleLbl="node1"')
    );
    connections.push(
      '<dgm:cxn modelId="' +
        guid(id, 800 + slot) +
        '" type="presOf" srcId="' +
        guid(id, assoc) +
        '" destId="' +
        guid(id, slot) +
        '" srcOrd="0" destOrd="0" presId="' +
        uid +
        '"/>'
    );
  };
  pres(400, 0, 'diagram');
  model.nodes.forEach((node, i) => {
    pres(401 + i, i + 1, 'node');
    connections.push(
      '<dgm:cxn modelId="' +
        guid(id, 1500 + i) +
        '" type="presParOf" srcId="' +
        guid(id, 400) +
        '" destId="' +
        guid(id, 401 + i) +
        '" srcOrd="' +
        i +
        '" destOrd="0" presId="' +
        uid +
        '"/>'
    );
  });
  model.connections.forEach((c, i) => {
    pres(500 + i, model.layout === 'hierarchy' ? 201 + i : 300 + i, 'transition');
    connections.push(
      '<dgm:cxn modelId="' +
        guid(id, 1600 + i) +
        '" type="presParOf" srcId="' +
        guid(id, 400) +
        '" destId="' +
        guid(id, 500 + i) +
        '" srcOrd="' +
        (n + i) +
        '" destOrd="0" presId="' +
        uid +
        '"/>'
    );
  });
  const data =
    '<dgm:dataModel xmlns:dgm="' +
    SMARTART_NS +
    '" xmlns:a="' +
    A +
    '" xmlns:dsp="' +
    SMARTART_CACHE_NS +
    '"><dgm:ptLst>' +
    points.join('') +
    '</dgm:ptLst><dgm:cxnLst>' +
    connections.join('') +
    '</dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri="' +
    SMARTART_CACHE_NS +
    format.XML_DE7C10CFEE +
    SMARTART_NS +
    '"/></a:ext></dgm:extLst></dgm:dataModel>';
  const nodeLayout =
    '<dgm:layoutNode name="node" styleLbl="node1"><dgm:alg type="sp"/><dgm:presOf axis="self" ptType="node"/><dgm:shape type="roundRect"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>';
  const layout =
    '<dgm:layoutDef xmlns:dgm="' +
    SMARTART_NS +
    '" xmlns:a="' +
    A +
    '" uniqueId="' +
    uid +
    format.XML_B100457FA8 +
    sample.label +
    '"/><dgm:desc val="Repository-authored bounded sample"/><dgm:catLst><dgm:cat type="' +
    (model.layout === 'hierarchy' ? 'hierarchy' : model.layout) +
    '" pri="1"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:alg type="' +
    sample.algorithm +
    '"/><dgm:presOf axis="self"/><dgm:forEach axis="' +
    (model.layout === 'hierarchy' ? 'des' : 'ch') +
    '" ptType="node" hideLastTrans="' +
    (model.layout === 'cycle' ? '0' : '1') +
    '">' +
    nodeLayout +
    '</dgm:forEach><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:layoutDef>';
  const references =
    '<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>';
  const style =
    '<dgm:styleDef xmlns:dgm="' +
    SMARTART_NS +
    '" xmlns:a="' +
    A +
    format.XML_988BB12941 +
    references +
    '</dgm:style></dgm:styleLbl></dgm:styleDef>';
  const colors = '<dgm:colorsDef xmlns:dgm="' + SMARTART_NS + '" xmlns:a="' + A + format.XML_FBD63D1D52;
  const g = smartartGeometry(model),
    w = anchor.widthEmu,
    h = anchor.heightEmu,
    rx = (x) => Math.round((x * w) / g.width),
    ry = (y) => Math.round((y * h) / g.height);
  const nodeShapes = g.boxes
    .map(
      (b, i) =>
        '<dsp:sp modelId="' +
        guid(id, 401 + i) +
        '"><dsp:nvSpPr><dsp:cNvPr id="' +
        (i + 1) +
        '" name="' +
        b.id +
        '"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="' +
        rx(b.x) +
        '" y="' +
        ry(b.y) +
        '"/><a:ext cx="' +
        rx(b.w) +
        '" cy="' +
        ry(b.h) +
        '"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
        fill('4472C4') +
        '<a:ln w="12700">' +
        fill('2F5597') +
        '</a:ln></dsp:spPr><dsp:txBody>' +
        textBody(b.text) +
        '</dsp:txBody></dsp:sp>'
    )
    .join('');
  const edgeShapes = g.edges
    .map(
      (e, i) =>
        '<dsp:sp modelId="' +
        guid(id, 500 + i) +
        '"><dsp:nvSpPr><dsp:cNvPr id="' +
        (100 + i) +
        '" name="' +
        e.from +
        '-' +
        e.to +
        '"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm' +
        (e.x2 < e.x1 ? ' flipH="1"' : '') +
        (e.y2 < e.y1 ? ' flipV="1"' : '') +
        '><a:off x="' +
        rx(Math.min(e.x1, e.x2)) +
        '" y="' +
        ry(Math.min(e.y1, e.y2)) +
        '"/><a:ext cx="' +
        rx(Math.abs(e.x2 - e.x1)) +
        '" cy="' +
        ry(Math.abs(e.y2 - e.y1)) +
        '"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="19050">' +
        fill('4472C4') +
        '<a:tailEnd type="triangle" w="med" len="med"/></a:ln></dsp:spPr></dsp:sp>'
    )
    .join('');
  const drawing =
    '<dsp:drawing xmlns:dsp="' +
    SMARTART_CACHE_NS +
    '" xmlns:a="' +
    A +
    '"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name="SmartArt"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    w +
    '" cy="' +
    h +
    '"/><a:chOff x="0" y="0"/><a:chExt cx="' +
    w +
    '" cy="' +
    h +
    '"/></a:xfrm></dsp:grpSpPr>' +
    edgeShapes +
    nodeShapes +
    '</dsp:spTree></dsp:drawing>';
  const dataRels =
    '<Relationships xmlns="' + P + format.XML_28C7718E93 + names.drawing.split('/').at(-1) + '"/></Relationships>';
  const parts = Object.fromEntries(
    Object.entries({ data, layout, style, colors, drawing }).map(([role, xml]) => [names[role], xml])
  );
  parts['xl/diagrams/_rels/' + names.data.split('/').at(-1) + '.rels'] = dataRels;
  return {
    model,
    names,
    parts,
    types: Object.fromEntries(
      Object.entries(names).map(([role, path]) => [
        path,
        role === 'drawing'
          ? 'application/vnd.ms-office.drawingml.diagramDrawing+xml'
          : 'application/vnd.openxmlformats-officedocument.drawingml.diagram' +
            { data: 'Data', layout: 'Layout', style: 'Style', colors: 'Colors' }[role] +
            '+xml',
      ])
    ),
  };
}
const parse = (bytes) => {
  const raw = typeof bytes === 'string' ? bytes : strFromU8(bytes ?? new Uint8Array());
  if (raw.length > SMARTART_LIMITS.partBytes || /<!DOCTYPE|<!ENTITY/i.test(raw))
    throw refusal('internal-missing-field', 'SmartArt XML is missing, extended or oversized');
  const errors = [],
    doc = new DOMParser({
      errorHandler: { warning: (s) => errors.push(s), error: (s) => errors.push(s), fatalError: (s) => errors.push(s) },
    }).parseFromString(raw, 'application/xml');
  if (errors.length || !doc.documentElement) throw refusal('object-malformed', 'Invalid SmartArt XML');
  return doc;
};
const serialize = (n) => new XMLSerializer().serializeToString(n),
  elements = (n, tag, ns) => Array.from(n.getElementsByTagNameNS(ns, tag));
const equal = (node, xml) => serialize(node) === serialize(parse(xml).documentElement);
export function smartartModelFromPackage(anchorNode, parts, drawingPart, anchor, id) {
  const format = FORMAT;
  const names = smartartPartNames(id, format),
    e = smartartEdgeIds(id, format),
    props = elements(anchorNode, 'cNvPr', D);
  if (props.length !== 1 || props[0].getAttribute('name') !== format.FORMAT_183450 + id)
    throw refusal('object-bad-input', 'SmartArt identity is not owned');
  const number = Number(props[0].getAttribute('id')),
    alt = props[0].getAttribute('descr');
  if (!equal(anchorNode, smartartAnchorXml(id, number, anchor, alt, format)))
    throw refusal('object-protected', 'Imported or extended SmartArt anchor is protected');
  const types = parse(parts['[Content_Types].xml']),
    rels = parse(parts[drawingPart.replace(/([^/]+)$/, '_rels/$1.rels')]);
  for (const [attr, role] of [
    ['dm', 'data'],
    ['lo', 'layout'],
    ['qs', 'style'],
    ['cs', 'colors'],
  ]) {
    const links = elements(rels, 'Relationship', P).filter((n) => n.getAttribute('Id') === e[attr]);
    if (
      links.length !== 1 ||
      links[0].getAttribute('Type') !==
        R + '/diagram' + { data: 'Data', layout: 'Layout', style: 'QuickStyle', colors: 'Colors' }[role] ||
      links[0].getAttribute('Target') !== '../diagrams/' + names[role].split('/').at(-1) ||
      links[0].hasAttribute('TargetMode')
    )
      throw refusal('object-stale', 'SmartArt edge is shared, external or changed');
  }
  const data = parse(parts[names.data]),
    layout = parse(parts[names.layout]),
    kind = layout.documentElement.getAttribute('uniqueId').replace(format.URN_SMARTART_LAYOUT, '');
  if (data.documentElement.namespaceURI !== SMARTART_NS || data.documentElement.localName !== 'dataModel')
    throw refusal('object-unsupported', 'Unsupported SmartArt data model');
  const nodes = elements(data, 'pt', SMARTART_NS)
    .filter((n) => n.getAttribute('type') === 'node')
    .map((n, i) => ({
      id: 'n' + (i + 1),
      text: elements(n, 't', A)
        .map((t) => t.textContent)
        .join(''),
    }));
  const model = normalizeSmartArt({ layout: kind, nodes }),
    expected = smartartPackage(id, model, anchor, format);
  let bytes = 0;
  for (const [path, xml] of Object.entries(expected.parts)) {
    const raw = parts[path];
    if (!raw) throw refusal('internal-missing-field', 'Missing SmartArt graph part');
    bytes += typeof raw === 'string' ? raw.length : raw.byteLength;
    if (!equal(parse(raw).documentElement, xml))
      throw refusal('object-protected', 'Imported or extended SmartArt data/layout/cache/style is protected');
  }
  if (bytes > SMARTART_LIMITS.graphBytes) throw refusal('object-limit', 'SmartArt graph exceeds its bound');
  for (const [path, mime] of Object.entries(expected.types)) {
    const matches = elements(types, 'Override', T).filter((n) => n.getAttribute('PartName') === '/' + path);
    if (matches.length !== 1 || matches[0].getAttribute('ContentType') !== mime)
      throw refusal('internal-mismatch', 'SmartArt content type mismatch');
  }
  // Exactly one incoming owner per graph part; no unexpected outgoing part edges.
  const owners = new Map(Object.values(names).map((path) => [path, 0]));
  for (const [path, raw] of Object.entries(parts)) {
    if (!path.endsWith('.rels')) continue;
    const owner = path === '_rels/.rels' ? '' : path.replace('/_rels/', '/').replace(/\.rels$/, ''),
      base = owner.slice(0, owner.lastIndexOf('/') + 1),
      dataRels = 'xl/diagrams/_rels/' + names.data.split('/').at(-1) + '.rels';
    if (Object.values(names).includes(owner) && path !== dataRels)
      throw refusal('object-protected', 'Dependent SmartArt part graph is protected');
    const root = parse(raw);
    for (const node of elements(root, 'Relationship', P)) {
      const mode = node.getAttribute('TargetMode');
      if (mode === 'External') continue;
      if (mode && mode !== 'Internal') throw refusal('internal-unknown-reference', 'Unknown relationship target mode');
      let target;
      try {
        target = decodeURIComponent(node.getAttribute('Target'));
      } catch {
        throw refusal('object-ambiguous', 'Ambiguous relationship target');
      }
      if (target.includes('\\') || target.includes(String.fromCharCode(0)))
        throw refusal('object-malformed', 'Invalid relationship target');
      const stack = target.startsWith('/') ? [] : base.split('/').filter(Boolean);
      for (const part of target.split('/'))
        if (part === '..') {
          if (!stack.length) throw refusal('object-unsupported', 'Relationship target escapes the package');
          stack.pop();
        } else if (part !== '.' && part) stack.push(part);
      const to = stack.join('/');
      if (owners.has(to)) owners.set(to, owners.get(to) + 1);
    }
  }
  for (const path of Object.values(names))
    if (owners.get(path) !== 1) throw refusal('object-protected', 'Shared SmartArt graph is protected');
  return { model, names, parts: expected.parts };
}

// Internal owner commands; the public authoring gallery is qualified separately.
export const isSmartArtCommand = (c) => ['createSmartArt', 'updateSmartArt', 'deleteSmartArt'].includes(c?.operation);
export function captureSmartArtCommands(commands) {
  const copied = structuredClone(commands);
  for (const c of copied) {
    if (!isSmartArtCommand(c)) continue;
    const allowed = [
      'operation',
      'sheet',
      'objectId',
      'expectedObjectRevision',
      ...(c.operation === 'createSmartArt'
        ? ['smartart', 'anchor', 'altText']
        : c.operation === 'updateSmartArt'
          ? ['patch', 'anchor', 'altText']
          : []),
    ];
    if (Object.keys(c).some((k) => !allowed.includes(k)))
      throw refusal('internal-invalid-field', 'Invalid SmartArt command fields');
  }
  return copied;
}
