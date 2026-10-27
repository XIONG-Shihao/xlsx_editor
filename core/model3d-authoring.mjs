import { FORMAT, formatTemplate } from './format.mjs';
/** Closed Office model3d anchors. GLB bytes remain the model authority. */
import { DOMParser } from '@xmldom/xmldom';
import { normalizeModel3D } from './model3d-model.mjs';
import { assertXmlText, cellPosition } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const MODEL3D_NS = 'http://schemas.microsoft.com/office/drawing/2017/model3d';
export const MODEL3D_REL = 'http://schemas.microsoft.com/office/2017/06/relationships/model3d';
export const MODEL3D_MIME = 'model/gltf-binary';
const D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  A = 'http://schemas.openxmlformats.org/drawingml/2006/main',
  R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  M = FORMAT.URN_MODEL3D_SAMPLES_V1;
const quote = (s) => {
  assertXmlText(s);
  return s.replace(
    /[&<>"\r\n\t]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' })[c]
  );
};
export const model3DEdgeId = (id, format = FORMAT) => format.MODEL3D + id.replaceAll('-', '');
export function model3DAnchorXml(id, number, placement, value, alt = '', format = FORMAT) {
  const model = normalizeModel3D(value),
    p = cellPosition(placement.cell),
    rid = model3DEdgeId(id, format),
    { widthEmu: w, heightEmu: h } = placement;
  return formatTemplate(
    [
      format.XML_E4B9095B4E,
      format.XMLNS_A,
      format.XMLNS_R,
      format.XMLNS_AM3D,
      format.XMLNS,
      format.XML_FBC4CC58D8,
      format.XML_29D0C584D4,
      format.XML_4277E7D203,
      format.XML_603729CE77,
      format.XML_D6227924D9,
      format.CY,
      format.XML_1521CC04F4,
      format.NAME,
      format.DESCR,
      format.XML_6A6654760F,
      format.CY,
      format.XML_59B7FAB6D1,
      format.XML_C3C4F0DD62,
      format.XML_C15A4BC0B7,
      format.D,
      format.XML_BB0C576314,
      format.XML_BB1DBC2272,
      format.COLOR,
      format.X,
      format.Y,
      format.Z,
      format.XML_7561F05DA1,
    ],
    [
      D,
      A,
      R,
      MODEL3D_NS,
      format.URN_MODEL3D_SAMPLES_V1,
      p.col,
      placement.dxEmu,
      p.row,
      placement.dyEmu,
      w,
      h,
      number,
      id,
      quote(alt),
      w,
      h,
      MODEL3D_NS,
      rid,
      Math.min(w, h),
      5 * h,
      format.URN_MODEL3D_SAMPLES_V1,
      model.sample,
      model.color,
      model.rotation.x,
      model.rotation.y,
      model.rotation.z,
    ]
  );
}
const elements = (node, name, ns) => Array.from(node.getElementsByTagNameNS(ns, name));
const xmlModel = (node) => {
  const attrs = Object.fromEntries(
    Array.from(node.attributes ?? [])
      .filter((a) => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
      .map((a) => [a.namespaceURI + ':' + a.localName, a.value])
      .sort(([a], [b]) => a.localeCompare(b))
  );
  const children = Array.from(node.childNodes ?? []).filter((n) => n.nodeType === 1),
    text = Array.from(node.childNodes ?? [])
      .filter((n) => [3, 4].includes(n.nodeType))
      .map((n) => n.data)
      .join('');
  return [
    node.namespaceURI,
    node.localName,
    attrs,
    children.length && text.trim() === '' ? '' : text,
    children.map(xmlModel),
  ];
};
export function model3DFromAnchor(node, placement, id) {
  const props = elements(node, 'cNvPr', D),
    format = FORMAT,
    samples = elements(node, 'sample', format.URN_MODEL3D_SAMPLES_V1);
  if (props.length !== 1 || samples.length !== 1) throw refusal('object-ambiguous', 'Ambiguous 3D sample metadata');
  const s = samples[0],
    angle = (k) => {
      const n = s.getAttribute(k);
      if (!/^-?(?:0|[1-9]\d*)$/.test(n)) throw refusal('internal-invalid-field', 'Invalid 3D rotation');
      return Number(n);
    },
    model = normalizeModel3D({
      sample: s.getAttribute('sample'),
      color: s.getAttribute('color'),
      rotation: { x: angle('x'), y: angle('y'), z: angle('z') },
    }),
    number = Number(props[0].getAttribute('id')),
    alt = props[0].getAttribute('descr');
  if (!Number.isSafeInteger(number) || number < 1 || number > 2147483647 || alt.length > 256)
    throw refusal('object-malformed', 'Invalid 3D identity/description');
  const expected = new DOMParser().parseFromString(
    model3DAnchorXml(id, number, placement, model, alt, format),
    'application/xml'
  ).documentElement;
  if (JSON.stringify(xmlModel(node)) !== JSON.stringify(xmlModel(expected)))
    throw refusal('object-protected', 'Foreign or extended 3D anchor is protected');
  return model;
}
export const isModel3DCommand = (c) => ['createModel3D', 'updateModel3D', 'deleteModel3D'].includes(c?.operation);
export function captureModel3DCommands(commands) {
  const copied = structuredClone(commands);
  for (const c of copied) {
    if (!isModel3DCommand(c)) continue;
    const allowed = [
      'operation',
      'sheet',
      'objectId',
      'expectedObjectRevision',
      ...(c.operation === 'createModel3D'
        ? ['model3d', 'anchor', 'altText']
        : c.operation === 'updateModel3D'
          ? ['patch', 'anchor', 'altText']
          : []),
    ];
    if (Object.keys(c).some((k) => !allowed.includes(k)))
      throw refusal('internal-invalid-field', 'Invalid 3D command fields');
  }
  return copied;
}
