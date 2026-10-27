import { FORMAT, stripFormatPrefix } from './format.mjs';
/** Repository-authored editable DrawingML samples; no bitmap surrogate. */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { WORDART_SAMPLES, WORDART_LIMITS, wordArtRunXml } from './wordart-samples.mjs';
import { ICON_SAMPLES, ICON_COLOR, iconGeometryXml } from './icon-samples.mjs';
import { assertXmlText, cellPosition } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const SHAPE_DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const SHAPE_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const SHAPE_SAMPLES = Object.freeze([
  Object.freeze({
    id: 'text-plain',
    family: 'text',
    label: 'Plain text box',
    geometry: 'rect',
    fill: null,
    line: null,
    color: '1F2937',
    fontSize: 14,
    align: 'l',
    textBox: true,
  }),
  Object.freeze({
    id: 'text-note',
    family: 'text',
    label: 'Note text box',
    geometry: 'rect',
    fill: 'EAF2F8',
    line: '4472C4',
    color: '17365D',
    fontSize: 14,
    align: 'l',
    textBox: true,
  }),
  Object.freeze({
    id: 'shape-rectangle',
    family: 'shape',
    label: 'Rectangle',
    geometry: 'rect',
    fill: 'DCE6F1',
    line: '4472C4',
    color: '17365D',
    fontSize: 14,
    align: 'ctr',
    textBox: false,
  }),
  Object.freeze({
    id: 'shape-ellipse',
    family: 'shape',
    label: 'Ellipse',
    geometry: 'ellipse',
    fill: 'E2EFDA',
    line: '70AD47',
    color: '375623',
    fontSize: 14,
    align: 'ctr',
    textBox: false,
  }),
  Object.freeze({
    id: 'shape-arrow',
    family: 'shape',
    label: 'Right arrow',
    geometry: 'rightArrow',
    fill: 'FFF2CC',
    line: 'BF9000',
    color: '7F6000',
    fontSize: 14,
    align: 'ctr',
    textBox: false,
  }),
]);
export const ILLUSTRATION_SAMPLES = Object.freeze([...SHAPE_SAMPLES, ...ICON_SAMPLES, ...WORDART_SAMPLES]);
export const SHAPE_LIMITS = Object.freeze({ text: 1024, paragraphs: 8, alt: 256 });
export function shapeSample(id) {
  const sample = ILLUSTRATION_SAMPLES.find((s) => s.id === id);
  if (!sample) throw refusal('object-bad-input', 'Choose a supported text or shape sample');
  return sample;
}
export function normalizeShape(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((k) => !['sample', 'text', 'color'].includes(k))
  )
    throw refusal('internal-invalid-field', 'Invalid shape model');
  const sample = shapeSample(value.sample);
  if (sample.family === 'icon') {
    const color = value.color ?? ICON_COLOR;
    if (typeof color !== 'string' || !/^[0-9A-F]{6}$/.test(color) || (value.text !== undefined && value.text !== ''))
      throw refusal('object-bad-input', 'Icons support an empty text model and six-digit uppercase sRGB color');
    return { sample: sample.id, text: '', color };
  }
  if (value.color !== undefined)
    throw refusal('object-bad-input', 'Color editing is currently available for icon samples only');
  const text = value.text ?? sample.label;
  if (
    typeof text !== 'string' ||
    text.length > SHAPE_LIMITS.text ||
    text.split('\n').length > SHAPE_LIMITS.paragraphs ||
    /[\r\t]/.test(text) ||
    /_x[0-9a-f]{4}_/i.test(text)
  )
    throw refusal(
      'object-limit',
      'Shape text supports at most 1,024 characters/eight paragraphs without tabs or escape patterns'
    );
  if (
    sample.family === 'wordart' &&
    (text.length > WORDART_LIMITS.text || text.split('\n').length > WORDART_LIMITS.paragraphs)
  )
    throw refusal('object-limit', 'WordArt text supports at most 256 characters/four paragraphs');
  assertXmlText(text);
  return { sample: sample.id, text };
}
const q = (value) => {
  assertXmlText(String(value));
  return String(value).replace(
    /[&<>"\r\n\t]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' })[c]
  );
};
const fill = (color) => (color ? '<a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' : '<a:noFill/>');
export function shapeAnchorXml(id, number, shape, placement, altText = '', format = FORMAT) {
  const model = normalizeShape(shape),
    sample = shapeSample(model.sample),
    p = cellPosition(placement.cell),
    paragraphs = model.text
      .split('\n')
      .map(
        (text) =>
          '<a:p><a:pPr algn="' +
          sample.align +
          '"/><a:r>' +
          (sample.family === 'wordart'
            ? wordArtRunXml(sample)
            : '<a:rPr lang="en-US" sz="' +
              sample.fontSize * 100 +
              '" b="0" i="0">' +
              fill(sample.color) +
              '<a:latin typeface="Arial"/></a:rPr>') +
          '<a:t xml:space="preserve">' +
          q(text) +
          '</a:t></a:r></a:p>'
      )
      .join('');
  return (
    '<xdr:oneCellAnchor xmlns:xdr="' +
    SHAPE_DRAW_NS +
    '" xmlns:a="' +
    SHAPE_MAIN_NS +
    '"><xdr:from><xdr:col>' +
    p.col +
    '</xdr:col><xdr:colOff>' +
    placement.dxEmu +
    '</xdr:colOff><xdr:row>' +
    p.row +
    '</xdr:row><xdr:rowOff>' +
    placement.dyEmu +
    '</xdr:rowOff></xdr:from><xdr:ext cx="' +
    placement.widthEmu +
    '" cy="' +
    placement.heightEmu +
    '"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="' +
    number +
    format.NAME +
    id +
    '" descr="' +
    q(altText) +
    '"/><xdr:cNvSpPr txBox="' +
    Number(sample.textBox) +
    '"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' +
    placement.widthEmu +
    '" cy="' +
    placement.heightEmu +
    '"/></a:xfrm>' +
    (sample.family === 'icon'
      ? iconGeometryXml(sample)
      : '<a:prstGeom prst="' + sample.geometry + '"><a:avLst/></a:prstGeom>') +
    fill(sample.family === 'icon' ? model.color : sample.fill) +
    '<a:ln w="12700">' +
    fill(sample.line) +
    '</a:ln></xdr:spPr>' +
    (sample.family === 'icon'
      ? ''
      : '<xdr:txBody><a:bodyPr ' +
        (sample.family === 'wordart' ? 'fromWordArt="1" ' : '') +
        'wrap="square" anchor="ctr" lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/>' +
        paragraphs +
        '</xdr:txBody>') +
    '</xdr:sp><xdr:clientData/></xdr:oneCellAnchor>'
  );
}
const elements = (node, name, ns) => Array.from(node.getElementsByTagNameNS(ns, name));
export function shapeModelFromAnchor(node, placement) {
  const props = elements(node, 'cNvPr', SHAPE_DRAW_NS),
    texts = elements(node, 't', SHAPE_MAIN_NS),
    paragraphs = elements(node, 'p', SHAPE_MAIN_NS);
  if (props.length !== 1 || texts.length !== paragraphs.length || paragraphs.length > SHAPE_LIMITS.paragraphs)
    throw refusal('object-unsupported', 'Unsupported shape text/identity profile');
  const name = props[0].getAttribute('name'),
    id = stripFormatPrefix(name, 'FORMAT_183450') ?? '',
    number = Number(props[0].getAttribute('id')),
    alt = props[0].getAttribute('descr'),
    format = FORMAT;
  if (
    !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(id) ||
    !Number.isSafeInteger(number) ||
    number < 1 ||
    number > 2147483647 ||
    alt.length > SHAPE_LIMITS.alt
  )
    throw refusal('object-unsupported', 'Unsupported shape identity/alternative text');
  const serialize = (n) => new XMLSerializer().serializeToString(n),
    actual = serialize(node),
    text = texts.map((n) => n.textContent).join('\n');
  const iconColor = elements(node, 'spPr', SHAPE_DRAW_NS)[0]
    ?.getElementsByTagNameNS(SHAPE_MAIN_NS, 'solidFill')[0]
    ?.getElementsByTagNameNS(SHAPE_MAIN_NS, 'srgbClr')[0]
    ?.getAttribute('val');
  if (paragraphs.length === 0) {
    for (const sample of ICON_SAMPLES) {
      let model;
      try {
        model = normalizeShape({ sample: sample.id, text: '', color: iconColor });
      } catch {
        continue;
      }
      const expected = new DOMParser().parseFromString(
        shapeAnchorXml(id, number, model, placement, alt, format),
        'application/xml'
      ).documentElement;
      if (actual === serialize(expected)) return model;
    }
    throw refusal('object-protected', 'Imported or extended icon path is preserved and protected');
  }
  for (const sample of [...SHAPE_SAMPLES, ...WORDART_SAMPLES]) {
    let model;
    try {
      model = normalizeShape({ sample: sample.id, text });
    } catch {
      continue;
    }
    const expected = new DOMParser().parseFromString(
      shapeAnchorXml(id, number, model, placement, alt, format),
      'application/xml'
    ).documentElement;
    if (actual === serialize(expected)) return model;
  }
  throw refusal('object-protected', 'Imported or extended shape profile is preserved and protected');
}
export const isShapeCommand = (c) => ['createShape', 'updateShape', 'deleteShape'].includes(c?.operation);
export function captureShapeCommands(commands) {
  const copied = structuredClone(commands);
  for (const c of copied) {
    if (!isShapeCommand(c)) continue;
    const allowed = [
      'operation',
      'sheet',
      'objectId',
      'expectedObjectRevision',
      ...(c.operation === 'createShape'
        ? ['shape', 'anchor', 'altText']
        : c.operation === 'updateShape'
          ? ['patch', 'anchor', 'altText']
          : []),
    ];
    if (Object.keys(c).some((k) => !allowed.includes(k)))
      throw refusal('internal-invalid-field', 'Invalid shape command fields');
  }
  return copied;
}
