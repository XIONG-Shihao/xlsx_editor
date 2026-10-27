import { FORMAT, formatTemplate, stripFormatPrefix } from './format.mjs';
/** Closed ordinary DrawingML path profile. Package ownership/native writes are separate. */
import { DOMParser } from '@xmldom/xmldom';
import { INK_LIMITS, inkPen, normalizeInkStroke } from './ink-stroke.mjs';
import { refusal } from './refusal-codes.mjs';
export const INK_DRAW_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  INK_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const integer = (v, min, max, label) => {
  if (!Number.isSafeInteger(v) || v < min || v > max)
    throw refusal('internal-invalid-field', 'Invalid ink ' + label + '.');
  return v;
};
const uuid = (v) => typeof v === 'string' && /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(v);
const text = (v, max) => {
  if (
    typeof v !== 'string' ||
    v.length > max ||
    Array.from(v).some((c) => {
      const n = c.codePointAt(0);
      return (n < 32 && ![9, 10, 13].includes(n)) || (n >= 0xd800 && n <= 0xdfff) || n === 0xfffe || n === 0xffff;
    }) ||
    /_x[0-9a-f]{4}_/i.test(v)
  )
    throw refusal('internal-invalid-field', 'Invalid ink alternative text.');
  return v;
};
const q = (v) =>
  text(v, INK_LIMITS.alt).replace(
    /[&<>"\r\n\t]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' })[c]
  );
const children = (n) => Array.from(n.childNodes).filter((c) => c.nodeType === 1),
  nodes = (n, name, ns) => Array.from(n.getElementsByTagNameNS(ns, name));
function cell(ref) {
  const m = typeof ref === 'string' && /^([A-Z]{1,3})([1-9]\d*)$/.exec(ref);
  if (!m) throw refusal('object-bad-input', 'Use a canonical ink anchor cell.');
  const col = Array.from(m[1]).reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(m[2]) - 1;
  integer(col, 0, 16383, 'anchor column');
  integer(row, 0, 1048575, 'anchor row');
  return { col, row };
}
function name(col, row) {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters + (row + 1);
}
function placement(p, stroke) {
  if (!p || Object.keys(p).sort().join(',') !== 'cell,dxEmu,dyEmu,heightEmu,kind,widthEmu' || p.kind !== 'oneCell')
    throw refusal('object-bad-input', 'Ink requires one-cell placement.');
  cell(p.cell);
  integer(p.dxEmu, 0, INK_LIMITS.offsetEmu, 'column offset');
  integer(p.dyEmu, 0, INK_LIMITS.offsetEmu, 'row offset');
  if (p.widthEmu !== stroke.widthEmu || p.heightEmu !== stroke.heightEmu)
    throw refusal('object-bad-input', 'Ink placement must keep the original stroke dimensions.');
  return p;
}
export function inkStrokeAnchorXml(id, number, value, anchor, altText = '', format = FORMAT) {
  if (!uuid(id)) throw refusal('object-malformed', 'Invalid ink object identity.');
  integer(number, 1, 2147483647, 'drawing identity');
  const stroke = normalizeInkStroke(value),
    pen = inkPen(stroke.pen),
    p = cell(placement(anchor, stroke).cell),
    geometry = stroke.points
      .map(([x, y], i) => `<a:${i ? 'lnTo' : 'moveTo'}><a:pt x="${x}" y="${y}"/></a:${i ? 'lnTo' : 'moveTo'}>`)
      .join('');
  return formatTemplate(
    [
      format.XML_E4B9095B4E,
      format.XMLNS_A,
      format.XML_FBC4CC58D8,
      format.XML_29D0C584D4,
      format.XML_4277E7D203,
      format.XML_603729CE77,
      format.XML_D6227924D9,
      format.CY,
      format.XML_5A4B0FAB49,
      format.NAME,
      format.DESCR,
      format.XML_556D7186B8,
      format.CY,
      format.XML_EF944F0E32,
      format.B,
      format.XML_7CC1C7CAE4,
      format.H,
      format.XML_67A1FCCFE5,
      format.XML_D0E4EED93F,
      format.CAP,
      format.XML_B4E7779FDC,
      format.XML_29A6366017,
      format.XML_F7447F0FBD,
    ],
    [
      INK_DRAW_NS,
      INK_MAIN_NS,
      p.col,
      anchor.dxEmu,
      p.row,
      anchor.dyEmu,
      stroke.widthEmu,
      stroke.heightEmu,
      number,
      id,
      q(altText),
      stroke.widthEmu,
      stroke.heightEmu,
      stroke.widthEmu,
      stroke.heightEmu,
      stroke.widthEmu,
      stroke.heightEmu,
      geometry,
      pen.widthEmu,
      pen.cap === 'round' ? 'rnd' : 'flat',
      pen.color,
      pen.alpha,
    ]
  );
}
function parse(source) {
  if (
    typeof source !== 'string' ||
    new TextEncoder().encode(source).byteLength > INK_LIMITS.xmlBytes ||
    /<!DOCTYPE|<!ENTITY/i.test(source)
  )
    throw refusal('object-limit', 'Ink XML exceeds its closed profile.');
  const fail = (m) => {
      throw refusal('package-malformed', m);
    },
    doc = new DOMParser({
      normalizeLineEndings: (v) => v,
      errorHandler: { warning: fail, error: fail, fatalError: fail },
    }).parseFromString(source, 'application/xml');
  if (doc.documentElement?.namespaceURI !== INK_DRAW_NS || doc.documentElement.localName !== 'oneCellAnchor')
    throw refusal('object-unsupported', 'Unsupported ink anchor.');
  return doc.documentElement;
}
// Prefix/order/insignificant spacing do not change this known semantic profile.
function semantic(n) {
  const attrs = Array.from(n.attributes)
    .filter((a) => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
    .map((a) => [a.namespaceURI ?? '', a.localName, a.value])
    .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
  const entries = children(n),
    texts = Array.from(n.childNodes)
      .filter((c) => c.nodeType === 3 || c.nodeType === 4)
      .map((c) => c.nodeValue)
      .join('');
  return [n.namespaceURI, n.localName, attrs, entries.length ? texts.trim() : texts, entries.map(semantic)];
}
export function readInkStrokeAnchor(source) {
  const root = parse(source),
    one = (name, ns) => {
      const found = nodes(root, name, ns);
      if (found.length !== 1) throw refusal('object-ambiguous', 'Ambiguous ink ' + name + '.');
      return found[0];
    },
    number = (s) => {
      if (typeof s !== 'string' || !/^(0|[1-9]\d*)$/.test(s))
        throw refusal('internal-invalid-field', 'Invalid ink numeric metadata.');
      return integer(Number(s), 0, INK_LIMITS.extentEmu, 'coordinate');
    };
  const props = one('cNvPr', INK_DRAW_NS),
    id = stripFormatPrefix(props.getAttribute('name'), 'FORMAT_183450'),
    drawingNumber = Number(props.getAttribute('id')),
    path = one('path', INK_MAIN_NS),
    line = one('ln', INK_MAIN_NS),
    color = one('srgbClr', INK_MAIN_NS),
    alpha = one('alpha', INK_MAIN_NS);
  if (!uuid(id)) throw refusal('object-unsupported', 'Unsupported ink identity.');
  const pen = inkPen(INK_PENS_ID(line, color, alpha)),
    widthEmu = number(path.getAttribute('w')),
    heightEmu = number(path.getAttribute('h')),
    steps = children(path);
  if (steps.length < 2 || steps.length > INK_LIMITS.points)
    throw refusal('object-unsupported', 'Ink path point count is unsupported.');
  const points = steps.map((step, i) => {
    if (step.namespaceURI !== INK_MAIN_NS || step.localName !== (i ? 'lnTo' : 'moveTo'))
      throw refusal('object-unsupported', 'Only ordinary open ink polylines are qualified.');
    const pts = children(step);
    if (pts.length !== 1 || pts[0].namespaceURI !== INK_MAIN_NS || pts[0].localName !== 'pt')
      throw refusal('object-malformed', 'Invalid ink path point.');
    return [number(pts[0].getAttribute('x')), number(pts[0].getAttribute('y'))];
  });
  const stroke = normalizeInkStroke({ schema: 1, pen: pen.id, widthEmu, heightEmu, points }),
    col = number(one('col', INK_DRAW_NS).textContent),
    row = number(one('row', INK_DRAW_NS).textContent),
    anchor = {
      kind: 'oneCell',
      cell: name(integer(col, 0, 16383, 'column'), integer(row, 0, 1048575, 'row')),
      dxEmu: number(one('colOff', INK_DRAW_NS).textContent),
      dyEmu: number(one('rowOff', INK_DRAW_NS).textContent),
      widthEmu,
      heightEmu,
    },
    altText = props.getAttribute('descr');
  const expected = parse(inkStrokeAnchorXml(id, drawingNumber, stroke, anchor, altText, FORMAT));
  if (JSON.stringify(semantic(root)) !== JSON.stringify(semantic(expected)))
    throw refusal('object-unsupported', 'Imported or extended ink metadata is outside this stroke profile.');
  return { id, number: drawingNumber, stroke, anchor, altText };
}
function INK_PENS_ID(line, color, alpha) {
  const width = Number(line.getAttribute('w')),
    rgb = color.getAttribute('val'),
    opacity = Number(alpha.getAttribute('val')),
    cap = line.getAttribute('cap');
  for (const id of ['black', 'red', 'yellow-highlighter']) {
    const p = inkPen(id);
    if (p.widthEmu === width && p.color === rgb && p.alpha === opacity && (p.cap === 'round' ? 'rnd' : 'flat') === cap)
      return id;
  }
  throw refusal('object-unsupported', 'Unsupported ink pen style.');
}
