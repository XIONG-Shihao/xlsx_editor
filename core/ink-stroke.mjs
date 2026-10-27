/** Portable bounded vector strokes; no pressure, native device or imported-ink claim. */
import { refusal } from './refusal-codes.mjs';

export const INK_EMU_PER_PIXEL = 9525;
export const INK_LIMITS = Object.freeze({
  points: 512,
  inputPoints: 2048,
  extentEmu: 50_000_000,
  offsetEmu: 10_000_000,
  alt: 256,
  xmlBytes: 256 * 1024,
});
export const INK_PENS = Object.freeze([
  Object.freeze({ id: 'black', label: 'Black pen', color: '000000', widthEmu: 19050, alpha: 100000, cap: 'round' }),
  Object.freeze({ id: 'red', label: 'Red pen', color: 'D32F2F', widthEmu: 19050, alpha: 100000, cap: 'round' }),
  Object.freeze({
    id: 'yellow-highlighter',
    label: 'Yellow highlighter',
    color: 'FFFF00',
    widthEmu: 114300,
    alpha: 30000,
    cap: 'flat',
  }),
]);
const exact = (value, keys, label) => {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== [...keys].sort().join(',')
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields.');
};
const integer = (value, min, max, label) => {
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label + '.');
  return value;
};
export function inkPen(id) {
  const pen = INK_PENS.find((p) => p.id === id);
  if (!pen) throw refusal('object-bad-input', 'Choose a supported pen.');
  return pen;
}
function point(value, min, max) {
  if (!Array.isArray(value) || value.length !== 2)
    throw refusal('object-bad-input', 'Ink points need two coordinates.');
  return value.map((n) => integer(n, min, max, 'ink coordinate'));
}
/** Local coordinates include a half-width padding box, so vertical/horizontal paths do not clip. */
export function normalizeInkStroke(value) {
  exact(value, ['schema', 'pen', 'widthEmu', 'heightEmu', 'points'], 'ink stroke');
  if (value.schema !== 1) throw refusal('object-unsupported', 'Unsupported ink stroke schema.');
  const pen = inkPen(value.pen),
    pad = Math.ceil(pen.widthEmu / 2),
    widthEmu = integer(value.widthEmu, 2 * pad, INK_LIMITS.extentEmu, 'stroke width'),
    heightEmu = integer(value.heightEmu, 2 * pad, INK_LIMITS.extentEmu, 'stroke height');
  if (!Array.isArray(value.points) || value.points.length < 2 || value.points.length > INK_LIMITS.points)
    throw refusal('object-bad-input', 'A stroke needs 2–512 distinct consecutive points.');
  const points = value.points.map((p) => point(p, pad, INK_LIMITS.extentEmu)),
    xs = points.map((p) => p[0]),
    ys = points.map((p) => p[1]);
  if (
    points.some(
      (p, i) =>
        p[0] > widthEmu - pad ||
        p[1] > heightEmu - pad ||
        (i > 0 && p[0] === points[i - 1][0] && p[1] === points[i - 1][1])
    ) ||
    Math.min(...xs) !== pad ||
    Math.min(...ys) !== pad ||
    Math.max(...xs) !== widthEmu - pad ||
    Math.max(...ys) !== heightEmu - pad
  )
    throw refusal('object-limit', 'Stroke bounds must tightly contain its padded points.');
  return { schema: 1, pen: pen.id, widthEmu, heightEmu, points };
}
/** Capture worksheet logical pixels once. Zoom/scroll conversion belongs to the view adapter. */
export function inkStrokeFromLogicalPoints(input, penId = 'black') {
  const pen = inkPen(penId),
    pad = Math.ceil(pen.widthEmu / 2);
  if (!Array.isArray(input) || input.length < 2 || input.length > INK_LIMITS.inputPoints)
    throw refusal('object-bad-input', 'A pointer draft needs 2–2,048 input points.');
  const points = [];
  for (const item of input) {
    if (
      !Array.isArray(item) ||
      item.length !== 2 ||
      item.some((n) => typeof n !== 'number' || !Number.isFinite(n) || n < 0 || n > 1_000_000_000)
    )
      throw refusal('internal-invalid-field', 'Invalid worksheet pointer coordinate.');
    const p = item.map((n) => Math.round(n * INK_EMU_PER_PIXEL)),
      last = points.at(-1);
    if (!last || p[0] !== last[0] || p[1] !== last[1]) points.push(p);
  }
  if (points.length < 2 || points.length > INK_LIMITS.points)
    throw refusal('object-bad-input', 'Make a stroke with 2–512 distinct consecutive points.');
  const xs = points.map((p) => p[0]),
    ys = points.map((p) => p[1]),
    left = Math.min(...xs) - pad,
    top = Math.min(...ys) - pad;
  if (left < 0 || top < 0) throw refusal('object-bad-input', 'Keep the stroke inside the worksheet edge.');
  const stroke = normalizeInkStroke({
    schema: 1,
    pen: pen.id,
    widthEmu: Math.max(...xs) - left + pad,
    heightEmu: Math.max(...ys) - top + pad,
    points: points.map(([x, y]) => [x - left, y - top]),
  });
  return { stroke, originEmu: { x: left, y: top } };
}
export function inkStrokePath(value) {
  const stroke = normalizeInkStroke(value);
  return stroke.points.map(([x, y], i) => (i ? 'L' : 'M') + x + ' ' + y).join(' ');
}
/** Whole-path picking groundwork, not a lasso or segment eraser. Coordinates are local EMU. */
export function inkStrokeHit(value, x, y, toleranceEmu = 0) {
  const stroke = normalizeInkStroke(value),
    radius = inkPen(stroke.pen).widthEmu / 2 + integer(toleranceEmu, 0, INK_LIMITS.offsetEmu, 'pick tolerance');
  if (typeof x !== 'number' || typeof y !== 'number' || !Number.isFinite(x) || !Number.isFinite(y))
    throw refusal('internal-invalid-field', 'Invalid stroke pick position.');
  if (x < -radius || y < -radius || x > stroke.widthEmu + radius || y > stroke.heightEmu + radius) return false;
  return stroke.points.slice(1).some((b, i) => {
    const a = stroke.points[i],
      dx = b[0] - a[0],
      dy = b[1] - a[1],
      u = Math.max(0, Math.min(1, ((x - a[0]) * dx + (y - a[1]) * dy) / (dx * dx + dy * dy)));
    return Math.hypot(x - a[0] - u * dx, y - a[1] - u * dy) <= radius;
  });
}

export const isInkStrokeCommand = (c) =>
  ['createInkStroke', 'updateInkStroke', 'deleteInkStroke'].includes(c?.operation);
/** Detach model/placement at queue admission; existing stroke edits only move/update alt text. */
export function captureInkStrokeCommands(commands) {
  if (!Array.isArray(commands) || !commands.length) throw refusal('internal-missing-field', 'Missing stroke commands.');
  const copied = structuredClone(commands);
  for (const c of copied) {
    if (!isInkStrokeCommand(c)) continue;
    const allowed = [
      'operation',
      'sheet',
      'objectId',
      'expectedObjectRevision',
      ...(c.operation === 'createInkStroke'
        ? ['stroke', 'anchor', 'altText']
        : c.operation === 'updateInkStroke'
          ? ['anchor', 'altText']
          : []),
    ];
    if (Object.keys(c).some((k) => !allowed.includes(k)))
      throw refusal('internal-invalid-field', 'Invalid stroke command fields.');
    if (c.operation === 'createInkStroke') c.stroke = normalizeInkStroke(c.stroke);
  }
  if (new TextEncoder().encode(JSON.stringify(copied)).byteLength > 2 * 1024 * 1024)
    throw refusal('object-limit', 'Stroke command batch exceeds its 2 MiB bound.');
  return copied;
}
