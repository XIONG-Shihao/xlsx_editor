import { FORMAT } from './format.mjs';
/** Closed GLB 2.0 sample assets and genuine mesh projection; no bitmap surrogate. */
import { refusal } from './refusal-codes.mjs';

const deepFreeze = (value) => {
  if (value && typeof value === 'object') {
    Object.values(value).forEach(deepFreeze);
    Object.freeze(value);
  }
  return value;
};
export const MODEL3D_LIMITS = Object.freeze({ assetBytes: 64 * 1024, vertices: 512, triangles: 256, coordinate: 10 });
export const MODEL3D_SAMPLES = deepFreeze([
  {
    id: 'cube',
    label: 'Cube',
    positions: [
      [-1, -1, -1],
      [1, -1, -1],
      [1, 1, -1],
      [-1, 1, -1],
      [-1, -1, 1],
      [1, -1, 1],
      [1, 1, 1],
      [-1, 1, 1],
    ],
    faces: [
      [4, 5, 6],
      [4, 6, 7],
      [1, 0, 3],
      [1, 3, 2],
      [5, 1, 2],
      [5, 2, 6],
      [0, 4, 7],
      [0, 7, 3],
      [7, 6, 2],
      [7, 2, 3],
      [0, 1, 5],
      [0, 5, 4],
    ],
  },
  {
    id: 'pyramid',
    label: 'Pyramid',
    positions: [
      [-1, -1, -1],
      [1, -1, -1],
      [1, -1, 1],
      [-1, -1, 1],
      [0, 1, 0],
    ],
    faces: [
      [0, 1, 4],
      [1, 2, 4],
      [2, 3, 4],
      [3, 0, 4],
      [0, 3, 2],
      [0, 2, 1],
    ],
  },
  {
    id: 'octahedron',
    label: 'Octahedron',
    positions: [
      [1, 0, 0],
      [-1, 0, 0],
      [0, 1, 0],
      [0, -1, 0],
      [0, 0, 1],
      [0, 0, -1],
    ],
    faces: [
      [2, 0, 4],
      [2, 4, 1],
      [2, 1, 5],
      [2, 5, 0],
      [3, 4, 0],
      [3, 1, 4],
      [3, 5, 1],
      [3, 0, 5],
    ],
  },
]);
const fields = (value, allowed, label) => {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((k) => !allowed.includes(k))
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields');
};
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
export function normalizeModel3D(value) {
  fields(value, ['sample', 'color', 'rotation'], '3D model');
  if (!MODEL3D_SAMPLES.some((s) => s.id === value.sample))
    throw refusal('object-bad-input', 'Choose Cube, Pyramid or Octahedron');
  const color = value.color ?? '217346';
  if (typeof color !== 'string' || !/^[0-9A-F]{6}$/.test(color))
    throw refusal('object-bad-input', 'Choose a six-digit uppercase sRGB material color');
  const rotation = value.rotation ?? { x: 20, y: -30, z: 0 };
  fields(rotation, ['x', 'y', 'z'], '3D rotation');
  const angles = {};
  for (const axis of ['x', 'y', 'z']) {
    const angle = rotation[axis];
    if (!Number.isSafeInteger(angle) || angle < -180 || angle > 180)
      throw refusal('object-refused', '3D rotation uses whole degrees from -180 to 180');
    angles[axis] = angle === 0 ? 0 : angle;
  }
  return { sample: value.sample, color, rotation: angles };
}
const sub = (a, b) => a.map((v, i) => v - b[i]),
  cross = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]],
  dot = (a, b) => a.reduce((sum, v, i) => sum + v * b[i], 0),
  unit = (v) => {
    const n = Math.hypot(...v);
    if (!Number.isFinite(n) || n < 1e-9) throw refusal('object-malformed', 'Degenerate 3D triangle');
    return v.map((x) => x / n);
  };
const linear = (c) => {
  const s = c / 255;
  return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
};
const srgb = (c) => Math.round(255 * (c <= 0.0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - 0.055));
function sampleMesh(id) {
  const sample = MODEL3D_SAMPLES.find((s) => s.id === id);
  if (!sample) throw refusal('object-unsupported', 'Unsupported 3D sample');
  const positions = [],
    normals = [],
    triangles = [];
  for (const face of sample.faces) {
    let points = face.map((i) => sample.positions[i]),
      normal = unit(cross(sub(points[1], points[0]), sub(points[2], points[0]))),
      center = points[0].map((_, i) => points.reduce((n, p) => n + p[i], 0) / 3);
    if (dot(normal, center) < 0) {
      points = [points[0], points[2], points[1]];
      normal = normal.map((v) => -v);
    }
    const start = positions.length;
    positions.push(...points.map((p) => [...p]));
    normals.push(...points.map(() => [...normal]));
    triangles.push([start, start + 1, start + 2]);
  }
  return { positions, normals, triangles };
}
function quaternion(rotation) {
  const [x, y, z] = ['x', 'y', 'z'].map((k) => (rotation[k] * Math.PI) / 360),
    sx = Math.sin(x),
    cx = Math.cos(x),
    sy = Math.sin(y),
    cy = Math.cos(y),
    sz = Math.sin(z),
    cz = Math.cos(z);
  // Rz * Ry * Rx, applied to column vectors in the right-handed glTF scene.
  return [
    sx * cy * cz - cx * sy * sz,
    cx * sy * cz + sx * cy * sz,
    cx * cy * sz - sx * sy * cz,
    cx * cy * cz + sx * sy * sz,
  ].map((n) => (n === 0 ? 0 : n));
}
export function model3DGlb(value) {
  const model = normalizeModel3D(value),
    mesh = sampleMesh(model.sample),
    count = mesh.positions.length,
    positionBytes = count * 12,
    normalBytes = count * 12,
    indexBytes = mesh.triangles.length * 6,
    rawLength = positionBytes + normalBytes + indexBytes,
    binLength = Math.ceil(rawLength / 4) * 4,
    bin = new Uint8Array(binLength),
    dv = new DataView(bin.buffer);
  for (const [base, vectors] of [
    [0, mesh.positions],
    [positionBytes, mesh.normals],
  ])
    vectors.forEach((v, i) => v.forEach((n, k) => dv.setFloat32(base + i * 12 + k * 4, n, true)));
  mesh.triangles.flat().forEach((n, i) => dv.setUint16(positionBytes + normalBytes + i * 2, n, true));
  const bounds = (axis) => mesh.positions.map((p) => p[axis]),
    min = [0, 1, 2].map((i) => Math.min(...bounds(i))),
    max = [0, 1, 2].map((i) => Math.max(...bounds(i))),
    rgba = [0, 2, 4].map((i) => linear(Number.parseInt(model.color.slice(i, i + 2), 16))).concat(1);
  const json = {
    asset: { version: '2.0', generator: FORMAT.CLOSED_MODEL3D_SAMPLES_V1 },
    scene: 0,
    scenes: [{ nodes: [0] }],
    nodes: [{ mesh: 0, rotation: quaternion(model.rotation) }],
    meshes: [
      {
        name: MODEL3D_SAMPLES.find((s) => s.id === model.sample).label,
        primitives: [{ attributes: { POSITION: 0, NORMAL: 1 }, indices: 2, material: 0, mode: 4 }],
      },
    ],
    materials: [
      {
        pbrMetallicRoughness: { baseColorFactor: rgba, metallicFactor: 0, roughnessFactor: 1 },
        alphaMode: 'OPAQUE',
        doubleSided: false,
      },
    ],
    buffers: [{ byteLength: rawLength }],
    bufferViews: [
      { buffer: 0, byteOffset: 0, byteLength: positionBytes, target: 34962 },
      { buffer: 0, byteOffset: positionBytes, byteLength: normalBytes, target: 34962 },
      { buffer: 0, byteOffset: positionBytes + normalBytes, byteLength: indexBytes, target: 34963 },
    ],
    accessors: [
      { bufferView: 0, componentType: 5126, count, type: 'VEC3', min, max },
      { bufferView: 1, componentType: 5126, count, type: 'VEC3' },
      {
        bufferView: 2,
        componentType: 5123,
        count: mesh.triangles.length * 3,
        type: 'SCALAR',
        min: [0],
        max: [count - 1],
      },
    ],
    extras: { [FORMAT.MODEL3D_84DD49]: model },
  };
  const text = new TextEncoder().encode(JSON.stringify(json)),
    jsonLength = Math.ceil(text.length / 4) * 4,
    bytes = new Uint8Array(12 + 8 + jsonLength + 8 + binLength),
    out = new DataView(bytes.buffer);
  out.setUint32(0, 0x46546c67, true);
  out.setUint32(4, 2, true);
  out.setUint32(8, bytes.length, true);
  out.setUint32(12, jsonLength, true);
  out.setUint32(16, 0x4e4f534a, true);
  bytes.fill(32, 20, 20 + jsonLength);
  bytes.set(text, 20);
  out.setUint32(20 + jsonLength, binLength, true);
  out.setUint32(24 + jsonLength, 0x004e4942, true);
  bytes.set(bin, 28 + jsonLength);
  return bytes;
}
export function parseModel3DGlb(bytes) {
  if (!(bytes instanceof Uint8Array) || bytes.length < 28 || bytes.length > MODEL3D_LIMITS.assetBytes)
    throw refusal('object-limit', '3D GLB exceeds the supported asset bound');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    view.getUint32(0, true) !== 0x46546c67 ||
    view.getUint32(4, true) !== 2 ||
    view.getUint32(8, true) !== bytes.length
  )
    throw refusal('internal-invalid-field', 'Invalid GLB 2.0 header');
  const jsonLength = view.getUint32(12, true),
    binaryHeader = 20 + jsonLength;
  if (jsonLength % 4 || jsonLength < 4 || binaryHeader + 8 > bytes.length || view.getUint32(16, true) !== 0x4e4f534a)
    throw refusal('internal-invalid-field', 'Invalid GLB JSON chunk');
  const binLength = view.getUint32(binaryHeader, true),
    binStart = binaryHeader + 8;
  if (binLength % 4 || binStart + binLength !== bytes.length || view.getUint32(binaryHeader + 4, true) !== 0x004e4942)
    throw refusal('internal-invalid-field', 'Invalid GLB binary chunk');
  let document;
  try {
    document = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(20, binaryHeader)));
  } catch {
    throw refusal('internal-invalid-field', 'Invalid GLB JSON');
  }
  const model = normalizeModel3D(document?.extras[FORMAT.MODEL3D_84DD49]),
    canonical = model3DGlb(model),
    expectedView = new DataView(canonical.buffer),
    expectedJsonLength = expectedView.getUint32(12, true),
    expectedDocument = JSON.parse(new TextDecoder().decode(canonical.subarray(20, 20 + expectedJsonLength))),
    expectedBin = canonical.subarray(28 + expectedJsonLength);
  // JSON object order/whitespace are immaterial; all declared fields and binary data must match the closed sample.
  const canonicalJson = (v) =>
    v && typeof v === 'object'
      ? Array.isArray(v)
        ? v.map(canonicalJson)
        : Object.fromEntries(
            Object.keys(v)
              .sort()
              .map((k) => [k, canonicalJson(v[k])])
          )
      : v;
  if (
    !same(canonicalJson(document), canonicalJson(expectedDocument)) ||
    expectedBin.length !== binLength ||
    !expectedBin.every((b, i) => b === bytes[binStart + i])
  )
    throw refusal('object-stale', 'Foreign, extended, textured or changed 3D mesh is protected');
  const binary = new DataView(bytes.buffer, bytes.byteOffset + binStart, binLength),
    readVectors = (accessor) => {
      const a = document.accessors[accessor],
        buffer = document.bufferViews[a.bufferView];
      return Array.from({ length: a.count }, (_, i) =>
        [0, 1, 2].map((k) => binary.getFloat32(buffer.byteOffset + i * 12 + k * 4, true))
      );
    },
    positions = readVectors(0),
    normals = readVectors(1),
    indices = document.accessors[2],
    offset = document.bufferViews[indices.bufferView].byteOffset,
    triangles = Array.from({ length: indices.count / 3 }, (_, i) =>
      [0, 1, 2].map((k) => binary.getUint16(offset + (i * 3 + k) * 2, true))
    );
  return { model, mesh: { positions, normals, triangles } };
}
export function model3DSampleAsset(sample) {
  return parseModel3DGlb(model3DGlb({ sample }));
}
function rotate(v, rotation) {
  let [x, y, z] = v;
  const [a, b, c] = ['x', 'y', 'z'].map((k) => (rotation[k] * Math.PI) / 180),
    ca = Math.cos(a),
    sa = Math.sin(a),
    cb = Math.cos(b),
    sb = Math.sin(b),
    cc = Math.cos(c),
    sc = Math.sin(c);
  [y, z] = [y * ca - z * sa, y * sa + z * ca];
  [x, z] = [x * cb + z * sb, -x * sb + z * cb];
  return [x * cc - y * sc, x * sc + y * cc, z];
}
export function projectModel3D(asset, width = 320, height = 240) {
  fields(asset, ['model', 'mesh'], '3D asset');
  const model = normalizeModel3D(asset.model);
  fields(asset.mesh, ['positions', 'normals', 'triangles'], '3D mesh');
  const expected = parseModel3DGlb(model3DGlb(model)).mesh;
  if (!['positions', 'normals', 'triangles'].every((key) => same(asset.mesh[key], expected[key])))
    throw refusal('object-bad-input', '3D preview requires the decoded closed mesh');
  if (![width, height].every((n) => Number.isFinite(n) && n >= 1 / 9525 && n <= 50000000 / 9525))
    throw refusal('object-unsupported', 'Unsupported 3D viewport size');
  const vertices = asset.mesh.positions.map((v) => rotate(v, model.rotation)),
    radius = Math.max(...asset.mesh.positions.map((p) => Math.hypot(...p))),
    scale = (Math.min(width, height) * 0.42) / radius,
    light = unit([-0.3, 0.7, 1]),
    base = [0, 2, 4].map((i) => linear(Number.parseInt(model.color.slice(i, i + 2), 16))),
    faces = [];
  asset.mesh.triangles.forEach((triangle, index) => {
    const points = triangle.map((i) => vertices[i]),
      normal = unit(rotate(asset.mesh.normals[triangle[0]], model.rotation));
    if (normal[2] <= 1e-9) return;
    const brightness = 0.35 + 0.65 * Math.max(0, dot(normal, light)),
      color =
        '#' +
        base
          .map((c) =>
            srgb(Math.min(1, c * brightness))
              .toString(16)
              .padStart(2, '0')
          )
          .join('');
    faces.push({
      index,
      points: points.map((p) => [width / 2 + p[0] * scale, height / 2 - p[1] * scale]),
      depth: points.reduce((n, p) => n + p[2], 0) / 3,
      color,
    });
  });
  faces.sort((a, b) => a.depth - b.depth || a.index - b.index);
  return { width, height, faces };
}
