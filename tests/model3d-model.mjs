import './support/scratch-dir.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import {
  MODEL3D_SAMPLES,
  model3DGlb,
  normalizeModel3D,
  parseModel3DGlb,
  projectModel3D,
} from '../core/model3d-model.mjs';
const phase = path.resolve('.build/tests/in-m3-model3d-foundation-phase-1');
await fs.mkdir(phase, { recursive: true });
test('three current GLB assets decode actual distinct closed meshes and pass an independent format/geometry oracle', async () => {
  const files = [];
  for (const sample of MODEL3D_SAMPLES) {
    const file = path.resolve('core/assets/models/' + sample.id + '.glb'),
      bytes = await fs.readFile(file),
      asset = parseModel3DGlb(bytes);
    const current = model3DGlb({ sample: sample.id });
    assert.deepEqual(parseModel3DGlb(current), asset);
    assert.deepEqual(bytes, Buffer.from(current), 'Bundled GLB samples use the current neutral metadata');
    assert.equal(asset.mesh.triangles.length, sample.faces.length);
    assert.equal(asset.model.sample, sample.id);
    files.push(file);
  }
  const result = spawnSync('python3', [path.resolve('tests/support/check-model3d-glb.py'), ...files], {
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, result.stderr);
  console.log(result.stdout.trim());
});
test('rotation/color are detached and produce real changed GLB quaternions/material and mesh projection', async () => {
  const files = [];
  for (const sample of MODEL3D_SAMPLES) {
    const source = { sample: sample.id, color: 'DC2626', rotation: { x: -35, y: 70, z: 25 } },
      model = normalizeModel3D(source),
      bytes = model3DGlb(model),
      asset = parseModel3DGlb(bytes),
      before = projectModel3D(parseModel3DGlb(model3DGlb({ sample: sample.id }))),
      after = projectModel3D(asset);
    source.rotation.x = 0;
    assert.equal(model.rotation.x, -35);
    assert.equal(asset.model.color, 'DC2626');
    assert.notDeepEqual(after.faces, before.faces);
    assert(after.faces.length > 0);
    assert(after.faces.every((f) => f.points.flat().every(Number.isFinite)));
    const file = path.join(phase, sample.id + '-edited.glb');
    await fs.writeFile(file, bytes);
    files.push(file);
  }
  const result = spawnSync('python3', [path.resolve('tests/support/check-model3d-glb.py'), ...files], {
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, result.stderr);
});
test('truncated/version/length/type/binary mutation and foreign model fields cannot become a guessed mesh', () => {
  const bytes = model3DGlb({ sample: 'cube' }),
    mutate = (offset, value) => {
      const copy = bytes.slice();
      new DataView(copy.buffer).setUint32(offset, value, true);
      return copy;
    };
  for (const bad of [
    bytes.subarray(0, 27),
    mutate(4, 1),
    mutate(8, bytes.length + 4),
    mutate(12, 0xffffffff),
    mutate(16, 0),
    Uint8Array.from({ length: 65537 }),
  ])
    assert.throws(() => parseModel3DGlb(bad));
  const changed = bytes.slice();
  changed[changed.length - 12] ^= 1;
  assert.throws(() => parseModel3DGlb(changed));
  for (const bad of [
    { sample: 'sphere' },
    { sample: 'cube', color: 'javascript:red' },
    { sample: 'cube', rotation: { x: 0, y: 0, z: Infinity } },
    { sample: 'cube', rotation: { x: 181, y: 0, z: 0 } },
    { sample: 'cube', rotation: { x: 0.5, y: 0, z: 0 } },
    { sample: 'cube', uri: 'https://example.com/a.glb' },
  ])
    assert.throws(() => normalizeModel3D(bad));
});
test('all declared poses and viewport bounds keep real front faces finite and do not modify the decoded asset', () => {
  for (const sample of MODEL3D_SAMPLES)
    for (const rotation of [
      { x: 0, y: 0, z: 0 },
      { x: 0, y: 90, z: 0 },
      { x: 20, y: -30, z: 0 },
      { x: 180, y: 180, z: 180 },
    ]) {
      const asset = parseModel3DGlb(model3DGlb({ sample: sample.id, rotation })),
        before = structuredClone(asset);
      for (const [w, h] of [
        [1 / 9525, 1 / 9525],
        [1, 1],
        [320, 240],
        [50000000 / 9525, 50000000 / 9525],
      ]) {
        const frame = projectModel3D(asset, w, h);
        assert(frame.faces.length > 0);
        assert(frame.faces.every((f) => f.points.flat().every(Number.isFinite) && /^#[0-9a-f]{6}$/.test(f.color)));
      }
      assert.deepEqual(asset, before);
    }
  const asset = parseModel3DGlb(model3DGlb({ sample: 'cube' }));
  assert.throws(() => projectModel3D(asset, NaN, 240));
  asset.mesh.positions[0][0] = 9;
  assert.throws(() => projectModel3D(asset));
});
