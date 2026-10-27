import test from 'node:test';
import assert from 'node:assert/strict';
import { collectCurrentRegion, CURRENT_REGION_LIMITS } from '../ui/currentRegion.ts';
import {
  collectSpecialSelectionPreview,
  specialSelectionPreviewPage,
  DEFAULT_SPECIAL_SELECTION_OPTIONS,
} from '../ui/specialSelection.ts';
const position = (ref) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return { row: +m[2] - 1, column: [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1 };
};
const ref = (r, c) => {
  let s = '';
  for (let n = c + 1; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s + (r + 1);
};
function reader(data, trace = []) {
  return async (sheet, range) => {
    assert.equal(sheet, 'Data');
    const [a, z = a] = range.split(':'),
      first = position(a),
      last = position(z);
    assert((last.row - first.row + 1) * (last.column - first.column + 1) <= 1000);
    trace.push(range);
    return {
      range,
      complete: true,
      cells: Object.fromEntries(
        Object.entries(data).filter(([id]) => {
          const p = position(id);
          return p.row >= first.row && p.row <= last.row && p.column >= first.column && p.column <= last.column;
        })
      ),
    };
  };
}
const options = (data) => ({
  sheet: 'Data',
  anchor: 'B2',
  merges: [],
  read: reader(data),
  check() {},
  validateCell: (x) => x,
});
const block = {
  A1: { value: 'Header' },
  B1: { value: 0 },
  C1: { value: false },
  A2: { value: '#N/A', dataType: 'error' },
  B2: { value: 2 },
  C2: { value: '', dataType: 'text' },
  A3: { value: null, formula: 'IF(TRUE,"",1)' },
  C3: { value: 3 },
  A5: { value: 'separate' },
  E2: { value: 'separate' },
};
const binding = {
  documentId: 'doc',
  ownerIdentity: 'owner',
  state: 'head',
  sheet: 'Data',
  selection: 'B2',
  selectedCell: 'B2',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
test('plain data block retains holes, typed zero/false/error/formula/empty Text and stops at blank separators', async () => {
  const before = structuredClone(block),
    trace = [],
    found = await collectCurrentRegion({ ...options(block), read: reader(block, trace) });
  assert.equal(found.range, 'A1:C3');
  assert.equal(found.anchor, 'B2');
  assert.equal(found.positions, 9);
  assert.deepEqual(found.references, ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']);
  assert.deepEqual(block, before);
  assert(trace.length > 0);
});
test('Special preview captures the active cell and a complete regular rectangle with paged inspection', async () => {
  const preview = await collectSpecialSelectionPreview({
    binding,
    options: { ...DEFAULT_SPECIAL_SELECTION_OPTIONS, kind: 'current-region' },
    geometry: { name: 'Data', merges: [] },
    read: reader(block),
    isCurrent: () => true,
  });
  assert.equal(preview.scope.kind, 'current-region');
  assert.equal(preview.range, 'A1:C3');
  assert.equal(preview.activeCell, 'B2');
  assert.equal(specialSelectionPreviewPage(preview).total, 9);
  assert.throws(
    () => specialSelectionPreviewPage({ ...preview, references: preview.references.slice(1) }),
    /Current Region/
  );
});
test('worksheet last row/column edges and single-cell regions have valid bounded reads', async () => {
  for (const anchor of ['A1', 'XFD1048576']) {
    const found = await collectCurrentRegion({ ...options({ [anchor]: { value: false } }), anchor });
    assert.equal(found.range, anchor);
    assert.equal(found.positions, 1);
  }
});
test('long ordinary blocks use adaptive source windows without enumerating the whole worksheet', async () => {
  const data = Object.fromEntries(Array.from({ length: 1000 }, (_, r) => [ref(r, 1), { value: r }])),
    trace = [];
  const found = await collectCurrentRegion({ ...options(data), read: reader(data, trace) });
  assert.equal(found.range, 'B1:B1000');
  assert(trace.length < 300, 'adaptive bands reduce owner round trips for the long block');
  assert(found.references.length === 1000);
});
test('merged, grouped, blank-anchor and diagonal-only blocks have explicit limitations', async () => {
  await assert.rejects(() => collectCurrentRegion({ ...options(block), merges: ['A1:B1'] }), /merged/);
  await assert.rejects(
    () => collectCurrentRegion(options({ ...block, B2: { value: 2, formulaType: 'array' } })),
    /array/
  );
  await assert.rejects(() => collectCurrentRegion(options({})), /populated/);
  await assert.rejects(() => collectCurrentRegion(options({ B2: { value: 2 }, C3: { value: 3 } })), /Diagonal/);
});
test('oversized regions, incomplete replies and foreign cells refuse without a partial result', async () => {
  const data = Object.fromEntries(
    Array.from({ length: 101 }, (_, r) => Array.from({ length: 100 }, (_, c) => [ref(r, c), { value: 1 }])).flat()
  );
  await assert.rejects(() => collectCurrentRegion(options(data)), /10,000/);
  await assert.rejects(
    () => collectCurrentRegion({ ...options(block), read: async () => ({ range: 'B2', complete: false, cells: {} }) }),
    /complete/
  );
  await assert.rejects(
    () =>
      collectCurrentRegion({
        ...options(block),
        read: async (_, range) => ({ range, complete: true, cells: { Z99: { value: 1 } } }),
      }),
    /foreign/
  );
  assert.equal(CURRENT_REGION_LIMITS.positions, 10000);
});
test('cancelled and retired original owners stop before publishing a selection', async () => {
  let current = true;
  await assert.rejects(
    () =>
      collectCurrentRegion({
        ...options(block),
        check() {
          if (!current) throw Error('retired owner');
        },
        read: async (...args) => {
          const result = await reader(block)(...args);
          current = false;
          return result;
        },
      }),
    /retired/
  );
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(
    () =>
      collectSpecialSelectionPreview({
        binding,
        options: { ...DEFAULT_SPECIAL_SELECTION_OPTIONS, kind: 'current-region' },
        geometry: { name: 'Data', merges: [] },
        read: reader(block),
        isCurrent: () => true,
        signal: controller.signal,
      }),
    /cancelled/
  );
});
