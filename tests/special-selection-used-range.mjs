import './support/scratch-dir.mjs';
/** Actual Special collector with controlled owner extent/source replies. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import {
  providerReceipt,
  initialCells,
  completeUsedRange,
  incompleteUsedRange,
  sourceReader,
} from './support/used-range-selection-fixture.mjs';

const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(
    path.join(component, '.build/tests/special-selection-used-range-model-')
  );
const bundle = path.join(output, 'model.mjs');
await build({
  entryPoints: [path.join(component, 'ui/specialSelection.ts')],
  bundle: true,
  format: 'esm',
  platform: 'node',
  outfile: bundle,
  logLevel: 'silent',
});
const api = await import(pathToFileURL(bundle));
const binding = {
  documentId: 'doc',
  ownerIdentity: 'standalone-registry-alias',
  state: 'registry:accepted-head',
  sheet: 'Sheet1',
  selection: 'A5',
  selectedCell: 'A5',
  activationEpoch: 0,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
};
const options = (kind) => ({ kind, types: { number: true, text: true, logical: true, error: true } }),
  defer = () => {
    let resolve;
    const promise = new Promise((yes) => (resolve = yes));
    return { promise, resolve };
  };
const collect = (cells = initialCells(), patch = {}) =>
  api.collectSpecialSelectionPreview({
    binding: structuredClone(binding),
    options: options('constants'),
    geometry: { name: 'Sheet1', merges: [] },
    read: sourceReader(cells),
    isCurrent: () => true,
    usedRange: completeUsedRange(),
    usedRangeReceipt: { ...providerReceipt },
    ...patch,
  });
const check = (name, fn) => {
  test(name, async () => {
    await fn();
  });
};

check(
  'Single-cell Constants, Formulas and Blanks inspect UsedRange with observed exact matches and active cursors',
  async () => {
    const expected = {
      constants: { references: ['A1', 'C4', 'A5', 'C5', 'C6'], activeCell: 'A1' },
      formulas: { references: ['A3', 'C3'], activeCell: 'A3' },
      blanks: { references: ['B1', 'C1', 'A2', 'B2', 'C2', 'B3', 'A4', 'B4', 'B5', 'A6', 'B6'], activeCell: 'A2' },
    };
    for (const [kind, value] of Object.entries(expected)) {
      const calls = [],
        preview = await collect(initialCells(), { options: options(kind), read: sourceReader(initialCells(), calls) });
      assert.deepEqual(preview.references, value.references, kind);
      assert.equal(preview.activeCell, value.activeCell, kind);
      assert.equal(preview.range, 'A1:C6');
      assert.equal(preview.positions, 18);
      assert.equal(preview.binding.selection, 'A5');
      assert.equal(preview.scope.kind, 'used-range');
      assert.deepEqual(calls, [{ sheet: 'Sheet1', range: 'A1:C6' }]);
      assert.equal(api.specialSelectionPreviewPage(preview).total, value.references.length);
    }
  }
);
check('Original provider authority remains independent from standalone Home owner and state aliases', async () => {
  const preview = await collect();
  assert.equal(preview.binding.ownerIdentity, 'standalone-registry-alias');
  assert.equal(preview.binding.state, 'registry:accepted-head');
  assert.deepEqual(preview.scope.receipt, providerReceipt);
  assert.deepEqual(preview.scope.usedRange.receipt, providerReceipt);
  assert.equal(preview.scope.usedRange.native.version, 2);
  assert.equal(preview.scope.usedRange.native.sheetPath, 'xl/worksheets/original-path.xml');
  assert(
    Object.isFrozen(preview.scope) &&
      Object.isFrozen(preview.scope.usedRange.native) &&
      Object.isFrozen(preview.scope.receipt)
  );
});
check('Explicit rectangles and Visible single-cell scope do not borrow UsedRange', async () => {
  const explicit = await collect(initialCells(), {
    binding: { ...binding, selection: 'A1:B2' },
    usedRange: undefined,
    usedRangeReceipt: undefined,
  });
  assert.equal(explicit.scope.kind, 'selection');
  assert.equal(explicit.range, 'A1:B2');
  assert.deepEqual(explicit.references, ['A1']);
  const visible = await collect(initialCells(), {
    options: options('visible'),
    visibility: { name: 'Sheet1', widths: [] },
    usedRange: undefined,
    usedRangeReceipt: undefined,
  });
  assert.equal(visible.range, 'A5');
  assert.deepEqual(visible.references, ['A5']);
  assert.equal(visible.scope.kind, 'selection');
  assert.equal(api.specialSelectionNeedsUsedRange('A5', { name: 'Sheet1', merges: [] }, 'visible'), false);
  await assert.rejects(
    collect(initialCells(), { binding: { ...binding, selection: 'A1:B2' } }),
    /explicit.*extent authority/
  );
});
check(
  'Single full merged areas default to UsedRange; partial single-cell merge selections close before classification',
  async () => {
    const geometry = { name: 'Sheet1', merges: ['G8:I10'] },
      cells = { C10: { value: 9, dataType: 'number' } },
      usedRange = completeUsedRange('C8:I10', { merges: 'G8:I10' });
    for (const selection of ['G8:I10', 'I10', 'H9']) {
      assert(api.specialSelectionNeedsUsedRange(selection, geometry, 'constants'));
      const preview = await collect(cells, { binding: { ...binding, selection }, geometry, usedRange });
      assert.equal(preview.scope.kind, 'used-range');
      assert.equal(preview.range, 'C8:I10');
      assert.deepEqual(preview.references, ['C10']);
      assert.equal(preview.activeCell, 'C10');
    }
    assert.equal(api.specialSelectionNeedsUsedRange('C8:I10', geometry, 'constants'), false);
  }
);
check(
  'Last Cell uses the worksheet corner for explicit selections and reads no cells even with full worksheet extent',
  async () => {
    let reads = 0;
    const read = async () => {
      reads++;
      throw Error('Last Cell must not read cells');
    };
    for (const [usedRange, expected] of [
      [completeUsedRange(), 'C6'],
      [completeUsedRange('A1:Z100'), 'Z100'],
      [completeUsedRange('A1:XFD1048576'), 'XFD1048576'],
    ]) {
      const preview = await collect(
        {},
        { binding: { ...binding, selection: 'A1:B2' }, options: options('last-cell'), usedRange, read }
      );
      assert.equal(preview.scope.kind, 'last-cell');
      assert.equal(preview.range, expected);
      assert.equal(preview.activeCell, expected);
      assert.equal(preview.positions, 1);
      assert.deepEqual(preview.references, [expected]);
      assert.deepEqual(preview.readRanges, []);
      assert.equal(api.specialSelectionPreviewPage(preview).total, 1);
    }
    assert.equal(reads, 0);
  }
);
check('Last Cell raw merged corner expands the complete area and activates its anchor without cell reads', async () => {
  let reads = 0;
  const preview = await collect(
    {},
    {
      options: options('last-cell'),
      usedRange: completeUsedRange('C8:I10', { merges: 'G8:I10' }),
      geometry: { name: 'Sheet1', merges: ['G8:I10'] },
      read: async () => {
        reads++;
        throw Error('No source read');
      },
    }
  );
  assert.equal(preview.scope.usedRange.lastCell, 'I10');
  assert.equal(preview.range, 'G8:I10');
  assert.deepEqual(preview.references, ['G8', 'H8', 'I8', 'G9', 'H9', 'I9', 'G10', 'H10', 'I10']);
  assert.equal(preview.activeCell, 'G8');
  assert.equal(preview.positions, 9);
  assert.equal(reads, 0);
  assert.equal(api.specialSelectionPreviewPage(preview).total, 9);
});
check('Empty worksheet Last Cell selects A1 and source authorship is not manufactured', async () => {
  const usedRange = completeUsedRange('A1', { empty: true }),
    preview = await collect(
      {},
      {
        options: options('last-cell'),
        usedRange,
        read: async () => {
          throw Error('No cells');
        },
      }
    );
  assert.deepEqual(preview.references, ['A1']);
  assert.equal(preview.scope.usedRange.empty, true);
  assert.equal(preview.scope.usedRange.provenance.effectiveCells, null);
  assert.equal(preview.activeCell, 'A1');
});
check('Missing, incomplete, foreign and retired UsedRange authority refuses before any source cell read', async () => {
  let reads = 0;
  const read = async () => {
      reads++;
      throw Error('No cells');
    },
    valid = completeUsedRange();
  const cases = [
    { usedRange: undefined },
    { usedRangeReceipt: undefined },
    { usedRange: incompleteUsedRange() },
    { usedRange: { ...valid, sheet: 'Other' } },
    { usedRange: { ...valid, receipt: { ...providerReceipt, ownerId: 'foreign' } } },
    { usedRange: { ...valid, receipt: { ...providerReceipt, state: 'other-head' } } },
    { usedRange: { ...valid, receipt: { ...providerReceipt, revision: 'other-revision' } } },
    { binding: { ...binding, documentId: 'foreign' } },
    { usedRange: { ...valid, native: { ...valid.native, sourceRevision: 'not-a-hash' } } },
    { usedRange: { ...valid, lastCell: 'C5' } },
    { usedRange: { ...valid, scopeMinMax: { ...valid.scopeMinMax, maxRow: 4 } } },
  ];
  for (const patch of cases)
    await assert.rejects(collect({}, { ...patch, read }), /UsedRange|document|source|geometry|bounds/);
  assert.equal(reads, 0);
});
check('1004 UsedRange targets use bounded complete windows and immutable100-row pages', async () => {
  const cells = Object.fromEntries(
      Array.from({ length: 1004 }, (_, i) => ['A' + (i + 1), { value: i, dataType: 'number' }])
    ),
    calls = [],
    preview = await collect(cells, { usedRange: completeUsedRange('A1:A1004'), read: sourceReader(cells, calls) });
  assert.deepEqual(calls, [
    { sheet: 'Sheet1', range: 'A1:A1000' },
    { sheet: 'Sheet1', range: 'A1001:A1004' },
  ]);
  assert.equal(preview.positions, 1004);
  const all = [];
  let offset = 0,
    pages = 0;
  do {
    const page = api.specialSelectionPreviewPage(preview, offset);
    all.push(...page.rows.map((row) => row.cell));
    pages++;
    assert(Object.isFrozen(page.rows));
    if (page.nextOffset === null) break;
    offset = page.nextOffset;
  } while (true);
  assert.equal(pages, 11);
  assert.deepEqual(all, Object.keys(cells));
});
check(
  'UsedRange materialization and last-cell merged closure enforce10k without reading over-limit scopes',
  async () => {
    let reads = 0;
    const read = async () => {
      reads++;
      throw Error('No cells');
    };
    await assert.rejects(collect({}, { usedRange: completeUsedRange('A1:A10001'), read }), /10,000/);
    await assert.rejects(
      collect(
        {},
        {
          options: options('last-cell'),
          usedRange: completeUsedRange('A1:CW100', { merges: 'A1:CW100' }),
          geometry: { name: 'Sheet1', merges: ['A1:CW100'] },
          read,
        }
      ),
      /10,000/
    );
    assert.equal(reads, 0);
    const cells = { A10000: { value: 1, dataType: 'number' } },
      calls = [],
      preview = await collect(cells, { usedRange: completeUsedRange('A1:A10000'), read: sourceReader(cells, calls) });
    assert.equal(preview.positions, 10000);
    assert.equal(calls.length, 10);
    assert.deepEqual(preview.references, ['A10000']);
  }
);
check(
  'Preview copies cannot forge extent scope, provider identity, native geometry, corner or merged active anchor',
  async () => {
    const preview = await collect();
    const changes = [
      (copy) => (copy.scope.kind = 'selection'),
      (copy) => (copy.scope.receipt.ownerId = 'foreign'),
      (copy) => (copy.scope.usedRange.native.version = -1),
      (copy) => (copy.scope.usedRange.native.sourceRevision = 'x'),
      (copy) => (copy.scope.usedRange.sheet = 'Other'),
      (copy) => (copy.scope.usedRange.range = 'A1:C7'),
      (copy) => (copy.range = 'A5'),
      (copy) => (copy.positions = 1),
      (copy) => (copy.readRanges = ['A5:A5']),
      (copy) => delete copy.scope,
    ];
    for (const change of changes) {
      const copy = structuredClone(preview);
      change(copy);
      assert.throws(() => api.specialSelectionPreviewPage(copy));
    }
    const last = await collect(
      {},
      {
        options: options('last-cell'),
        usedRange: completeUsedRange('C8:I10', { merges: 'G8:I10' }),
        geometry: { name: 'Sheet1', merges: ['G8:I10'] },
      }
    );
    for (const change of [
      (copy) => (copy.activeCell = 'I10'),
      (copy) => (copy.references = copy.references.slice(1)),
      (copy) => (copy.readRanges = ['G8:I10']),
      (copy) => (copy.scope.usedRange.lastCell = 'G8'),
    ]) {
      const copy = structuredClone(last);
      change(copy);
      assert.throws(() => api.specialSelectionPreviewPage(copy));
    }
  }
);
check(
  'Captured extent and receipt cannot retarget a held read; cancellation and late retirement mask results',
  async () => {
    const held = defer(),
      usedRange = completeUsedRange(),
      receipt = { ...providerReceipt },
      input = {
        binding: { ...binding },
        options: options('constants'),
        geometry: { name: 'Sheet1', merges: [] },
        usedRange,
        usedRangeReceipt: receipt,
        read: async (_sheet, range) => ({ range, complete: true, cells: await held.promise }),
        isCurrent: () => true,
      };
    const pending = api.collectSpecialSelectionPreview(input);
    usedRange.receipt.ownerId = 'changed';
    receipt.state = 'changed';
    usedRange.native.version = 999;
    held.resolve(initialCells());
    const preview = await pending;
    assert.deepEqual(preview.scope.receipt, providerReceipt);
    assert.equal(preview.scope.usedRange.native.version, 2);
    const gate = defer(),
      controller = new AbortController(),
      cancelled = collect(
        {},
        {
          signal: controller.signal,
          read: async (_sheet, range) => ({ range, complete: true, cells: await gate.promise }),
        }
      );
    controller.abort();
    gate.resolve({});
    await assert.rejects(cancelled, (error) => error.name === 'AbortError');
    let current = true;
    const gate2 = defer(),
      retired = collect(
        {},
        {
          isCurrent: () => current,
          read: async (_sheet, range) => ({ range, complete: true, cells: await gate2.promise }),
        }
      );
    current = false;
    gate2.resolve({});
    await assert.rejects(retired, /initiating workbook/);
  }
);
