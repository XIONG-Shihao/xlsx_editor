import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { normalizeSortKeys } from '../core/sort-keys.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';
import { projectWorkbookCommands } from '../core/workbook-projection.mjs';
import { createWorkbook, readWorkbook, editWorkbook, applyProjectedOperation } from '../core/workbook-package.mjs';
import * as engine from '../core/calculation-node.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';

const roots = [],
  clients = [],
  ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function check(name, fn) {
  test(name, async (t) => {
    const value = await fn(t);
  });
}
after(async () => {
  for (const client of clients) await client.close();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
const key = (column, direction = 'asc') => ({ column, direction });
const command = (keys) => ({ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:E7', keys });
const mixedKeys = [key('A'), key('B', 'desc'), key('C')];
const values = [
  [2, 'B', 1],
  [1, 'a', 1],
  [1, 'b', 2],
  [1, 'b', 1],
  [1, 'b', 1],
  [2, 'a', 3],
  [1, 'a', 0],
];
const expectedOrder = [4, 5, 3, 7, 2, 1, 6];
function mixedBytes() {
  const parts = unzipSync(createWorkbook());
  parts['custom/ordered-sort.bin'] = new Uint8Array([7, 0, 255, 17]);
  const rows = values
    .map(([a, b, c], i) => {
      const r = i + 1;
      return `<row r="${r}" ht="${20 + r}" customHeight="1"${r === 5 ? ' hidden="1"' : ''}><c r="A${r}"><v>${a}</v></c><c r="B${r}" t="inlineStr"><is><t>${b}</t></is></c><c r="C${r}"><v>${c}</v></c><c r="D${r}" t="inlineStr"><is><t>row-${r}</t></is></c><c r="E${r}"><f>A${r}*100+C${r}+$F$12</f><v>999</v></c></row>`;
    })
    .join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${rows}<row r="12"><c r="F12"><v>10</v></c></row></sheetData></worksheet>`
  );
  return editWorkbook(
    zipSync(parts),
    values.map((_, i) => ({ sheet: 'Sheet1', cell: 'D' + (i + 1), fontColor: '00000' + (i + 1), bold: i % 2 === 0 }))
  );
}
async function fixture(bytes) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ordered-sort-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, bytes);
  return { root, source };
}
async function owner(source) {
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  clients.push(client);
  const calls = [];
  const native = (input) => {
    calls.push(structuredClone(input));
    return client.call(input);
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('Native sort used the whole-workbook fallback'),
    (request) => calculateIndexedWorkbook(request, native)
  );
  await book.open(source);
  return { book, calls, native };
}
function verifyMixed(actual, original) {
  const sheet = actual.sheets[0],
    old = original.sheets[0];
  assert.deepEqual(
    expectedOrder.map((_, i) => sheet.cells['D' + (i + 1)].value),
    expectedOrder.map((r) => 'row-' + r)
  );
  for (const [from, i] of expectedOrder.map((r, i) => [r, i])) {
    const to = i + 1;
    assert.equal(sheet.cells['D' + to].fontColor, old.cells['D' + from].fontColor);
    assert.equal(sheet.cells['D' + to].bold, old.cells['D' + from].bold);
    assert.equal(sheet.cells['E' + to].formula, `A${to}*100+C${to}+$F$12`);
    assert.equal(sheet.cells['A' + to].value, old.cells['A' + from].value);
    assert.equal(sheet.cells['C' + to].value, old.cells['C' + from].value);
    assert.equal(sheet.rowHeights.find((r) => r.row === i).height, 20 + from);
  }
  assert.deepEqual(sheet.hiddenRows, [1]);
  assert.equal(sheet.cells.F12.value, 10);
}

check('ordered keys close shape, range, uniqueness and 16-key bounds while retaining legacy', async () => {
  const view = readWorkbook(mixedBytes()),
    valid = command(mixedKeys);
  assert.deepEqual(normalizeSortKeys(valid), mixedKeys);
  assert(Object.isFrozen(normalizeSortKeys(valid)[0]));
  assert.deepEqual(resolveWorkbookEdits(view, [valid])[0], valid);
  assert.deepEqual(normalizeSortKeys({ range: 'A1:E7', column: 'B', direction: 'desc' }), [key('B', 'desc')]);
  assert.deepEqual(normalizeSortKeys({ range: '$a$1:$e$7', column: 'b', direction: 'desc' }), [key('B', 'desc')]);
  const invalid = [
    { keys: [] },
    { keys: Array.from({ length: 17 }, (_, i) => key(String.fromCharCode(65 + i))) },
    { keys: [key('A'), key('a', 'desc')] },
    { keys: [key('F')] },
    { keys: [key('$A')] },
    { keys: [key('XFE')] },
    { keys: [{ column: 'A', direction: 'asc', extra: true }] },
    { keys: [{ column: 'A' }] },
    { keys: [{ column: 'A', direction: 'ascending' }] },
    { keys: [null] },
    { keys: [key('A')], column: 'A' },
    { keys: [key('A')], direction: undefined },
    { keys: undefined },
  ];
  for (const patch of invalid) {
    const bad = { ...valid, ...patch };
    assert.throws(() => normalizeSortKeys(bad));
    assert.throws(() => resolveWorkbookEdits(view, [bad]));
  }
  assert.equal(
    normalizeSortKeys({ range: 'A1:P2', keys: Array.from({ length: 16 }, (_, i) => key(String.fromCharCode(65 + i))) })
      .length,
    16
  );
  return { invalidShapes: invalid.length, keys: mixedKeys };
});
check('portable three-key mixed directions preserve full row associations and exact writer envelope', async () => {
  const bytes = mixedBytes(),
    original = readWorkbook(bytes),
    before = structuredClone(original),
    out = await projectWorkbookCommands(original, [command(mixedKeys)], engine);
  assert.deepEqual(original, before);
  verifyMixed(out.view, original);
  assert.equal(out.edits.length, 1);
  const saved = applyProjectedOperation(bytes, out.edits[0], out.view).bytes,
    actual = readWorkbook(saved);
  verifyMixed(actual, original);
  const beforeParts = unzipSync(bytes),
    afterParts = unzipSync(saved);
  for (const name of Object.keys(beforeParts))
    if (name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(afterParts[name], beforeParts[name], name);
  const forged = structuredClone(out.view);
  forged.sheets[0].rowOrder[0].to = 20;
  assert.throws(() => applyProjectedOperation(bytes, out.edits[0], forged), /permutation/);
  return { expectedOrder, oneProjectedEdit: true, unrelatedZipMembersExact: true };
});
check(
  'native three-key sort is one transform and preserves formulas, caches, marker styles and row attributes',
  async () => {
    const bytes = mixedBytes(),
      original = readWorkbook(bytes),
      f = await fixture(bytes),
      { book, calls } = await owner(f.source);
    try {
      calls.length = 0;
      const result = await book.apply('0', [command(mixedKeys)]);
      assert.equal(result.edits.length, 1);
      assert.equal(calls.filter((c) => c.operation === 'transform').length, 1);
      assert.equal(calls.filter((c) => c.operation === 'range').length, 0);
      const accepted = (await book.read(result.view.storage.state, 'Sheet1', 'A1:E7')).cells;
      assert.deepEqual(
        expectedOrder.map((_, i) => accepted['D' + (i + 1)].value),
        expectedOrder.map((r) => 'row-' + r)
      );
      const exported = await fs.readFile((await book.export(result.view.storage.state)).path),
        actual = readWorkbook(exported);
      verifyMixed(actual, original);
      for (let r = 1; r <= 7; r++)
        assert.equal(
          actual.sheets[0].cells['E' + r].cachedValue,
          actual.sheets[0].cells['A' + r].value * 100 + actual.sheets[0].cells['C' + r].value + 10
        );
      assert.deepEqual(unzipSync(exported)['custom/ordered-sort.bin'], unzipSync(bytes)['custom/ordered-sort.bin']);
      assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.value, 'row-1');
      assert.deepEqual(await fs.readFile(f.source), bytes);
      return { expectedOrder, transforms: 1, acceptedState: result.view.storage.state, originalStateRetained: true };
    } finally {
      await book.close();
    }
  }
);
check('native and portable preserve typed order, blanks-last per key and stable complete ties', async () => {
  const keys = [key('A'), key('B', 'desc')],
    rows = [
      [1, 2, 'n1'],
      [1, 2, 'n2'],
      [1, true, 'n-bool'],
      [1, null, 'n-blank'],
      [2, false, 'n-two'],
      ['x', 1, 'text'],
      [false, 1, 'false'],
      [true, 1, 'true'],
      ['#N/A', 1, 'error-na'],
      ['#DIV/0!', 1, 'error-div'],
      [null, 1, 'blank-low'],
      ['', 2, 'empty-high'],
    ];
  const parts = unzipSync(createWorkbook());
  const cell = (ref, v) =>
    v === null
      ? ''
      : typeof v === 'number'
        ? `<c r="${ref}"><v>${v}</v></c>`
        : typeof v === 'boolean'
          ? `<c r="${ref}" t="b"><v>${Number(v)}</v></c>`
          : v.startsWith('#')
            ? `<c r="${ref}" t="e"><v>${v}</v></c>`
            : `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${rows.map(([a, b, tag], i) => `<row r="${i + 1}">${cell('A' + (i + 1), a)}${cell('B' + (i + 1), b)}${cell('C' + (i + 1), tag)}</row>`).join('')}</sheetData></worksheet>`
  );
  const bytes = zipSync(parts),
    edit = { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:C12', keys },
    expected = [
      'n-bool',
      'n1',
      'n2',
      'n-blank',
      'n-two',
      'text',
      'false',
      'true',
      'error-na',
      'error-div',
      'empty-high',
      'blank-low',
    ],
    portable = await projectWorkbookCommands(readWorkbook(bytes), [edit], engine);
  assert.deepEqual(
    expected.map((_, i) => portable.view.sheets[0].cells['C' + (i + 1)].value),
    expected
  );
  const f = await fixture(bytes),
    { book } = await owner(f.source);
  try {
    const result = await book.apply('0', [edit]),
      cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C12')).cells;
    assert.deepEqual(
      expected.map((_, i) => cells['C' + (i + 1)].value),
      expected
    );
    return { expected, stableEqualRows: ['n1', 'n2'], typedErrorTie: ['error-na', 'error-div'] };
  } finally {
    await book.close();
  }
});
check('native malformed key requests refuse before changing source or accepted rows', async () => {
  const bytes = mixedBytes(),
    f = await fixture(bytes),
    { book, calls, native } = await owner(f.source);
  try {
    calls.length = 0;
    await book.apply('0', [command(mixedKeys)]);
    const request = calls.find((c) => c.operation === 'transform'),
      before = await book.read('0', 'Sheet1', 'A1:E7'),
      artifact = await fs.readFile(path.join(request.root, 'transform.xlsx'));
    const invalid = [
      { keys: [] },
      { keys: [key('A'), key('A')] },
      { keys: [key('F')] },
      { keys: [{ column: 'A', direction: 'asc', extra: 1 }] },
      { keys: [key('A')], column: 'A' },
      { keys: [key('A')], direction: 'asc' },
      { keys: Array.from({ length: 17 }, () => key('A')) },
    ];
    for (const patch of invalid) {
      await assert.rejects(native({ ...request, ...patch }));
      assert.deepEqual(await book.read('0', 'Sheet1', 'A1:E7'), before);
      assert.deepEqual(await fs.readFile(path.join(request.root, 'transform.xlsx')), artifact);
      assert.deepEqual(await fs.readFile(f.source), bytes);
    }
    for (const field of ['column', 'direction']) {
      calls.length = 0;
      await assert.rejects(book.apply('0', [{ ...command(mixedKeys), [field]: undefined }]), /mixed/);
      assert.equal(calls.filter((c) => ['export', 'transform'].includes(c.operation)).length, 0);
      assert.deepEqual(await book.read('0', 'Sheet1', 'A1:E7'), before);
    }
    return { nativeRefusals: invalid.length, indexedPreExportAmbiguityRefusals: 2, sourceAndArtifactExact: true };
  } finally {
    await book.close();
  }
});
check('legacy and keys single-sort routes yield the same native rows and admission', async () => {
  const f = await fixture(mixedBytes()),
    { book } = await owner(f.source);
  try {
    const legacy = { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:E7', column: 'C', direction: 'desc' },
      a = await book.apply('0', [legacy]),
      b = await book.apply('0', [{ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:E7', keys: [key('C', 'desc')] }]);
    const cells = async (result) =>
      Object.fromEntries(
        Object.entries((await book.read(result.view.storage.state, 'Sheet1', 'A1:E7')).cells).map(([r, c]) => [
          r,
          { value: c.value, formula: c.formula, styleId: c.styleId, calculated: c.calculation?.value },
        ])
      );
    assert.deepEqual(await cells(a), await cells(b));
  } finally {
    await book.close();
  }
  const parts = unzipSync(createWorkbook()),
    keys = Array.from({ length: 16 }, (_, i) => key(String.fromCharCode(65 + i), i === 15 ? 'desc' : 'asc'));
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${[1, 2].map((r) => `<row r="${r}">${keys.map((k, i) => `<c r="${k.column}${r}"><v>${i === 15 ? r : 0}</v></c>`).join('')}<c r="Q${r}" t="inlineStr"><is><t>row-${r}</t></is></c></row>`).join('')}</sheetData></worksheet>`
  );
  const g = await fixture(zipSync(parts)),
    other = await owner(g.source);
  try {
    other.calls.length = 0;
    const result = await other.book.apply('0', [{ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:Q2', keys }]),
      cells = (await other.book.read(result.view.storage.state, 'Sheet1', 'Q1:Q2')).cells;
    assert.deepEqual([cells.Q1.value, cells.Q2.value], ['row-2', 'row-1']);
    assert.equal(other.calls.filter((c) => c.operation === 'transform').length, 1);
    return { legacyEquivalent: true, sixteenthKeyDeterminesNativeOrder: true, maxKeys: 16 };
  } finally {
    await other.book.close();
  }
});
