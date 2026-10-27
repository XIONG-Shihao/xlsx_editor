import assert from 'node:assert/strict';
import { test } from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/index.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { encodeOwnerResult, decodeOwnerResult, ownerCommands, decodedCommands } from '../core/transport.ts';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const row = (n, cell) => `<row r="${n}">${cell}</row>`;
const number = (n, value) => row(n, `<c r="A${n}"><v>${value}</v></c>`);
const text = (n, value) => row(n, `<c r="A${n}" t="inlineStr"><is><t>${value}</t></is></c>`);
const bool = (n, value) => row(n, `<c r="A${n}" t="b"><v>${value ? 1 : 0}</v></c>`);
const formula = (n, value) => row(n, `<c r="A${n}"><f>${value}</f><v>999</v></c>`);
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function fixture(body, { calculate } = {}) {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-filter-values-'));
  const client = createWorkbookNativeClient(nativeWorkbookBinary()),
    calls = [];
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
  parts['custom/opaque.bin'] = new Uint8Array([0, 1, 255]);
  const source = path.join(temporary, 'source.xlsx'),
    bytes = zipSync(parts);
  await fs.writeFile(source, bytes);
  let root;
  const native = async (input) => {
    if (input.operation === 'index') root = input.root;
    const result = await client.call({ ...input, diagnostics: true });
    calls.push({ operation: input.operation, version: input.version, diagnostics: result.diagnostics });
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('Unexpected writer fallback'),
    (input) => (calculate ? calculate(input, native) : calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
  } catch (error) {
    await book.close();
    client.close();
    await fs.rm(temporary, { recursive: true, force: true });
    throw error;
  }
  return {
    book,
    native,
    calls,
    source,
    bytes,
    root,
    sheet: book.view('0').sheets[0].path,
    close: async () => {
      await book.close();
      client.close();
      await fs.rm(temporary, { recursive: true, force: true });
    },
  };
}
const safe = { errors: 0, unresolved: 0, special: 0, unsafeNumbers: 0 };

test('owner choices retain first-seen typed identities, exact adjacent numbers and calculated values', async () => {
  const f = await fixture(
    text(1, 'Header') +
      number(2, 1) +
      text(3, '1') +
      bool(4, true) +
      text(5, 'true') +
      text(6, '') +
      row(7, '<c r="A7"/>') +
      number(9, '-0') +
      number(10, '0') +
      number(11, '1.0000000000000002') +
      formula(12, '1+1') +
      formula(13, 'TRUE()') +
      formula(14, '&quot;text&quot;') +
      text(15, '1')
  );
  try {
    const page = await f.book.filterValues('0', 'Sheet1', 'A2:A15', 'A');
    assert.deepEqual(
      page.values.map((v) => v.value),
      [1, '1', true, 'true', '', null, -0, 0, 1.0000000000000002, 2, 'text']
    );
    assert.deepEqual(
      page.values.map((v) => v.count),
      [1, 2, 2, 1, 1, 2, 1, 1, 1, 1, 1]
    );
    assert.equal(new Set(page.values.map((v) => v.key)).size, 11);
    assert.equal(page.total, 11);
    assert.equal(page.nextOffset, null);
    assert.deepEqual(page.protected, safe);
    assert.equal(page.values[0].key, 'n:3ff0000000000000');
    assert.equal(page.values[8].key, 'n:3ff0000000000001');
    const sliced = await f.book.filterValues('0', 'Sheet1', 'A2:A15', 'A', '', 3, 2);
    assert.deepEqual(sliced.values, page.values.slice(3, 5));
    assert.equal(sliced.nextOffset, 5);
    const changed = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'A3', value: 'branch' }]);
    const branch = await f.book.filterValues(changed.view.storage.state, 'Sheet1', 'A2:A15', 'A');
    assert.equal(branch.values.find((v) => v.value === '1').count, 1);
    assert.equal(branch.values.find((v) => v.value === 'branch').count, 1);
    assert.deepEqual(await f.book.filterValues('0', 'Sheet1', 'A2:A15', 'A'), page);
    assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
    assert.equal(Object.keys(f.book.view('0').sheets[0].cells).length, 0);
  } finally {
    await f.close();
  }
});

test('protected errors, unresolved results, special formulas and lossy numbers never become ordinary choices', async () => {
  const changes = [
    [6, { formula: '2+2', value: 999, dataType: 'number', calculation: { status: 'calculated', value: 4 } }],
    [7, { formula: '1/0', value: null, calculation: { status: 'error', error: '#DIV/0!' } }],
    [8, { formula: 'A8', value: null, calculation: { status: 'circular', error: '#CYCLE!' } }],
    [9, { formula: 'UNSUPPORTED()', value: 999, calculation: { status: 'unresolved' } }],
    [10, { formula: '1+1', value: 999, calculation: { status: 'calculated' } }],
    [
      11,
      {
        formula: '1+1',
        formulaType: 'array',
        formulaReference: 'A11',
        value: 2,
        calculation: { status: 'calculated', value: 2 },
      },
    ],
    [
      12,
      {
        formula: '1+1',
        formulaType: 'shared',
        sharedFormulaId: '0',
        value: 2,
        calculation: { status: 'calculated', value: 2 },
      },
    ],
    [13, { scenarioResult: true, value: 2, calculation: { status: 'calculated', value: 2 } }],
    [
      15,
      {
        formula: '&quot;resolved&quot;',
        value: '#N/A',
        dataType: 'error',
        calculation: { status: 'calculated', value: 'resolved' },
      },
    ],
  ];
  const f = await fixture(
    text(1, 'Header') +
      row(2, '<c r="A2" t="e"><v>#DIV/0!</v></c>') +
      text(3, '#DIV/0!') +
      number(4, '12345678901234567890') +
      number(5, '1e999') +
      formula(6, '2+2') +
      Array.from({ length: 7 }, (_, i) => number(i + 7, 999)).join('') +
      bool(14, true) +
      number(15, 999),
    {
      calculate: (input, native) =>
        native({
          action: 'workbook',
          operation: 'patch',
          root: input.root,
          version: input.version,
          changes: changes.map(([n, value]) => ({ sheet: 'xl/worksheets/sheet1.xml', cell: `A${n}`, value })),
        }),
    }
  );
  try {
    const page = await f.book.filterValues('0', 'Sheet1', 'A2:A15', 'A');
    assert.deepEqual(
      page.values.map((v) => v.value),
      ['#DIV/0!', 4, true, 'resolved']
    );
    assert.deepEqual(page.protected, { errors: 3, unresolved: 2, special: 3, unsafeNumbers: 2 });
    const searched = await f.book.filterValues('0', 'Sheet1', 'A2:A15', 'A', 'resolved');
    assert.equal(searched.total, 1);
    assert.deepEqual(searched.protected, page.protected, 'Search must not hide protected cells elsewhere in the range');
    assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
  } finally {
    await f.close();
  }
});

test('calculated strings supersede preserved error caches while raw errors remain protected in apply', async () => {
  const calculated = await fixture(
    text(1, 'Header') + row(2, '<c r="A2" t="e"><f>&quot;resolved&quot;</f><v>#N/A</v></c>')
  );
  try {
    const page = await calculated.book.filterValues('0', 'Sheet1', 'A2:A2', 'A');
    assert.deepEqual(
      page.values.map((v) => v.value),
      ['resolved']
    );
    assert.deepEqual(page.protected, safe);
    const next = await calculated.book.apply('0', [
      { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:A2', column: 'A', values: ['resolved'] },
    ]);
    assert.deepEqual(next.view.sheets[0].hiddenRows, []);
    const saved = readWorkbook(await fs.readFile((await calculated.book.export(next.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.A2.cachedValue, 'resolved');
  } finally {
    await calculated.close();
  }
  const raw = await fixture(text(1, 'Header') + row(2, '<c r="A2" t="e"><v>#DIV/0!</v></c>') + text(3, '#DIV/0!'));
  try {
    const page = await raw.book.filterValues('0', 'Sheet1', 'A2:A3', 'A');
    assert.deepEqual(
      page.values.map((v) => v.value),
      ['#DIV/0!']
    );
    assert.equal(page.values[0].count, 1);
    assert.equal(page.protected.errors, 1);
    await assert.rejects(
      raw.book.apply('0', [
        { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:A3', column: 'A', values: ['#DIV/0!'] },
      ]),
      { code: 'filter-unsupported' }
    );
    assert.deepEqual(await raw.book.filterValues('0', 'Sheet1', 'A2:A3', 'A'), page);
  } finally {
    await raw.close();
  }
});

test('large offscreen high-cardinality ranges use bounded pages and arithmetic sparse blank counts', async () => {
  const size = 4207,
    start = 20000,
    last = 1048576;
  const f = await fixture(
    text(1, 'Header') +
      Array.from({ length: size }, (_, i) => number(start + i, i)).join('') +
      text(last, 'offscreen-tail')
  );
  try {
    const first = await f.book.filterValues('0', 'Sheet1', `A2:A${last}`, 'A');
    assert.equal(first.total, size + 2);
    assert.equal(first.values.length, 100);
    assert.equal(first.nextOffset, 100);
    assert.deepEqual(first.values[0], { key: 'null', value: null, count: last - 1 - size - 1 });
    assert.deepEqual(
      first.values.slice(1).map((v) => v.value),
      Array.from({ length: 99 }, (_, i) => i)
    );
    const next = await f.book.filterValues('0', 'Sheet1', `A2:A${last}`, 'A', '', 100);
    assert.deepEqual(
      next.values.map((v) => v.value),
      Array.from({ length: 100 }, (_, i) => i + 99)
    );
    assert(!next.values.some((v) => first.values.some((prior) => prior.key === v.key)));
    const tail = await f.book.filterValues('0', 'Sheet1', `A2:A${last}`, 'A', '', 4200);
    assert.equal(tail.values.length, 9);
    assert.equal(tail.values.at(-1).value, 'offscreen-tail');
    assert.equal(tail.nextOffset, null);
    const found = await f.book.filterValues('0', 'Sheet1', `A2:A${last}`, 'A', 'OFFSCREEN');
    assert.equal(found.total, 1);
    assert.equal(found.values[0].value, 'offscreen-tail');
    for (const call of f.calls.filter((c) => c.operation === 'filter-values')) {
      const counters = call.diagnostics.counters;
      assert.equal(counters['filter_values.stored_cells'], size + 1);
      assert.equal(counters['filter_values.blank_positions'], last - 1 - size - 1);
      assert(counters['filter_values.page_cells'] <= 2048);
      assert.equal(counters['filter_values.pages'], 3);
      assert(counters['filter_values.returned_values'] <= 100);
    }
    assert.deepEqual(first.protected, safe);
    assert.equal(Object.keys(f.book.view('0').sheets[0].cells).length, 0);
  } finally {
    await f.close();
  }
});

test('search uses literal Unicode-aware labels and owner/native inputs enforce their pinned scope', async () => {
  const f = await fixture(text(1, 'Header') + text(2, 'Äpfel') + text(3, 'a%b_c') + text(4, 'other') + number(5, 42));
  try {
    assert.deepEqual(
      (await f.book.filterValues('0', 'Sheet1', 'A2:A5', 'A', 'äP')).values.map((v) => v.value),
      ['Äpfel']
    );
    assert.deepEqual(
      (await f.book.filterValues('0', 'Sheet1', 'A2:A5', 'A', '%b_')).values.map((v) => v.value),
      ['a%b_c']
    );
    assert.equal((await f.book.filterValues('0', 'Sheet1', 'A2:A5', 'A', 'absent')).total, 0);
    assert.deepEqual((await f.book.filterValues('0', 'Sheet1', 'A2:A5', 'A', '', 100)).values, []);
    for (const args of [
      ['missing', 'Sheet1', 'A2:A5', 'A'],
      ['0', 'Missing', 'A2:A5', 'A'],
      ['0', 'Sheet1', 'A2:A5', 'B'],
      ['0', 'Sheet1', 'A2:A5', 'a'],
      ['0', 'Sheet1', 'A5:A2', 'A'],
      ['0', 'Sheet1', 'A2:A5', 'A', '', -1],
      ['0', 'Sheet1', 'A2:A5', 'A', '', 0, 101],
      ['0', 'Sheet1', 'A2:A5', 'A', '', 0, 0],
      ['0', 'Sheet1', 'A2:A5', 'A', '中'.repeat(10923)],
    ])
      await assert.rejects(f.book.filterValues(...args));
    const request = {
      action: 'workbook',
      operation: 'filter-values',
      root: f.root,
      version: 0,
      sheet: f.sheet,
      range: 'A2:A5',
      column: 'A',
    };
    for (const extra of [
      { version: 99999 },
      { sheet: 'xl/worksheets/foreign.xml' },
      { column: 'B' },
      { column: 'A1' },
      { query: null },
      { query: '中'.repeat(10923) },
      { offset: -1 },
      { offset: 9007199254740992 },
      { limit: 101 },
      { limit: 0 },
    ]) {
      await assert.rejects(f.native({ ...request, ...extra }));
    }
    assert.equal((await f.book.filterValues('0', 'Sheet1', 'A2:A5', 'A')).total, 4);
  } finally {
    await f.close();
  }
});

test('native cancellation and TEMP quota refusal preserve accepted data and permit a later query', async () => {
  const size = 30000;
  const f = await fixture(
    text(1, 'Header') + Array.from({ length: size }, (_, i) => text(i + 2, `${i}-` + 'value'.repeat(60))).join('')
  );
  try {
    const request = {
      action: 'workbook',
      operation: 'filter-values',
      root: f.root,
      version: 0,
      sheet: f.sheet,
      range: `A2:A${size + 1}`,
      column: 'A',
    };
    const database = path.join(f.root, 'cells.sqlite'),
      before = digest(await fs.readFile(database)),
      copiedSource = digest(await fs.readFile(path.join(f.root, 'source.xlsx')));
    const pending = f.native(request),
      // Where the cancellation lands varies; a bare database interrupt carries no code.
      refused = assert.rejects(pending, /cancel|interrupt/i),
      deadline = Date.now() + 3000;
    const active = path.join(path.dirname(f.root), 'native-active.json');
    while (
      !(await fs.stat(active).then(
        () => true,
        () => false
      ))
    ) {
      assert(Date.now() < deadline, 'Native query did not enter its lease');
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
    await new Promise((resolve) => setTimeout(resolve, 20));
    await fs.writeFile(path.join(f.root, 'cancelled'), 'test');
    const stop = performance.now();
    await refused;
    assert(performance.now() - stop < 3000);
    await fs.rm(path.join(f.root, 'cancelled'));
    assert.equal(digest(await fs.readFile(database)), before);
    const lease = path.join(path.dirname(f.root), 'owner.json'),
      originalLease = await fs.readFile(lease, 'utf8');
    async function bytesAt(directory) {
      let total = 0;
      for (const name of await fs.readdir(directory)) {
        const file = path.join(directory, name),
          stat = await fs.stat(file);
        total += stat.isDirectory() ? await bytesAt(file) : stat.size;
      }
      return total;
    }
    // Leave enough free bytes to enter the readonly operation, but only ~64 KiB
    // for TEMP pages. Grouping 30,000 distinct long strings must refuse there.
    await fs.writeFile(
      lease,
      JSON.stringify({
        ...JSON.parse(originalLease),
        maxSessionBytes: (await bytesAt(path.dirname(f.root))) + 6 * 65536 + 4096,
      })
    );
    try {
      await assert.rejects(f.native(request), /full|disk/i);
    } finally {
      await fs.writeFile(lease, originalLease);
    }
    assert.equal(digest(await fs.readFile(database)), before);
    assert.equal(digest(await fs.readFile(path.join(f.root, 'source.xlsx'))), copiedSource);
    assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
    const recovered = await f.book.filterValues('0', 'Sheet1', `A2:A${size + 1}`, 'A', '29999-');
    assert.equal(recovered.total, 1);
    assert(recovered.values[0].value.startsWith('29999-'));
    assert.deepEqual(recovered.protected, safe);
    assert.deepEqual(
      (await f.book.filterValues('0', 'Sheet1', 'A2:A2', 'A')).values.map((v) => v.value),
      ['0-' + 'value'.repeat(60)]
    );
  } finally {
    await fs.rm(path.join(f.root, 'cancelled'), { force: true });
    await f.close();
  }
});

test('typed selections survive JSON transport; owned blank and empty criteria remain distinct through saved metadata', async () => {
  const f = await fixture(
    text(1, 'Header') +
      number(2, '-0') +
      number(3, '0') +
      bool(4, true) +
      text(5, 'true') +
      number(6, 1) +
      text(7, '1') +
      text(9, '')
  );
  try {
    const original = await f.book.filterValues('0', 'Sheet1', 'A2:A9', 'A');
    const roundtrip = decodeOwnerResult(
      JSON.parse(JSON.stringify(encodeOwnerResult({ filterValues: original })))
    ).filterValues;
    assert.deepEqual(roundtrip, original);
    assert(Object.is(roundtrip.values[0].value, -0));
    assert(Object.is(roundtrip.values[1].value, 0));
    assert.notEqual(roundtrip.values[0].key, roundtrip.values[1].key);
    const command = (values) => ({ operation: 'filterRows', sheet: 'Sheet1', range: 'A1:A9', column: 'A', values });
    const transmitted = decodedCommands(
      JSON.parse(JSON.stringify(ownerCommands([command([roundtrip.values[0].value])])))
    );
    assert(Object.is(transmitted[0].values[0], -0));
    const zero = await f.book.apply('0', transmitted);
    assert.deepEqual(zero.view.sheets[0].hiddenRows, [2, 3, 4, 5, 6, 7, 8]);
    const savedZero = readWorkbook(await fs.readFile((await f.book.export(zero.view.storage.state)).path));
    assert(Object.is(savedZero.sheets[0].cells.A2.value, -0));
    assert.deepEqual(savedZero.sheets[0].hiddenRows, [2, 3, 4, 5, 6, 7, 8]);
    const boolean = await f.book.apply('0', [command([true])]);
    assert.deepEqual(boolean.view.sheets[0].hiddenRows, [1, 2, 4, 5, 6, 7, 8]);
    const numeric = await f.book.apply('0', [command([1])]);
    assert.deepEqual(numeric.view.sheets[0].hiddenRows, [1, 2, 3, 4, 6, 7, 8]);
    const onlyBlank = await f.book.apply('0', [command([null])]);
    assert.deepEqual(onlyBlank.view.sheets[0].hiddenRows, [1, 2, 3, 4, 5, 6, 8]);
    const onlyEmpty = await f.book.apply('0', [command([''])]);
    assert.deepEqual(onlyEmpty.view.sheets[0].hiddenRows, [1, 2, 3, 4, 5, 6, 7]);
    assert.deepEqual(await f.book.filterValues('0', 'Sheet1', 'A2:A9', 'A'), original);
    const combined = await f.book.apply('0', [command([null, ''])]);
    assert.deepEqual(combined.view.sheets[0].hiddenRows, [1, 2, 3, 4, 5, 6]);
    const bytes = await fs.readFile((await f.book.export(combined.view.storage.state)).path),
      xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
    assert.match(xml, /<filters\b[^>]*blank="1"/);
    assert.equal(readWorkbook(bytes).sheets[0].cells.A9.value, '');
    assert.equal(digest(await fs.readFile(f.source)), digest(f.bytes));
    const edited = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'A3', value: -0 }]);
    const editedPage = await f.book.filterValues(edited.view.storage.state, 'Sheet1', 'A2:A9', 'A');
    assert.equal(editedPage.protected.unsafeNumbers, 1, 'The unqualified numeric overlay marker must remain protected');
    assert.equal(editedPage.values.filter((v) => typeof v.value === 'number' && Object.is(v.value, -0)).length, 1);
    await assert.rejects(f.book.apply(edited.view.storage.state, [command([-0])]), /precision/i);
  } finally {
    await f.close();
  }
});
