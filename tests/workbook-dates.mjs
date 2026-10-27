import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeFormula } from '../core/native-formula.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { indexedCalculationReusable } from '../core/indexed-checkpoint.mjs';
after(() => closeNativeWorkbookWorker());

for (const system of [1900, 1904])
  test(`indexed ${system} dates retain exact typed oracle results through export/reopen`, async (t) => {
    const source = new URL(`../fixtures/excel-date-${system}-source.xlsx`, import.meta.url),
      bytes = await fs.readFile(source);
    const oracle = JSON.parse(
      await fs.readFile(new URL(`../fixtures/excel-date-${system}-reference.json`, import.meta.url), 'utf8')
    );
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-dates-')),
      copy = path.join(root, 'input.xlsx');
    await fs.writeFile(copy, bytes);
    t.after(() => fs.rm(root, { recursive: true, force: true }));
    const book = new IndexedWorkbook(
      native,
      () => assert.fail('No writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
    await book.open(copy);
    t.after(() => book.close());
    async function verify(owner, state) {
      const cells = (await owner.read(state, 'Cases', 'B1:H80')).cells;
      for (const item of oracle.cases) {
        const result = cells[item.ref].calculation;
        if (Number(item.ref.slice(1)) >= 78) {
          assert.equal(result.status, 'unresolved', item.ref);
          continue;
        }
        if (item.dataType === 'error') {
          assert.equal(result.status, 'error', item.ref);
          assert.equal(result.error, item.expected, item.ref);
        } else {
          assert.equal(result.status, 'calculated', item.ref);
          assert.equal(result.value, item.expected, item.ref);
        }
      }
    }
    await verify(book, '0');
    const changed = await book.apply('0', [{ sheet: 'Cases', cell: 'A1', value: 'independent date roundtrip' }]);
    const output = (await book.export(changed.view.storage.state)).path;
    const expected = { 'Cases!A1': { value: 'independent date roundtrip' } };
    // The independent Python `value` check uses exact numeric equality. Numeric
    // cache spellings (e.g. 1900 versus 1900.0) need not equal Excel's serializer.
    for (const item of oracle.cases)
      if (Number(item.ref.slice(1)) < 78)
        expected[`Cases!${item.ref}`] = {
          formula: item.formula,
          ...(item.dataType === 'error' ? { rawValue: item.rawValue, error: item.expected } : { value: item.expected }),
        };
    assert.equal(checkXlsxPackage(copy, output, { targets: { Cases: ['A1'] }, expected }).passed, true);
    assert.equal(readWorkbook(await fs.readFile(output)).sheets[0].date1904, system === 1904);
    const reopened = new IndexedWorkbook(
      native,
      () => assert.fail('No writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
    await reopened.open(output);
    t.after(() => reopened.close());
    await verify(reopened, '0');
    assert.deepEqual(await fs.readFile(copy), bytes);
    assert.deepEqual(await fs.readFile(source), bytes);
  });

test('native numeric date analysis/evaluation and branch edits use exact scalar inputs', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-native-dates-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  for (const date1904 of [false, true]) {
    const input = {
      revision: 'dates',
      sheets: [
        {
          name: 'Cases',
          date1904,
          cells: {
            A1: { value: 60.75, dataType: 'number' },
            B1: { formula: 'YEAR(A1)' },
            C1: { formula: 'MONTH(A1)' },
            D1: { formula: 'DAY(A1)' },
            E1: { formula: 'DATE(1904,1,1)' },
          },
        },
      ],
    };
    const analyzed = await nativeFormula(native, root, 'analyze', { view: input });
    assert.equal(analyzed.supported, true, analyzed.reason);
    const result = await nativeFormula(native, root, 'evaluate', { view: input });
    assert.equal(result.supported, true, result.reason);
    assert.equal(result.results.find((item) => item.ref === 'D1').value, date1904 ? 1 : 29);
    assert.equal(result.results.find((item) => item.ref === 'E1').value, date1904 ? 0 : 1462);
    for (const formula of ['YEAR("2024-01-01")', 'DATE(2024,1201,1)', 'YEAR(TRUE)']) {
      const unsupported = await nativeFormula(native, root, 'evaluate', {
        view: { ...input, sheets: [{ ...input.sheets[0], cells: { A1: { formula } } }] },
      });
      assert.equal(unsupported.supported, false, formula);
      assert.equal(unsupported.results, undefined);
    }
    for (const formula of ['YEAR(-1)', 'DATE(9999,12,32)']) {
      const invalid = await nativeFormula(native, root, 'evaluate', {
        view: { ...input, sheets: [{ ...input.sheets[0], cells: { A1: { formula } } }] },
      });
      assert.equal(invalid.supported, true);
      assert.deepEqual(invalid.results, [{ sheet: 'Cases', ref: 'A1', error: '#NUM!' }]);
    }
    for (const invalid of ['true', 1, null]) {
      const refused = await nativeFormula(native, root, 'evaluate', {
        view: { ...input, sheets: [{ ...input.sheets[0], date1904: invalid }] },
      });
      assert.equal(refused.supported, false);
      assert.match(refused.reason, /date-system metadata/);
      assert.equal(refused.results, undefined);
    }
  }
});

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
async function finiteFixture(t, date1904, { shared = false } = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-date-proofs-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  if (date1904)
    parts['xl/workbook.xml'] = strToU8(
      strFromU8(parts['xl/workbook.xml']).replace(/(<workbook[^>]*>)/, '$1<workbookPr date1904="1"/>')
    );
  const formula = (row) =>
    shared
      ? `<f t="shared" si="0"${row === 1 ? ' ref="B1:B3"' : ''}>${row === 1 ? 'DAY(A1)' : ''}</f>`
      : `<f>DAY(A${row})</f>`;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${[60.75, 59.5, 61]
      .map((serial, index) => {
        const row = index + 1;
        return `<row r="${row}"><c r="A${row}"><v>${serial}</v></c><c r="B${row}">${formula(row)}<v>999</v></c>${row === 1 ? '<c r="C1"><f>YEAR(A1)</f><v>999</v></c><c r="D1"><f>DATE(1904,1,1)</f><v>999</v></c><c r="E1"><f>SUM(B1:B3)</f><v>999</v></c>' : ''}</row>`;
      })
      .join('')}</sheetData></worksheet>`
  );
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source),
    calls = [];
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  t.after(() => client.close());
  const transport = async (request) => {
    const result = await client.call(request);
    calls.push({ request, result });
    return result;
  };
  let last;
  const make = () =>
    new IndexedWorkbook(
      transport,
      () => assert.fail('No writer'),
      async (input) => {
        const result = await calculateIndexedWorkbook(input, transport);
        last = { input, result };
        return result;
      }
    );
  return {
    root,
    source,
    original,
    calls,
    transport,
    make,
    last: () => last,
    options: { cacheIdentity: 'qualified-date-test', cacheOptions: { directory: path.join(root, 'cache') } },
  };
}

test('finite shared dates retain compression, branch values and exact saved caches in both systems', async (t) => {
  for (const system of [false, true]) {
    const f = await finiteFixture(t, system, { shared: true }),
      book = f.make();
    await book.open(f.source);
    t.after(() => book.close());
    const values = system ? [1, 29, 2] : [29, 28, 1],
      original = (await book.read('0', 'Sheet1', 'A1:E3')).cells;
    for (let row = 1; row <= 3; row++) {
      assert.equal(original[`B${row}`].calculation.value, values[row - 1]);
      assert.equal(original[`B${row}`].formulaType, 'shared');
    }
    assert.equal(f.last().result.inputCells, 0);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 61.99 }]),
      branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 7 }]);
    const cells = (await book.read(changed.view.storage.state, 'Sheet1', 'A1:E3')).cells;
    assert.equal(cells.B1.calculation.value, system ? 2 : 1);
    assert.equal(cells.B2.calculation.value, values[1]);
    assert.equal(
      (await book.read(branch.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value,
      values[0] + 7 + values[2]
    );
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, values[0]);
    for (const [state, sharedCount] of [
      [changed.view.storage.state, 3],
      [branch.view.storage.state, 0],
    ]) {
      const output = await book.export(state),
        bytes = await fs.readFile(output.path),
        xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
      assert.equal((xml.match(/t="shared"/g) ?? []).length, sharedCount);
      const before = (await book.read(state, 'Sheet1', 'B1:E3')).cells,
        disk = readWorkbook(bytes).sheets[0].cells;
      for (const ref of ['B1', 'B3', 'C1', 'D1', 'E1'])
        assert.equal(disk[ref].cachedValue, before[ref].calculation.value, `Saved ${ref}`);
      assert.equal(
        disk.B2.formula || disk.B2.formulaType ? disk.B2.cachedValue : disk.B2.value,
        before.B2.calculation?.value ?? before.B2.value
      );
      const reopened = f.make();
      await reopened.open(output.path);
      try {
        const after = (await reopened.read('0', 'Sheet1', 'B1:E3')).cells;
        for (const ref of ['B1', 'B3', 'C1', 'D1', 'E1'])
          assert.equal(after[ref].calculation.value, before[ref].calculation.value, ref);
        assert.equal(after.B2.calculation?.value ?? after.B2.value, before.B2.calculation?.value ?? before.B2.value);
      } finally {
        await reopened.close();
      }
    }
    assert.deepEqual(await fs.readFile(f.source), f.original);
  }
});

test('changed saved dates import exact source fields and remain isolated across calendar changes', async (t) => {
  for (const system of [false, true]) {
    const f = await finiteFixture(t, system),
      saved = path.join(f.root, 'saved.xlsx');
    let book = f.make();
    await book.open(f.source, f.options);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A3', value: 61.99 }]),
      state = changed.view.storage.state;
    const output = await book.export(state);
    assert.equal(book.cachePublication.published, false);
    await fs.copyFile(output.path, saved);
    await book.close();
    book = f.make();
    t.after(() => book.close());
    f.calls.length = 0;
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert(
      f.calls.some(({ request }) => request.operation === 'index'),
      'Changed saved date packages must import their exact source fields'
    );
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, system ? 0 : 1462);
    const accepted = await fs.readFile((await book.export('0')).path);
    // Public source bytes and metadata select a separate cache, even with the
    // same formulas and cell coordinates. Existing restored state stays intact.
    const parts = unzipSync(accepted),
      other = path.join(f.root, 'other-calendar.xlsx');
    const xml = strFromU8(parts['xl/workbook.xml']).replace(/<workbookPr\b[^>]*\/?>(?:<\/workbookPr>)?/g, '');
    parts['xl/workbook.xml'] = strToU8(
      xml.replace(/(<workbook[^>]*>)/, `$1<workbookPr date1904="${system ? '0' : '1'}"/>`)
    );
    await fs.writeFile(other, zipSync(parts));
    const fresh = f.make();
    await fresh.open(other, f.options);
    try {
      assert.equal(fresh.cacheHit, false);
      assert.equal((await fresh.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, system ? 1462 : 0);
      assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, system ? 0 : 1462);
      assert.deepEqual(await fs.readFile((await book.export('0')).path), accepted);
    } finally {
      await fresh.close();
    }
    assert.deepEqual(await fs.readFile(f.source), f.original);
  }
});

test('calendar metadata cannot reuse old native plans or mutate an accepted state on refusal', async (t) => {
  const f = await finiteFixture(t, false),
    book = f.make();
  await book.open(f.source);
  t.after(() => book.close());
  const { input, result } = f.last(),
    version = result.version,
    originalView = structuredClone(input.view),
    changedView = structuredClone(input.view);
  changedView.sheets[0].date1904 = true;
  const accepted = await fs.readFile((await book.export('0')).path),
    count = book.diagnosticStats().retainedStates;
  const status = await indexedCalculationReusable({ root: input.root, version, view: changedView }, f.transport);
  assert.equal(status.qualified, false);
  f.calls.length = 0;
  const next = await calculateIndexedWorkbook(
    {
      ...input,
      version,
      view: changedView,
      revision: 'isolated-calendar-probe',
      changedRanges: [{ sheet: 'Sheet1', range: 'Z99' }],
    },
    f.transport
  );
  assert(
    f.calls.some(
      ({ request, result }) =>
        request.operation === 'dependencies' && request.method === 'select' && result.hit === false
    )
  );
  assert.notEqual(next.version, version);
  const page = await f.transport({
    action: 'workbook',
    operation: 'range',
    root: input.root,
    version: next.version,
    sheet: changedView.sheets[0].path,
    range: 'B1:D1',
    limit: 2048,
  });
  assert.equal(new Map(page.cells).get('B1').calculation.value, 1);
  assert.equal(new Map(page.cells).get('D1').calculation.value, 0);
  for (const invalid of ['true', 1, null]) {
    const bad = structuredClone(originalView);
    bad.sheets[0].date1904 = invalid;
    f.calls.length = 0;
    await assert.rejects(
      calculateIndexedWorkbook(
        { ...input, version, view: bad, changedRanges: [{ sheet: 'Sheet1', range: 'Z99' }] },
        f.transport
      ),
      { code: /^(formula-unsupported|internal-invalid-field)$/ }
    );
    assert.equal(f.calls.length, 0, 'Malformed metadata is refused before native proof lookup or mutation');
  }
  assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 29);
  assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 1462);
  assert.equal(book.diagnosticStats().retainedStates, count);
  assert.deepEqual(await fs.readFile((await book.export('0')).path), accepted);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
