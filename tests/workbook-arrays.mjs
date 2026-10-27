import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook, indexedCalculationMetadata } from '../core/workbook-package.mjs';
import { indexedSpillMetadata } from '../core/indexed-spills.mjs';
import { rangeBounds, cellPosition } from '../core/workbook-commands.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const xml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;');

test('actual Excel qualified SEQUENCE outputs and scalar transformations coexist with protected arrays through indexed export/reopen', async (t) => {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-dynamic-reference.json', import.meta.url), 'utf8')
  );
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-dynamic-oracle-protection-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  for (const [name, digest] of [
    [oracle.source, oracle.sourceSha256],
    [oracle.oracleWorkbook, oracle.oracleWorkbookSha256],
  ]) {
    const bytes = await fs.readFile(new URL(`../fixtures/${name}`, import.meta.url));
    assert.equal(createHash('sha256').update(bytes).digest('hex'), digest);
    const source = path.join(root, name);
    await fs.writeFile(source, bytes);
    const book = new IndexedWorkbook(
      native,
      () => assert.fail('No writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
    await book.open(source);
    try {
      const cells = (await book.read('0', 'Cases', 'A1:G80')).cells;
      const producers = ['A1', 'A6', 'A41', 'A56', 'A60', 'A69'],
        consumers = ['D1', 'D2', 'D6', 'A47'],
        qualified = [...producers, ...consumers];
      for (const ref of Object.keys(oracle.authoredFormulas))
        assert.equal(
          cells[ref].calculation.status,
          qualified.includes(ref) ? (oracle.cells[ref].dataType === 'error' ? 'error' : 'calculated') : 'unresolved',
          `${name}: ${ref}`
        );
      assert.equal(cells.A47.calculation.value, oracle.cells.A47.expected);
      assert.equal(cells.A47.calculation.scalarArray.schema, 1);
      if (name === oracle.oracleWorkbook) {
        assert.equal(cells.B56.value, 2);
        assert.equal(cells.B56.styleId, readWorkbook(bytes).sheets[0].cells.B56.styleId);
        assert.equal(cells.B61.value, 0);
        assert.equal(cells.B42.value, oracle.cells.B42.expected);
      }
      const changed = await book.apply('0', [{ sheet: 'Cases', cell: 'F5', value: 9 }]),
        output = await book.export(changed.view.storage.state),
        saved = await fs.readFile(output.path);
      const originals = readWorkbook(bytes).sheets[0].cells,
        savedCells = readWorkbook(saved).sheets[0].cells,
        groups = [],
        expected = {};
      for (const ref of producers) {
        const range = oracle.cells[ref].formulaAttributes.ref,
          error = oracle.cells[ref].expected === '#SPILL!',
          group = {
            schema: 1,
            sheet: 'xl/worksheets/sheet1.xml',
            master: ref,
            range,
            status: error ? 'error' : 'calculated',
            ...(error ? { error: '#SPILL!', intendedRange: `${ref}:B${cellPosition(ref).row + 2}` } : {}),
          };
        groups.push(group);
        const bounds = rangeBounds(range);
        for (const [cell, value] of Object.entries(oracle.cells)) {
          const p = cellPosition(cell);
          if (
            p.row >= bounds.first.row &&
            p.row <= bounds.last.row &&
            p.col >= bounds.first.col &&
            p.col <= bounds.last.col
          )
            expected[cell] = {
              value: value.expected,
              ...(cell === ref ? { formula: oracle.authoredFormulas[ref] } : {}),
            };
        }
      }
      for (const ref of Object.keys(oracle.authoredFormulas))
        if (!qualified.includes(ref) && originals[ref].formulaType)
          assert.deepEqual(savedCells[ref], originals[ref], `Unqualified special source cache: ${ref}`);
      const metadata = indexedSpillMetadata(bytes, indexedCalculationMetadata(bytes, true), groups),
        specification = path.join(root, `${name}-spill-specification.json`),
        report = fileURLToPath(
          new URL(`../.build/tests/dynamic-protection/${name}-spill-export.json`, import.meta.url)
        );
      await fs.writeFile(
        specification,
        JSON.stringify({
          masters: metadata.masters,
          cells: { 'xl/worksheets/sheet1.xml': expected },
          consumers: {
            'xl/worksheets/sheet1.xml': Object.fromEntries(
              consumers.map((ref) => [
                ref,
                { formula: oracle.authoredFormulas[ref], value: oracle.cells[ref].expected, singleton: true },
              ])
            ),
          },
          literals: { 'xl/worksheets/sheet1.xml': { F5: 9 } },
          uncalculated: {
            'xl/worksheets/sheet1.xml': Object.keys(oracle.authoredFormulas).filter(
              (ref) => !qualified.includes(ref) && !originals[ref].formulaType
            ),
          },
        })
      );
      const checked = spawnSync(
        'python3',
        [
          fileURLToPath(new URL('./support/check-spill-export.py', import.meta.url)),
          source,
          output.path,
          specification,
          report,
        ],
        { encoding: 'utf8' }
      );
      assert.equal(checked.status, 0, checked.stderr);
      assert.equal(JSON.parse(await fs.readFile(report, 'utf8')).sourceSha256, digest);
      assert.equal(savedCells.F5.value, 9);
      const reopened = new IndexedWorkbook(
        native,
        () => assert.fail('No writer'),
        (input) => calculateIndexedWorkbook(input, native)
      );
      await reopened.open(output.path);
      try {
        const after = (await reopened.read('0', 'Cases', 'A1:G80')).cells;
        for (const ref of Object.keys(oracle.authoredFormulas))
          assert.equal(after[ref].calculation.status, cells[ref].calculation.status, `${name}: reopened ${ref}`);
      } finally {
        await reopened.close();
      }
      assert.deepEqual(await fs.readFile(source), bytes);
      assert.equal(
        createHash('sha256')
          .update(await fs.readFile(new URL(`../fixtures/${name}`, import.meta.url)))
          .digest('hex'),
        digest
      );
    } finally {
      await book.close();
    }
  }
});
async function fixture(t, formula, input = 3) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-array-protection-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook()),
    source = path.join(root, 'source.xlsx');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>${input}</v></c><c r="B1"><v>2</v></c><c r="C1"><f>${xml(formula)}</f><v>777</v></c><c r="D1"><f>C2+1</f><v>777</v></c><c r="E1"><f>IFERROR(C1,99)</f><v>777</v></c><c r="F1"><f>SUM(B1:B3)</f><v>777</v></c></row><row r="2"><c r="B2"><v>3</v></c></row><row r="3"><c r="B3"><v>4</v></c></row></sheetData></worksheet>`
  );
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source),
    calls = [];
  let diagnostic;
  const transport = async (request) => {
    calls.push(request);
    return native(request);
  };
  const make = () =>
    new IndexedWorkbook(
      transport,
      () => assert.fail('No writer'),
      async (input) => (diagnostic = await calculateIndexedWorkbook(input, transport))
    );
  return { root, source, original, calls, make, diagnostic: () => diagnostic };
}

test('indexed modern spill masters, absent followers and error-catching dependents remain unresolved through export/reopen', async (t) => {
  for (const formula of ['SEQUENCE(A1)', 'FILTER(B1:B3,B1:B3>0)', 'UNIQUE(B1:B3)', 'IF(TRUE,9,SEQUENCE(A1))']) {
    const f = await fixture(t, formula),
      book = f.make();
    await book.open(f.source);
    t.after(() => book.close());
    const options = { cacheIdentity: 'array-protection-test', cacheOptions: { directory: path.join(f.root, 'cache') } };
    const before = (await book.read('0', 'Sheet1', 'A1:F3')).cells;
    for (const ref of ['C1', 'D1', 'E1'])
      assert.equal(before[ref].calculation.status, 'unresolved', `${formula}: ${ref}`);
    assert.equal(before.F1.calculation.value, 9);
    assert.equal(before.C2, undefined);
    assert.equal(f.diagnostic().nativePartialCells, 0);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]),
      state = changed.view.storage.state,
      branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 8 }]);
    assert.equal((await book.read(state, 'Sheet1', 'D1')).cells.D1.calculation.status, 'unresolved');
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'F1')).cells.F1.calculation.value, 14);
    const output = await book.export(state),
      bytes = await fs.readFile(output.path),
      disk = readWorkbook(bytes).sheets[0].cells;
    assert.equal(disk.C1.formula, formula);
    assert.notEqual(disk.C1.cachedValue, 1);
    // A1 was edited. A formula that reads A1 can no longer vouch for its stored
    // 777, so that result is dropped; one that does not read A1 keeps it.
    assert.equal(disk.C1.cachedValue, formula.includes('A1') ? null : 777, formula);
    assert.equal(disk.C2, undefined);
    assert.equal(disk.F1.cachedValue, 9);
    assert.equal(
      checkXlsxPackage(f.source, output.path, {
        targets: { Sheet1: formula.includes('A1') ? ['A1', 'C1', 'D1'] : ['A1'] },
        expected: {
          'Sheet1!A1': { value: 5 },
          // D1 reads C2, where C1's array would spill, so it follows C1.
          'Sheet1!C1': { formula, rawValue: formula.includes('A1') ? null : '777' },
          'Sheet1!D1': { formula: 'C2+1', rawValue: formula.includes('A1') ? null : '777' },
          'Sheet1!F1': { formula: 'SUM(B1:B3)', value: 9 },
        },
      }).passed,
      true
    );
    const reopened = f.make();
    await reopened.open(output.path, options);
    try {
      const cells = (await reopened.read('0', 'Sheet1', 'A1:F3')).cells;
      for (const ref of ['C1', 'D1', 'E1']) assert.equal(cells[ref].calculation.status, 'unresolved', ref);
      assert.equal(cells.F1.calculation.value, 9);
      await reopened.export('0');
      assert.notEqual(
        reopened.cachePublication?.published,
        true,
        'An unowned spill cannot become a trusted saved checkpoint'
      );
    } finally {
      await reopened.close();
    }
    assert.deepEqual(await fs.readFile(f.source), f.original);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 3);
  }
});

test('indexed conditional geometry growth cannot retain an old certified blank-follower result', async (t) => {
  const f = await fixture(t, 'IF(A1=0,99,B1:B3)', 0),
    book = f.make();
  await book.open(f.source);
  t.after(() => book.close());
  assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 99);
  assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 1);
  let state = '0';
  for (const value of [1, 0, 1, 0]) {
    const changed = await book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', value }]);
    state = changed.view.storage.state;
    const cells = (await book.read(state, 'Sheet1', 'A1:F3')).cells;
    assert.equal(cells.C1.calculation.status, value ? 'unresolved' : 'calculated');
    assert.equal(cells.D1.calculation.status, value ? 'unresolved' : 'calculated');
    if (!value) {
      assert.equal(cells.C1.calculation.value, 99);
      assert.equal(cells.D1.calculation.value, 1);
    }
    assert.equal(cells.F1.calculation.value, 9);
    const output = await book.export(state),
      reopened = f.make();
    await reopened.open(output.path);
    try {
      const after = (await reopened.read('0', 'Sheet1', 'C1:F1')).cells;
      for (const ref of ['C1', 'D1', 'E1'])
        assert.equal(after[ref].calculation.status, cells[ref].calculation.status, ref);
      if (!value) assert.equal(after.D1.calculation.value, 1);
    } finally {
      await reopened.close();
    }
  }
  assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 99);
  assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 1);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});

test('indexed bare arrays and range-valued INDEX cannot certify one scalar result or imaginary followers', async (t) => {
  for (const formula of ['B1:B3', 'B1:B3+1', '{1;2;3}', 'INDEX(B1:B3,0)', 'ROW(B1:B3)', 'MUNIT(A1)']) {
    const f = await fixture(t, formula),
      book = f.make();
    await book.open(f.source);
    try {
      const cells = (await book.read('0', 'Sheet1', 'C1:F3')).cells;
      assert.equal(cells.C1.calculation.status, 'unresolved', formula);
      assert.equal(cells.D1.calculation.status, 'unresolved', formula);
      assert.equal(cells.E1.calculation.status, 'unresolved', formula);
      assert.equal(cells.F1.calculation.value, 9);
      assert.equal(cells.C2, undefined);
      assert.equal(f.diagnostic().nativePartialCells, 0);
    } finally {
      await book.close();
    }
  }
  const f = await fixture(t, 'MUNIT(A1)', 1),
    book = f.make();
  await book.open(f.source);
  try {
    assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 1);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 1);
    const grown = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 2 }]),
      cells = (await book.read(grown.view.storage.state, 'Sheet1', 'C1:D1')).cells;
    assert.equal(cells.C1.calculation.status, 'unresolved');
    assert.equal(cells.D1.calculation.status, 'unresolved');
    assert.equal(f.diagnostic().nativePartialCells, 0);
    const shrunk = await book.apply(grown.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 1 }]);
    assert.equal((await book.read(shrunk.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 1);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 1);
    assert.deepEqual(await fs.readFile(f.source), f.original);
  } finally {
    await book.close();
  }
});
