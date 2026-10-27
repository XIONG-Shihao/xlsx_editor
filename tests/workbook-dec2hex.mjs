import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeFormula } from '../core/native-formula.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const escape = (value) => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;');
// Authored -0 retains the existing global numeric-fidelity guard. Its actual
// text0 result is still checked in the full compatibility/mixed oracle tests.
const nativeRefs = new Set([
  'E1',
  'E4',
  'E7',
  'E10',
  'E13',
  'E16',
  'E19',
  'E22',
  'E25',
  'E70',
  'E76',
  'E79',
  'E82',
  'E88',
  'E91',
  'E94',
  'E97',
  'E100',
  'E103',
  'E109',
  'E118',
]);
for (const row of [1, 4, 7, 10, 13, 16]) for (const column of ['F', 'G', 'H']) nativeRefs.add(`${column}${row}`);
async function oracleFixture() {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-dec2hex-reference.json', import.meta.url), 'utf8')
  );
  const bytes = await fs.readFile(new URL(`../fixtures/${oracle.source}`, import.meta.url));
  assert.equal(hash(bytes), oracle.sourceSha256);
  assert.equal(
    hash(await fs.readFile(new URL(`../fixtures/${oracle.oracleWorkbook}`, import.meta.url))),
    oracle.oracleWorkbookSha256
  );
  return { oracle, bytes, view: readWorkbook(bytes) };
}
function projection(view, refs = nativeRefs) {
  return {
    ...view,
    revision: 'qualified-dec2hex',
    sheets: view.sheets.map((sheet) => ({
      ...sheet,
      cells: Object.fromEntries(
        Object.entries(sheet.cells).filter(([ref, cell]) => (!cell.formula && !cell.formulaType) || refs.has(ref))
      ),
    })),
  };
}
async function directory(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-dec2hex-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  return root;
}
async function owner(t, source) {
  let diagnostic;
  const calls = [];
  const transport = async (request) => {
    calls.push(request);
    return native(request);
  };
  const book = new IndexedWorkbook(
    transport,
    () => assert.fail('No writer'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, transport))
  );
  await book.open(source);
  t.after(() => book.close());
  return { book, calls, diagnostic: () => diagnostic };
}
async function scalarPackage(t, view) {
  const root = await directory(t),
    source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook()),
    rows = new Map();
  for (const [ref, cell] of Object.entries(view.sheets[0].cells)) {
    const row = Number(ref.match(/\d+$/)[0]);
    let content;
    if (cell.formula) content = `<c r="${ref}"><f>${escape(cell.formula)}</f><v>777</v></c>`;
    else if (cell.dataType === 'error') content = `<c r="${ref}" t="e"><v>${escape(cell.value)}</v></c>`;
    else if (cell.value == null) content = `<c r="${ref}"/>`;
    else if (typeof cell.value === 'boolean') content = `<c r="${ref}" t="b"><v>${Number(cell.value)}</v></c>`;
    else if (typeof cell.value === 'string')
      content = `<c r="${ref}" t="inlineStr"><is><t>${escape(cell.value)}</t></is></c>`;
    else content = `<c r="${ref}"><v>${cell.rawValue ?? cell.value}</v></c>`;
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push(content);
  }
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${[...rows]
      .sort((a, b) => a[0] - b[0])
      .map(([row, cells]) => `<row r="${row}">${cells.join('')}</row>`)
      .join('')}</sheetData></worksheet>`
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace('name="Sheet1"', 'name="Cases"')
      .replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="InputNumber">Cases!$B$1</definedName></definedNames>'
      )
  );
  await fs.writeFile(source, zipSync(parts));
  return { root, source, original: await fs.readFile(source) };
}

test('native DEC2HEX returns exact typed strings, dependents and dependency edges from the actual Excel oracle', async (t) => {
  const root = await directory(t),
    { oracle, view } = await oracleFixture(),
    input = projection(view);
  const analyzed = await nativeFormula(native, root, 'analyze', { view: input });
  assert.equal(analyzed.supported, true, analyzed.reason);
  assert(analyzed.items.find((item) => item.ref === 'E19').ranges.some((range) => range.ref === 'B1'));
  assert(analyzed.items.find((item) => item.ref === 'E118').ranges.some((range) => range.ref === 'B1'));
  const result = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(result.supported, true, result.reason);
  assert.equal(result.results.length, 39);
  for (const item of result.results) {
    assert.equal(item.value, oracle.cells[item.ref].expected, item.ref);
    assert.equal(typeof item.value, typeof oracle.cells[item.ref].expected, item.ref);
  }
});

test('outside-domain, coercion, places, precision and array cases miss atomically instead of inventing errors', async (t) => {
  const root = await directory(t),
    { oracle, view } = await oracleFixture(),
    arrays = new Set(oracle.sourceArrayMasters);
  for (const ref of Object.keys(oracle.authoredFormulas))
    if (!nativeRefs.has(ref) && !arrays.has(ref)) {
      const result = await nativeFormula(native, root, 'evaluate', {
        view: projection(view, new Set([...nativeRefs, ref])),
      });
      if (result.supported) {
        for (const item of result.results) {
          const key = oracle.cells[item.ref];
          assert(key, item.ref);
          assert.equal(key.dataType === 'error' ? item.error : item.value, key.expected, item.ref);
        }
      } else assert.equal(result.results, undefined, ref);
    }
  for (const formula of [
    'DEC2HEX(-1)',
    'IFERROR(DEC2HEX(-1),"wrong")',
    'DEC2HEX(255,8)',
    'DEC2HEX(B1:B2)',
    'DEC2HEX({1,2})',
    'DEC2HEX(1e309)',
  ]) {
    const result = await nativeFormula(native, root, 'evaluate', {
      view: {
        revision: 'outside-subset',
        sheets: [{ name: 'Cases', cells: { B1: { value: 1 }, B2: { value: 2 }, C1: { formula } } }],
      },
    });
    assert.equal(result.supported, false, formula);
    assert.equal(result.results, undefined);
  }
  const precision = {
    revision: 'unsafe-precision',
    sheets: [
      {
        name: 'Cases',
        cells: {
          B1: { value: 255, rawValue: '255.00000000000000001', dataType: 'number' },
          C1: { formula: 'DEC2HEX(B1)' },
        },
      },
    ],
  };
  assert.equal((await nativeFormula(native, root, 'evaluate', { view: precision })).supported, false);
});

test('native scalar workbook avoids provider first load and saves exact string caches without changing authored spellings', async (t) => {
  const { oracle, view } = await oracleFixture(),
    input = projection(view),
    f = await scalarPackage(t, input);
  const probe = spawnSync(
    process.execPath,
    [fileURLToPath(new URL('./support/dec2hex-native-probe.mjs', import.meta.url)), f.source],
    { encoding: 'utf8', env: process.env, timeout: 120000 }
  );
  assert.equal(probe.status, 0, probe.stderr);
  const proof = JSON.parse(probe.stdout);
  assert.deepEqual(proof.before, []);
  assert.deepEqual(proof.after, []);
  assert(proof.diagnostic.inputCells <= 1);
  for (const ref of nativeRefs) assert.equal(proof.results[ref].value, oracle.cells[ref].expected, ref);
  const fOwner = await owner(t, f.source);
  assert(fOwner.diagnostic().inputCells <= 1);
  const changed = await fOwner.book.apply('0', [
      { sheet: 'Cases', cell: 'B1', value: 16 },
      { sheet: 'Cases', cell: 'E13', bold: true },
    ]),
    state = changed.view.storage.state;
  assert.equal((await fOwner.book.read(state, 'Cases', 'E19')).cells.E19.calculation.value, oracle.cells.E10.expected);
  assert.equal((await fOwner.book.read('0', 'Cases', 'E19')).cells.E19.calculation.value, oracle.cells.E19.expected);
  const output = await fOwner.book.export(state),
    bytes = await fs.readFile(output.path),
    disk = readWorkbook(bytes).sheets[0].cells,
    xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
  for (const ref of ['E1', 'E13', 'E16', 'E19', 'E109', 'E118']) {
    const cellXml = new RegExp(`<c\\b[^>]*r="${ref}"[^>]*>[\\s\\S]*?</c>`).exec(xml)?.[0];
    assert.match(cellXml, /\bt="str"/);
    assert.equal(disk[ref].formula, input.sheets[0].cells[ref].formula);
  }
  assert.equal(disk.E13.bold, true);
  assert.equal(disk.E19.cachedValue, oracle.cells.E10.expected);
  const request = path.join(f.root, 'independent-package-request.json');
  await fs.writeFile(
    request,
    JSON.stringify({
      source: f.source,
      saved: output.path,
      targets: { Cases: ['B1', 'E13'] },
      styleTargets: { Cases: ['E13'] },
      appendStyles: true,
      allowDefaultBlankNormalization: true,
      expected: {
        'Cases!B1': { value: 16 },
        'Cases!E13': { formula: 'DEC2HEX(255)', rawValue: 'FF', value: 'FF' },
        'Cases!E19': { formula: 'DEC2HEX($B$1)', rawValue: '10', value: '10' },
      },
      expectedFormulaCaches: { 'Cases!E13': { value: 'FF' }, 'Cases!E19': { value: '10' } },
      expectedBold: { sheet: 'Cases', cell: 'E13', value: true },
    })
  );
  const checked = spawnSync(
    'python3',
    [fileURLToPath(new URL('../scripts/support/standalone-package-oracle.py', import.meta.url)), request],
    { encoding: 'utf8' }
  );
  assert.equal(checked.status, 0, checked.stderr || checked.stdout);
  assert.equal(JSON.parse(checked.stdout).passed, true, checked.stdout);
  const reopened = await owner(t, output.path);
  assert.equal((await reopened.book.read('0', 'Cases', 'E19')).cells.E19.calculation.value, oracle.cells.E10.expected);
  const replay = await owner(t, f.source),
    recovered = await replay.book.apply('0', changed.edits);
  assert.equal(
    (await replay.book.read(recovered.view.storage.state, 'Cases', 'E19')).cells.E19.calculation.value,
    oracle.cells.E10.expected
  );
  assert.deepEqual(await fs.readFile(f.source), f.original);
});

test('full mixed actual workbook preserves all scalar fallback results and fourteen protected array contexts', async (t) => {
  const { oracle, bytes } = await oracleFixture(),
    root = await directory(t),
    source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, bytes);
  const f = await owner(t, source),
    cells = (await f.book.read('0', 'Cases', 'E1:I164')).cells,
    arrays = new Set(oracle.sourceArrayMasters);
  for (const ref of Object.keys(oracle.authoredFormulas)) {
    const result = cells[ref].calculation,
      expected = oracle.cells[ref];
    if (arrays.has(ref)) {
      assert.equal(result.status, 'unresolved', ref);
      continue;
    }
    assert.equal(result.status, expected.dataType === 'error' ? 'error' : 'calculated', ref);
    assert.equal(expected.dataType === 'error' ? result.error : result.value, expected.expected, ref);
  }
  assert.deepEqual(await fs.readFile(source), bytes);
});

test('mixed native text and compatibility lookup recalculate transitively and preserve older branches', async (t) => {
  const { oracle, view } = await oracleFixture(),
    input = projection(view, new Set(['E19', 'E85']));
  // The lookup reads the native text cell, rather than recomputing DEC2HEX in
  // compatibility. Actual I13/I1/I16 caches qualify the three table lookups.
  input.sheets[0].cells.E85 = { formula: 'VLOOKUP(E19,$N$1:$O$3,2,FALSE)' };
  const root = await directory(t),
    nativeOnly = projection(input, new Set(['E19']));
  const nativeResult = await nativeFormula(native, root, 'evaluate', { view: nativeOnly });
  assert.equal(nativeResult.supported, true);
  assert.equal(nativeResult.results.find((item) => item.ref === 'E19').value, oracle.cells.E19.expected);
  const mixedResult = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(mixedResult.supported, true);
  assert.equal(mixedResult.results.find((i) => i.ref === 'E85').value, oracle.cells.I13.expected);
  const edges = await nativeFormula(native, root, 'analyze', { view: input });
  assert.equal(edges.supported, true);
  assert(edges.items.find((item) => item.ref === 'E85').ranges.some((range) => range.ref === 'E19'));
  const f = await scalarPackage(t, input),
    probe = spawnSync(
      process.execPath,
      [fileURLToPath(new URL('./support/dec2hex-native-probe.mjs', import.meta.url)), f.source],
      { encoding: 'utf8', env: process.env, timeout: 120000 }
    );
  assert.equal(probe.status, 0, probe.stderr);
  const proof = JSON.parse(probe.stdout);
  assert.deepEqual(proof.before, []);
  assert.deepEqual(proof.after, []);
  assert(proof.diagnostic.nativeSqlite);
  assert.equal(proof.diagnostic.inputCells, 0);
  assert.equal(proof.results.E19.value, oracle.cells.E19.expected);
  assert.equal(proof.results.E85.value, oracle.cells.I13.expected);
  const owned = await owner(t, f.source);
  assert(owned.diagnostic().nativeSqlite);
  assert.equal((await owned.book.read('0', 'Cases', 'E85')).cells.E85.calculation.value, oracle.cells.E85.expected);
  const changed = await owned.book.apply('0', [{ sheet: 'Cases', cell: 'B1', value: 0 }]),
    branch = await owned.book.apply('0', [{ sheet: 'Cases', cell: 'B1', value: 2147483647 }]);
  assert.equal(
    (await owned.book.read(changed.view.storage.state, 'Cases', 'E19')).cells.E19.calculation.value,
    oracle.cells.E1.expected
  );
  assert.equal(
    (await owned.book.read(changed.view.storage.state, 'Cases', 'E85')).cells.E85.calculation.value,
    oracle.cells.I1.expected
  );
  assert.equal(
    (await owned.book.read(branch.view.storage.state, 'Cases', 'E19')).cells.E19.calculation.value,
    oracle.cells.E16.expected
  );
  assert.equal(
    (await owned.book.read(branch.view.storage.state, 'Cases', 'E85')).cells.E85.calculation.value,
    oracle.cells.I16.expected
  );
  assert.equal((await owned.book.read('0', 'Cases', 'E19')).cells.E19.calculation.value, oracle.cells.E19.expected);
  assert.equal((await owned.book.read('0', 'Cases', 'E85')).cells.E85.calculation.value, oracle.cells.E85.expected);
});

test('optional places stays on compatibility and actually loads the pinned providers', async (t) => {
  const { oracle, view } = await oracleFixture(),
    input = projection(view, new Set(['E43'])),
    root = await directory(t);
  const missed = await nativeFormula(native, root, 'evaluate', { view: input });
  assert.equal(missed.supported, false);
  assert.equal(missed.results, undefined);
  const f = await scalarPackage(t, input),
    probe = spawnSync(
      process.execPath,
      [fileURLToPath(new URL('./support/dec2hex-native-probe.mjs', import.meta.url)), f.source],
      { encoding: 'utf8', env: process.env, timeout: 120000 }
    );
  assert.equal(probe.status, 0, probe.stderr);
  const proof = JSON.parse(probe.stdout);
  assert.deepEqual(proof.before, []);
  assert(proof.after.length > 0);
  assert.equal(proof.diagnostic.nativePartialCells, 0);
  assert.equal(proof.diagnostic.nativeEvaluator, false);
  assert.equal(proof.results.E43.value, oracle.cells.E43.expected);
  assert.equal(typeof proof.results.E43.value, 'string');
});

test('qualified shared DEC2HEX groups retain typed caches and immutable normalization branches', async (t) => {
  const root = await directory(t),
    source = path.join(root, 'shared.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${[255, 255, 255]
      .map((value, index) => {
        const row = index + 1;
        return `<row r="${row}"><c r="A${row}"><v>${value}</v></c><c r="B${row}"><f t="shared" si="0"${row === 1 ? ' ref="B1:B3"' : ''}>${row === 1 ? 'DEC2HEX(A1)' : ''}</f><v>777</v></c></row>`;
      })
      .join('')}</sheetData></worksheet>`
  );
  await fs.writeFile(source, zipSync(parts));
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  await book.open(source);
  t.after(() => book.close());
  const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A2', value: 16 }]),
    branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 'replacement' }]);
  assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'B2')).cells.B2.calculation.value, '10');
  assert.equal((await book.read('0', 'Sheet1', 'B2')).cells.B2.calculation.value, 'FF');
  for (const [state, sharedCount] of [
    [changed.view.storage.state, 3],
    [branch.view.storage.state, 0],
  ]) {
    const output = await book.export(state),
      bytes = await fs.readFile(output.path),
      saved = readWorkbook(bytes).sheets[0].cells,
      xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
    assert.equal((xml.match(/t="shared"/g) ?? []).length, sharedCount);
    assert.equal(saved.B1.cachedValue, 'FF');
    assert.equal(saved.B3.cachedValue, 'FF');
    assert.equal(
      state === changed.view.storage.state ? saved.B2.cachedValue : saved.B2.value,
      state === changed.view.storage.state ? '10' : 'replacement'
    );
  }
});
