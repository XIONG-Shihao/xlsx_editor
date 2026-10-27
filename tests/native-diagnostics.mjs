import assert from 'node:assert/strict';
import { test } from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbook } from '../core/index.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';

test('native diagnostics are opt-in, content-free, and do not alter storage or accepted results', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-diagnostics-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary());
  const traces = [],
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>17</v></c><c r="B1"><f>A1*2</f></c><c r="C1" t="inlineStr"><is><t>private-content-marker</t></is></c></row></sheetData></worksheet>'
  );
  const source = path.join(root, 'secret-source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  let enabled = true;
  const native = async (input) => {
    const result = await client.call({ ...input, diagnostics: enabled });
    if (enabled) {
      assert.equal(result.diagnostics.schema, 1);
      assert(Number.isFinite(result.diagnostics.elapsedMs));
      traces.push({ operation: input.operation, ...result.diagnostics });
      delete result.diagnostics;
    } else assert.equal(result.diagnostics, undefined);
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(source);
    const first = await book.read('0', 'Sheet1', 'A1:C1');
    assert.equal(first.cells.B1.calculation.value, 34);
    const index = traces.find((t) => t.operation === 'index');
    assert.equal(index.counters['import.cells'], 3);
    assert.equal(index.counters['import.rows'], 1);
    assert.equal(index.counters['import.formulas'], 1);
    assert.equal(index.counters['import.insert_rows'], 3);
    assert.equal(index.counters['import.insert_statements'], 1);
    assert(index.spans.some((s) => s.name === 'import.worksheet'));
    assert(index.spans.some((s) => s.name === 'import.cell_pack_sample' && s.calls === 1));
    const calculation = traces.find((t) => t.operation === 'calculate');
    assert.equal(calculation.counters['calculation.evaluated'], 1);
    assert(calculation.spans.some((s) => s.name === 'calculation.evaluate'));
    for (const marker of ['private-content-marker', 'secret-source.xlsx', 'A1*2', root])
      assert(!JSON.stringify(traces).includes(marker));
    enabled = false;
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 21 }]);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 42);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 34);
    const saved = await book.export(next.view.storage.state);
    assert(await fs.stat(saved.path));
    await assert.rejects(
      client.call({ action: 'workbook', operation: 'not-an-operation', root, diagnostics: true }),
      { code: 'internal-unknown-reference' }
    );
    assert.equal(
      (
        await client.call({
          action: 'workbook',
          operation: 'artifact-write',
          root,
          name: 'display-formats.json',
          data: '[]',
        })
      ).diagnostics,
      undefined
    );
  } finally {
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('shared edits scope dependency snapshots while calculation retains full, immutable inputs', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-scoped-snapshot-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary()),
    traces = [];
  const parts = unzipSync(createWorkbook()),
    size = 4096;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: size }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${size}"` : ''}>${i === 0 ? '1+1' : ''}</f><v>2</v></c>${i === size - 1 ? `<c r="B${size}"><f>SUM(A1:A${size})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    const result = await client.call({ ...input, diagnostics: true });
    traces.push({ operation: input.operation, method: input.method, ...result.diagnostics });
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(source);
    traces.length = 0;
    const next = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A2000', value: 5 },
      { sheet: 'Sheet1', cell: 'A2001', value: 7 },
    ]);
    const normalization = traces.find((t) => t.operation === 'normalize-shared');
    assert.equal(normalization.counters['shared.validation_proof_hits'], 1);
    assert.equal(normalization.counters['shared.validation_reused_members'], size);
    assert.equal(
      normalization.counters['shared.validated_members'] ?? 0,
      0,
      'The exact calculated version already proved every follower'
    );
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B4096')).cells.B4096.calculation.value, 8200);
    assert.equal((await book.read('0', 'Sheet1', 'B4096')).cells.B4096.calculation.value, 8192);
    const update = traces.find((t) => t.operation === 'dependencies' && t.method === 'update-formulas');
    assert.equal(
      update.counters['calculation.snapshot_rows'],
      2,
      'Formula-edge replacement must not materialize thousands of unrelated overlays'
    );
    const calculation = traces.find((t) => t.operation === 'calculate');
    assert.equal(calculation.counters['calculation.evaluated'], 1);
    assert.equal(calculation.counters['shared.special_scan_skipped'], 1);
    assert.equal(calculation.counters['calculation.shared_overlay_reused_rows'] ?? 0, 0);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', formula: '3+3' }]);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'B4096')).cells.B4096.calculation.value, 8196);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B4096')).cells.B4096.calculation.value, 8200);
    const saved = await book.export(next.view.storage.state),
      reopened = new IndexedWorkbook(
        native,
        () => assert.fail('legacy writer'),
        (input) => calculateIndexedWorkbook(input, native)
      );
    try {
      await reopened.open(saved.path);
      assert.equal((await reopened.read('0', 'Sheet1', 'B4096')).cells.B4096.calculation.value, 8200);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('shared validation proofs are exact-version, rollback-safe and pruned with abandoned revisions', async () => {
  const { DatabaseSync } = await import('node:sqlite');
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shared-proof-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary()),
    size = 2048,
    parts = unzipSync(createWorkbook()),
    traces = [];
  let root, originalVersion;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: size }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${size}"` : ''}>${i === 0 ? '$C1+1' : ''}</f><v>999</v></c><c r="C${i + 1}"><v>${i + 1}</v></c>${i === size - 1 ? `<c r="D${size}"><f>SUM(A1:A${size})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  const source = path.join(temporary, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    if (input.operation === 'index') root = input.root;
    const result = await client.call({ ...input, diagnostics: true });
    if (input.operation === 'calculate' && originalVersion === undefined) originalVersion = result.version;
    traces.push({ operation: input.operation, ...result.diagnostics });
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  const inventory = () => {
    const db = new DatabaseSync(path.join(root, 'cells.sqlite'), { readOnly: true });
    try {
      return Object.fromEntries(
        ['versions', 'shared_normalizations', 'shared_validations'].map((table) => [
          table,
          db.prepare(`SELECT count(*) AS count FROM ${table}`).get().count,
        ])
      );
    } finally {
      db.close();
    }
  };
  try {
    await book.open(source);
    assert.equal(inventory().shared_validations, 1);
    const original = inventory(),
      request = {
        action: 'workbook',
        operation: 'normalize-shared',
        root,
        version: originalVersion,
        ranges: [{ sheet: 'xl/worksheets/sheet1.xml', range: 'A1000' }],
      };
    await fs.writeFile(path.join(root, 'cancelled'), 'test');
    await assert.rejects(native(request), { code: /^(internal-invalid-field|operation-cancelled)$/ });
    await fs.rm(path.join(root, 'cancelled'));
    assert.deepEqual(inventory(), original);
    await assert.rejects(
      native({ ...request, ranges: [...request.ranges, { sheet: 'xl/worksheets/sheet1.xml', range: 'A0' }] }),
      { code: /^(hyperlink-unsupported|package-unsupported)$/ }
    );
    assert.deepEqual(inventory(), original, 'A later invalid command rolls back an earlier proof-backed normalization');
    traces.length = 0;
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1000', value: 5 }]);
    assert.equal(
      traces.find((t) => t.operation === 'normalize-shared').counters['shared.validation_reused_members'],
      size
    );
    const expected = (size * (size + 1)) / 2 + size;
    assert.equal((await book.read('0', 'Sheet1', `D${size}`)).cells[`D${size}`].calculation.value, expected);
    assert.equal(
      (await book.read(changed.view.storage.state, 'Sheet1', `D${size}`)).cells[`D${size}`].calculation.value,
      expected - 1001 + 5
    );
    const styled = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A2', bold: true }]);
    traces.length = 0;
    const styledEdit = await book.apply(styled.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1000', value: 7 }]);
    const fallback = traces.find((t) => t.operation === 'normalize-shared');
    assert.equal(
      fallback.counters['shared.validation_proof_hits'] ?? 0,
      0,
      'An ancestor proof must not authorize a different native version'
    );
    assert.equal(fallback.counters['shared.validated_members'], size);
    assert.equal((await book.read(styledEdit.view.storage.state, 'Sheet1', 'A2')).cells.A2.bold, true);
    assert.equal((await book.read(styledEdit.view.storage.state, 'Sheet1', 'A2')).cells.A2.formula, '$C2+1');
    await book.apply('0', [{ type: 'recalculate' }]);
    assert.equal(inventory().shared_validations, 2);
    await book.compact(['0', changed.view.storage.state]);
    assert.equal(inventory().shared_validations, 1, 'Abandoned calculation proofs must not retain history');
    const exported = await book.export(changed.view.storage.state),
      reopened = new IndexedWorkbook(
        native,
        () => assert.fail('writer fallback'),
        (input) => calculateIndexedWorkbook(input, native)
      );
    try {
      await reopened.open(exported.path);
      assert.equal(
        (await reopened.read('0', 'Sheet1', `D${size}`)).cells[`D${size}`].calculation.value,
        expected - 1001 + 5
      );
      assert.equal((await reopened.read('0', 'Sheet1', `A${size}`)).cells[`A${size}`].formula, `$C${size}+1`);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
    client.close();
    await fs.rm(temporary, { recursive: true, force: true });
  }
});

test('compact shared definitions translate mixed references and reuse only constant results within one calculation', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shared-definition-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary()),
    traces = [];
  const size = 1024,
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: size }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${size}"` : ''}>${i === 0 ? 'SUM(1,2)' : ''}</f><v>999</v></c><c r="C${i + 1}"><f t="shared" si="1"${i === 0 ? ` ref="C1:C${size}"` : ''}>${i === 0 ? '$D1+1' : ''}</f><v>999</v></c><c r="D${i + 1}"><v>${i + 1}</v></c>${i === size - 1 ? `<c r="E${size}"><f>SUM(A1:A${size})+SUM(C1:C${size})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    const result = await client.call({ ...input, diagnostics: true });
    traces.push({ operation: input.operation, ...result.diagnostics });
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(source);
    const calculation = traces.find((t) => t.operation === 'calculate');
    assert.equal(calculation.counters['calculation.shared_groups'], 2);
    assert.equal(calculation.counters['calculation.shared_members_validated'], size * 2);
    assert.equal(calculation.counters['calculation.constant_hits'], size - 1);
    assert.equal(calculation.counters['calculation.constant_cache_entries'], 1);
    assert(calculation.counters['calculation.constant_cache_bytes'] <= 64 * 1024);
    assert(calculation.counters['calculation.syntax_cache_bytes'] <= 1024 * 1024);
    assert(calculation.counters['calculation.syntax_cache_entries'] <= 128);
    assert(calculation.counters['calculation.syntax_cache_hits'] > size);
    const initial = (await book.read('0', 'Sheet1', 'A1024:E1024')).cells;
    assert.equal(initial.A1024.calculation.value, 3);
    assert.equal(initial.C1024.calculation.value, 1025);
    const total = 3 * size + (size * (size + 3)) / 2;
    assert.equal(initial.E1024.calculation.value, total);
    traces.length = 0;
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D1024', value: 5000 }]);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'C1024')).cells.C1024.calculation.value, 5001);
    assert.equal(
      (await book.read(next.view.storage.state, 'Sheet1', 'E1024')).cells.E1024.calculation.value,
      total + 3976
    );
    assert.equal((await book.read('0', 'Sheet1', 'C1024')).cells.C1024.calculation.value, 1025);
    assert.equal(
      traces.find((t) => t.operation === 'calculate').counters['calculation.constant_hits'],
      0,
      'A new calculation must not reuse constant values from an earlier command'
    );
    traces.length = 0;
    const mixed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
    assert.equal((await book.read(mixed.view.storage.state, 'Sheet1', 'C1024')).cells.C1024.calculation.value, 1025);
    assert.equal(
      (await book.read(mixed.view.storage.state, 'Sheet1', 'E1024')).cells.E1024.calculation.value,
      total + 6
    );
    assert.equal(
      traces.find((t) => t.operation === 'calculate').counters['calculation.shared_groups'],
      1,
      'Normalizing one group must not bypass validation of another shared group'
    );
    const exported = await book.export(next.view.storage.state),
      reopened = new IndexedWorkbook(
        native,
        () => assert.fail('writer fallback'),
        (i) => calculateIndexedWorkbook(i, native)
      );
    try {
      await reopened.open(exported.path);
      assert.equal((await reopened.read('0', 'Sheet1', 'E1024')).cells.E1024.calculation.value, total + 3976);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('constant and syntax reuse budgets stay bounded across many distinct formulas and reset on edits', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-constant-budget-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary()),
    traces = [];
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: 300 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f>${i + 1}+1</f></c></row>`).join('')}</sheetData></worksheet>`
  );
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    const result = await client.call({ ...input, diagnostics: true });
    if (input.operation === 'calculate') traces.push(result.diagnostics);
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('legacy writer'),
    (i) => calculateIndexedWorkbook(i, native)
  );
  try {
    await book.open(source);
    assert.equal(traces[0].counters['calculation.constant_cache_entries'], 256);
    assert(traces[0].counters['calculation.constant_cache_bytes'] <= 64 * 1024);
    assert.equal(traces[0].counters['calculation.syntax_cache_entries'], 128);
    assert(traces[0].counters['calculation.syntax_cache_bytes'] <= 1024 * 1024);
    assert.equal((await book.read('0', 'Sheet1', 'A300')).cells.A300.calculation.value, 301);
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', formula: '2+2' }]);
    assert.equal(traces.at(-1).counters['calculation.constant_cache_entries'], 1);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'A1')).cells.A1.calculation.value, 4);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.calculation.value, 2);
  } finally {
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('shared group and syntax reuse preserve two-dimensional offsets and sheet-local name context', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shared-context-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary()),
    traces = [];
  const parts = unzipSync(editWorkbook(createWorkbook(), [{ operation: 'addSheet', sheet: 'Other' }]));
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="Rate" localSheetId="0">2</definedName><definedName name="Rate" localSheetId="1">7</definedName></definedNames></workbook>'
    )
  );
  for (const [index, scale] of [
    [1, 1],
    [2, 10],
  ])
    parts[`xl/worksheets/sheet${index}.xml`] = strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>${10 * scale}</v></c><c r="B1"><v>${20 * scale}</v></c><c r="C1"><v>${30 * scale}</v></c><c r="F1"><f>Rate+1</f></c><c r="G1"><f>Rate+1</f></c></row><row r="2"><c r="A2"><v>${2 * scale}</v></c><c r="C2"><f t="shared" si="0" ref="C2:D3">$A2+B$1+$A$1+Rate</f><v>999</v></c><c r="D2"><f t="shared" si="0"/><v>999</v></c></row><row r="3"><c r="A3"><v>${3 * scale}</v></c><c r="C3"><f t="shared" si="0"/><v>999</v></c><c r="D3"><f t="shared" si="0"/><v>999</v></c></row></sheetData></worksheet>`
    );
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    const result = await client.call({ ...input, diagnostics: true });
    if (input.operation === 'calculate') traces.push({ ...result });
    delete result.diagnostics;
    return result;
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    (i) => calculateIndexedWorkbook(i, native)
  );
  const values = async (state, sheet) => {
    const cells = (await book.read(state, sheet, 'C1:G3')).cells;
    return ['C2', 'D2', 'C3', 'D3', 'F1', 'G1'].map((ref) => cells[ref].calculation.value);
  };
  try {
    await book.open(source);
    assert.equal(traces.length, 1);
    assert.equal(traces[0].nativeSqlite, true);
    assert.equal(traces[0].supported, true);
    assert.deepEqual(await values('0', 'Sheet1'), [34, 44, 35, 45, 3, 3]);
    assert.deepEqual(await values('0', 'Other'), [327, 427, 337, 437, 8, 8]);
    const counters = traces[0].diagnostics.counters;
    assert.equal(counters['calculation.shared_groups'], 2);
    assert.equal(counters['calculation.shared_members_validated'], 8);
    assert.equal(counters['calculation.constant_cache_entries'], 2);
    assert.equal(counters['calculation.constant_hits'], 2);
    assert(counters['calculation.syntax_cache_hits'] > 0);
    const next = await book.apply('0', [{ sheet: 'Other', cell: 'A2', value: 50 }]);
    assert.deepEqual(await values(next.view.storage.state, 'Other'), [357, 457, 337, 437, 8, 8]);
    assert.deepEqual(await values(next.view.storage.state, 'Sheet1'), [34, 44, 35, 45, 3, 3]);
    assert.deepEqual(await values('0', 'Other'), [327, 427, 337, 437, 8, 8]);
    const saved = await book.export(next.view.storage.state),
      reopened = new IndexedWorkbook(
        native,
        () => assert.fail('writer fallback'),
        (i) => calculateIndexedWorkbook(i, native)
      );
    try {
      await reopened.open(saved.path);
      assert.equal((await reopened.read('0', 'Other', 'D2')).cells.D2.calculation.value, 457);
      assert.equal((await reopened.read('0', 'Sheet1', 'D2')).cells.D2.calculation.value, 44);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('partial shared edits persist compact group intent through branches, compaction, bulk and export', async () => {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-sparse-shared-')),
    client = createWorkbookNativeClient(nativeWorkbookBinary());
  let root, lastReadVersion;
  const size = 4096,
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: size }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${size}"` : ''}>${i === 0 ? '1+1' : ''}</f><v>999</v></c>${i === size - 1 ? `<c r="B${size}"><f>SUM(A1:A${size})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  const source = path.join(temporary, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const native = async (input) => {
    if (input.operation === 'index') root = input.root;
    if (input.operation === 'range') lastReadVersion = input.version;
    return client.call(input);
  };
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    (i) => calculateIndexedWorkbook(i, native)
  );
  const stats = () => native({ action: 'workbook', operation: 'stats', root });
  const sum = async (state) => (await book.read(state, 'Sheet1', `B${size}`)).cells[`B${size}`].calculation.value;
  try {
    await book.open(source);
    const before = await stats();
    assert.equal(before.normalizedSharedGroups, 0);
    const first = (await book.apply('0', [{ sheet: 'Sheet1', cell: 'A2000', value: 5 }])).view.storage.state;
    const after = await stats();
    assert.equal(after.normalizedSharedGroups, 1);
    assert.equal(
      after.changedCellRecords - before.changedCellRecords,
      2,
      'One cell edit plus affected SUM, not 4096 expanded formula records'
    );
    assert.equal(await sum(first), 8195);
    assert.equal(await sum('0'), 8192);
    const originalVersion = lastReadVersion;
    await sum(first);
    const firstVersion = lastReadVersion;
    const diff = async (before, version) => {
      let offset = 0,
        count = 0,
        changes = [];
      do {
        const page = await native({ action: 'workbook', operation: 'diff', root, before, version, offset, styles: [] });
        count += page.count;
        changes.push(...page.changes);
        offset = page.nextOffset;
      } while (offset !== null);
      return { count, changes };
    };
    for (const [before, after] of [
      [originalVersion, firstVersion],
      [firstVersion, originalVersion],
    ]) {
      const evidence = await diff(before, after);
      assert.equal(evidence.count, 2, 'Shared representation alone must not invent neighboring cell changes');
      assert.deepEqual(evidence.changes.map((c) => c.cell).sort(), ['A2000', 'B4096']);
      const edit = evidence.changes.find((c) => c.cell === 'A2000');
      assert.equal(edit.before.value, before === originalVersion ? 2 : 5);
      assert.equal(edit.after.value, after === firstVersion ? 5 : 2);
    }
    assert.equal(
      (await book.query(first, 'Sheet1', '1+1', 'formula', undefined, { range: 'A4000:A4000' })).cells.A4000.formula,
      '1+1'
    );
    const copied = (
      await book.apply(first, [
        { type: 'pasteFormulas', sheet: 'Sheet1', source: 'A4000', range: 'C20' },
        { type: 'fillRange', sheet: 'Sheet1', source: 'A4000', range: 'D20:D22' },
      ])
    ).view.storage.state;
    const copyCells = (await book.read(copied, 'Sheet1', 'C20:D22')).cells;
    assert.equal(copyCells.C20.formula, '1+1');
    assert.equal(copyCells.D22.calculation.value, 2);
    assert.equal(await sum(first), 8195);
    // Normalized source coordinates must not hide later array masters. Exercise
    // the low-level calculation patch route that legitimately stores metadata.
    for (const [cell, range, target] of [
      ['A1', 'A1:B2', 'B2'],
      ['C10', 'C10:D11', 'D11'],
    ]) {
      const special = await native({
        action: 'workbook',
        operation: 'patch',
        root,
        version: firstVersion,
        changes: [
          {
            sheet: 'xl/worksheets/sheet1.xml',
            cell,
            value: { value: null, formula: 'SUM(E1:E2)', formulaType: 'array', formulaReference: range },
          },
        ],
      });
      await assert.rejects(
        native({
          action: 'workbook',
          operation: 'patch',
          root,
          version: special.version,
          changes: [
            {
              sheet: 'xl/worksheets/sheet1.xml',
              cell: target,
              value: { value: 1, dataType: 'number' },
              edits: [{ value: 1 }],
            },
          ],
        }),
        /complete group|partial/
      );
    }
    assert.equal((await book.read(first, 'Sheet1', 'A1')).cells.A1.formula, '1+1');
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.formulaType, 'shared');
    const master = (
      await book.apply(first, [
        { sheet: 'Sheet1', cell: 'A1', formula: 'C1+3' },
        { sheet: 'Sheet1', cell: 'C1', value: 10 },
      ])
    ).view.storage.state;
    assert.equal(await sum(master), 8206);
    const masterVersion = lastReadVersion;
    for (const [before, after] of [
      [firstVersion, masterVersion],
      [masterVersion, firstVersion],
    ]) {
      const evidence = await diff(before, after);
      assert.equal(evidence.count, 3);
      assert.deepEqual(evidence.changes.map((c) => c.cell).sort(), ['A1', 'B4096', 'C1']);
    }
    const branch = (await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 6 }])).view.storage.state;
    assert.equal(await sum(branch), 8196);
    const branchVersion = lastReadVersion;
    assert.equal(await sum(first), 8195);
    for (const [before, after] of [
      [firstVersion, branchVersion],
      [branchVersion, firstVersion],
    ]) {
      const evidence = await diff(before, after);
      assert.equal(evidence.count, 3);
      assert.deepEqual(evidence.changes.map((c) => c.cell).sort(), ['A1', 'A2000', 'B4096']);
    }
    await book.compact([first, master]);
    assert.equal(
      (await stats()).normalizedSharedGroups,
      1,
      'Abandoned branch group intent must be pruned and private intent flattened'
    );
    assert.equal(await sum(master), 8206);
    assert.equal(await sum(first), 8195);
    assert.equal(await sum('0'), 8192);
    const cleared = (await book.apply(master, [{ type: 'clearRange', sheet: 'Sheet1', range: 'A2:A4' }])).view.storage
      .state;
    assert.equal(await sum(cleared), 8200);
    const styled = (
      await book.apply(cleared, [{ type: 'formatRange', sheet: 'Sheet1', range: 'A4000', format: { bold: true } }])
    ).view.storage.state;
    const saved = await book.export(styled),
      reopened = new IndexedWorkbook(
        native,
        () => assert.fail('writer fallback'),
        (i) => calculateIndexedWorkbook(i, native)
      );
    try {
      await reopened.open(saved.path);
      const cells = (await reopened.read('0', 'Sheet1', 'A4000')).cells;
      assert.equal(cells.A4000.bold, true);
      assert.equal(cells.A4000.formula, '1+1');
      assert.equal((await reopened.read('0', 'Sheet1', `B${size}`)).cells[`B${size}`].calculation.value, 8200);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
    client.close();
    await fs.rm(temporary, { recursive: true, force: true });
  }
});
