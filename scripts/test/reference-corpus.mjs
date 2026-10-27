/** Replay recorded external oracles through both current calculation routes.
 * Expected values come from the pinned external oracle, never our evaluator. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { readCalculatedWorkbook } from '../../core/workbook-service.mjs';
import { createWorkbookNativeClient } from '../../core/native-client.mjs';
import { nativeWorkbookBinary } from '../../core/native-runtime.mjs';

const component = fileURLToPath(new URL('../../', import.meta.url));
const fixtureDirectory = path.join(component, 'fixtures');
const args = process.argv.slice(2);
assert(
  args.length === 0 || (args.length === 2 && args[0] === '--output'),
  'Usage: node scripts/test/reference-corpus.mjs [--output DIRECTORY]'
);
const output = path.resolve(
  args[1] ?? path.join(component, '.build/tests/reference-corpus-' + Date.now())
);
await fs.mkdir(output, { recursive: true });
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const manifestBytes = await fs.readFile(path.join(fixtureDirectory, 'reference-corpus.json'));
const manifest = JSON.parse(manifestBytes);
assert.equal(manifest.schema, 'workbook.xlsx.reference-corpus.v1');
const binary = nativeWorkbookBinary();
const report = {
  schema: 'workbook.xlsx.reference-corpus-run.v1',
  startedAt: new Date().toISOString(),
  node: process.version,
  platform: process.platform,
  arch: process.arch,
  binary,
  binarySha256: digest(await fs.readFile(binary)),
  manifestSha256: digest(manifestBytes),
  provenance: manifest.provenance,
  unqualified: manifest.unqualified,
  fixtures: [],
  failures: [],
  passed: false,
};
const client = createWorkbookNativeClient(binary);
function member(name) {
  assert.match(name, /^[a-zA-Z0-9][a-zA-Z0-9_.-]*$/);
  return path.join(fixtureDirectory, name);
}
function expected(entry, item, lane) {
  if (lane === 'compatibility' && entry.compatibilityUnresolved?.includes(item.ref)) return { status: 'unresolved' };
  return entry.errorRefs?.includes(item.ref)
    ? { status: 'error', error: item.expected }
    : { status: 'calculated', value: item.expected };
}
function check(cell, wanted, revision, label) {
  assert(cell?.calculation, label + ': missing authoritative calculation');
  assert.equal(cell.calculation.revision, revision, label + ': calculation revision');
  assert.equal(cell.calculation.status, wanted.status, label + ': calculation status');
  if (wanted.status === 'calculated')
    assert(
      Object.is(cell.calculation.value, wanted.value),
      label + ': exact typed value ' + JSON.stringify({ actual: cell.calculation.value, expected: wanted.value })
    );
  if (wanted.status === 'error') assert.equal(cell.calculation.error, wanted.error, label + ': error');
}
function bounds(cases) {
  let row = 1,
    column = 1;
  for (const item of cases) {
    const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(item.ref);
    assert(match, 'Invalid oracle reference');
    row = Math.max(row, Number(match[2]));
    column = Math.max(
      column,
      [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
    );
  }
  assert(row * Math.max(column, 26) <= 10000, 'Reference corpus must remain a minimized bounded read');
  return 'A1:Z' + row;
}
try {
  for (const entry of manifest.fixtures) {
    const source = member(entry.source),
      oraclePath = member(entry.oracle),
      bytes = await fs.readFile(source),
      oracleBytes = await fs.readFile(oraclePath);
    assert.equal(digest(bytes), entry.sourceSha256, entry.id + ': source identity');
    assert.equal(digest(oracleBytes), entry.oracleSha256, entry.id + ': external oracle identity');
    const oracle = JSON.parse(oracleBytes);
    assert.equal(oracle.excel, '16.113');
    assert(oracle.cases.length > 0);
    const record = {
      id: entry.id,
      sourceSha256: entry.sourceSha256,
      oracleSha256: entry.oracleSha256,
      referenceApplication: 'Excel ' + oracle.excel,
      boundaries: entry.boundaries,
      cases: oracle.cases.length,
      compatibilityRefusals: entry.compatibilityUnresolved ?? [],
      compatibilityReason: entry.compatibilityReason,
      calculations: [],
      failures: [],
      passed: false,
    };
    report.fixtures.push(record);
    const verify = (callback) => {
      try {
        callback();
      } catch (error) {
        record.failures.push(String(error.message ?? error));
      }
    };
    const book = new IndexedWorkbook(
      client.call,
      async () => assert.fail('No whole-workbook mutation fallback'),
      async (input) => {
        const result = await calculateIndexedWorkbook(input, client.call);
        record.calculations.push({
          mode: result.mode,
          nativeSqlite: result.nativeSqlite,
          nativePartialCells: result.nativePartialCells,
          inputCells: result.inputCells,
          projectedBytes: result.projectedBytes,
          evaluatedFormulaCells: result.evaluatedFormulaCells,
        });
        return result;
      }
    );
    try {
      const initial = await book.open(source),
        range = bounds(oracle.cases);
      const before = (await book.read(initial.storage.state, entry.sheet, range)).cells;
      const compatible = await readCalculatedWorkbook(bytes),
        sheet = compatible.sheets.find((sheet) => sheet.name === entry.sheet);
      assert(sheet, entry.id + ': compatibility worksheet missing');
      for (const item of oracle.cases) {
        verify(() =>
          assert.equal(
            before[item.ref]?.formula,
            item.formula,
            entry.id + ' ' + item.ref + ': authored formula identity'
          )
        );
        verify(() =>
          check(before[item.ref], expected(entry, item, 'indexed'), initial.revision, entry.id + ' indexed ' + item.ref)
        );
        verify(() =>
          check(
            sheet.cells[item.ref],
            expected(entry, item, 'compatibility'),
            compatible.revision,
            entry.id + ' compatibility ' + item.ref
          )
        );
      }
      // A real accepted change forces preserving export of the calculated caches.
      const changed = await book.apply(initial.storage.state, [
        { sheet: entry.sheet, cell: 'Z1', value: 'reference-corpus-marker' },
      ]);
      const old = (await book.read(initial.storage.state, entry.sheet, range)).cells;
      for (const item of oracle.cases)
        verify(() => assert.deepEqual(old[item.ref], before[item.ref], entry.id + ': old revision ' + item.ref));
      const exported = await book.export(changed.view.storage.state),
        saved = path.join(output, entry.id + '.xlsx');
      await fs.copyFile(exported.path, saved);
      const request = {
        source,
        saved,
        sheet: entry.sheet,
        marker: { ref: 'Z1', value: 'reference-corpus-marker' },
        cases: oracle.cases.map((item) => ({
          ref: item.ref,
          formula: item.formula,
          ...expected(entry, item, 'indexed'),
        })),
      };
      const independent = spawnSync(
        'python3',
        [fileURLToPath(new URL('./reference-corpus-oracle.py', import.meta.url))],
        { input: JSON.stringify(request), encoding: 'utf8', timeout: 30000, maxBuffer: 1024 * 1024 }
      );
      assert(independent.stdout, independent.stderr || String(independent.error));
      record.saved = JSON.parse(independent.stdout);
      verify(() =>
        assert.equal(independent.status, 0, independent.stderr || independent.stdout || String(independent.error))
      );
      verify(() => assert(record.saved.passed));
      const reopened = await readCalculatedWorkbook(await fs.readFile(saved));
      for (const item of oracle.cases)
        verify(() =>
          check(
            reopened.sheets.find((s) => s.name === entry.sheet).cells[item.ref],
            expected(entry, item, 'compatibility'),
            reopened.revision,
            entry.id + ' cold reopen ' + item.ref
          )
        );
      assert.equal(digest(await fs.readFile(source)), entry.sourceSha256, entry.id + ': source preserved');
      record.passed = record.failures.length === 0;
    } finally {
      await book.close();
    }
  }
  report.passed = report.fixtures.every((fixture) => fixture.passed);
  if (!report.passed) {
    report.failures.push(...report.fixtures.flatMap((fixture) => fixture.failures));
    process.exitCode = 1;
  }
} catch (error) {
  report.failures.push(String(error.stack ?? error));
  process.exitCode = 1;
} finally {
  client.close();
  report.finishedAt = new Date().toISOString();
  await fs.writeFile(path.join(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(
    JSON.stringify({
      output,
      passed: report.passed,
      fixtures: report.fixtures.length,
      cases: report.fixtures.reduce((sum, f) => sum + f.cases, 0),
      failures: report.failures,
    })
  );
}
