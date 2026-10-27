import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import { formulaCacheExpectation, hashFile } from '../scripts/support/saved-package.mjs';
import { footprintTargets, parseFootprint, captureFootprint } from '../scripts/support/process-footprint.mjs';

test('native physical accounting never becomes an RSS sum or a false zero', async () => {
  const raw = {
    unit: 'byte',
    'bytes per unit': 1,
    'total footprint': 350,
    processes: [
      { pid: 10, footprint: 200 },
      { pid: 11, auxiliary: { phys_footprint: 200 } },
    ],
  };
  const measured = parseFootprint(raw, [10, 11]);
  assert.equal(measured.status, 'measured');
  assert.equal(measured.totalFootprintBytes, 350, 'Use OS aggregate; shared accounting is not the sum of children');
  const partial = parseFootprint({ ...raw, processes: raw.processes.slice(0, 1) }, [10, 11]);
  assert.equal(partial.status, 'partial');
  assert.equal(partial.totalFootprintBytes, null);
  assert.equal(parseFootprint({ ...raw, unit: 'kilobyte' }, [10, 11]).totalFootprintBytes, null);
  assert.equal(parseFootprint({ ...raw, processes: [...raw.processes, raw.processes[0]] }, [10, 11]).status, 'partial');
  assert.equal(parseFootprint(raw, []).status, 'unavailable');
  assert.deepEqual(
    footprintTargets(10, {
      processes: [
        { pid: 10, ppid: 1 },
        { pid: 11, ppid: 10 },
        { pid: 12, ppid: 11 },
        { pid: 20, ppid: 1 },
      ],
    }),
    [10, 11, 12]
  );
  assert.deepEqual(await captureFootprint({ rootPid: 10 }), { status: 'disabled', totalFootprintBytes: null });
});

test('independent saved-style checks catch a writer that loses Bold despite correct cell values', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-workload-oracle-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const { workbook } = await import('../fixtures/standalone.mjs'),
    { editWorkbook } = await import('../core/workbook-package.mjs');
  const source = path.join(directory, 'source.xlsx'),
    saved = path.join(directory, 'saved.xlsx'),
    request = path.join(directory, 'request.json');
  const bytes = workbook();
  await fs.writeFile(source, bytes);
  await fs.writeFile(saved, bytes);
  await fs.writeFile(
    request,
    JSON.stringify({
      source,
      saved,
      targets: { Sheet1: ['A1'] },
      styleTargets: { Sheet1: ['A1'] },
      appendStyles: true,
      expected: { 'Sheet1!A1': { value: 10 } },
      expectedBold: { sheet: 'Sheet1', cell: 'A1', value: true },
    })
  );
  const oracle = fileURLToPath(new URL('../scripts/support/standalone-package-oracle.py', import.meta.url)),
    exec = promisify(execFile);
  let failed;
  try {
    await exec('python3', [oracle, request]);
  } catch (error) {
    failed = JSON.parse(error.stdout);
  }
  assert.equal(failed.passed, false);
  assert(failed.failures.some((item) => item.kind === 'expected-saved-bold'));
  await fs.writeFile(saved, editWorkbook(bytes, [{ sheet: 'Sheet1', cell: 'A1', bold: true }]));
  const result = JSON.parse((await exec('python3', [oracle, request])).stdout);
  assert.equal(result.passed, true);
  assert.equal(result.boldCheck.actual, true);
  assert.equal(await hashFile(source), result.originalSha256);
});

test('saved formula cache checks retain text/boolean types and exact binary64 bits', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-workload-caches-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const { workbook } = await import('../fixtures/standalone.mjs'),
    { unzipSync, zipSync, strFromU8, strToU8 } = await import('fflate');
  const source = path.join(directory, 'source.xlsx'),
    saved = path.join(directory, 'saved.xlsx'),
    requestFile = path.join(directory, 'request.json');
  const original = workbook();
  await fs.writeFile(source, original);
  const cases = [
    { value: 'FF', formula: 'DEC2HEX(255)', type: 'str', payload: '<v>FF</v>', passed: true },
    { value: 'FF', formula: 'DEC2HEX(255)', type: 'inlineStr', payload: '<is><t>FF</t></is>', passed: false },
    { value: true, formula: 'TRUE', type: 'b', payload: '<v>1</v>', passed: true },
    { value: true, formula: 'TRUE', type: 'n', payload: '<v>1</v>', passed: false },
    { value: 0.1, formula: '1/10', type: 'n', payload: '<v>0.1</v>', passed: true },
    { value: 1, formula: '1', type: 'n', payload: '<v>1.0000000000000002</v>', passed: false },
    { value: -0, formula: '-0', type: 'n', payload: '<v>-0</v>', passed: true },
    { value: -0, formula: '-0', type: 'n', payload: '<v>0</v>', passed: false },
  ];
  const oracle = fileURLToPath(new URL('../scripts/support/standalone-package-oracle.py', import.meta.url)),
    exec = promisify(execFile);
  for (const item of cases) {
    const parts = unzipSync(original),
      xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      xml.replace('</row>', `<c r="D1" t="${item.type}"><f>${item.formula}</f>${item.payload}</c></row>`)
    );
    await fs.writeFile(saved, zipSync(parts));
    // Numeric calculated-only expectations intentionally exercise the inherited
    // approximate check. The new cache check must catch even one-bit changes.
    const expected =
      typeof item.value === 'number'
        ? { calculated: item.value, formula: item.formula }
        : { value: item.value, formula: item.formula };
    await fs.writeFile(
      requestFile,
      JSON.stringify({
        source,
        saved,
        targets: { Sheet1: ['D1'] },
        expected: { 'Sheet1!D1': expected },
        expectedFormulaCaches: { 'Sheet1!D1': formulaCacheExpectation(item.value) },
      })
    );
    let stdout;
    try {
      stdout = (await exec('python3', [oracle, requestFile])).stdout;
    } catch (error) {
      stdout = error.stdout;
    }
    const result = JSON.parse(stdout);
    assert.equal(result.passed, item.passed, `${item.type} ${item.payload}`);
    assert.equal(result.formulaCacheChecks[0].passed, item.passed);
    if (!item.passed)
      assert(
        result.checks.every((check) => check.passed),
        'Legacy equality alone would have accepted this wrong cache'
      );
  }
});
