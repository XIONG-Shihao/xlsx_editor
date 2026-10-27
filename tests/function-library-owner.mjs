/** Function-library scalar batch on a real native Document: Undo/Redo, recovery, Save and an
 * independent check of the saved cached values. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import {
  functionLibraryOwnerCases as functionLibraryCases,
  functionLibraryFixture,
  functionLibraryExpected,
} from './support/function-library-fixture.mjs';

const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-function-library-owner-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  python = process.env.XLSX_DESKTOP_PYTHON ?? 'python3',
  root = path.resolve(import.meta.dirname, '..'),
  copies = [];
const client = createWorkbookNativeClient(nativeWorkbookBinary()),
  owners = [];
const native = (request) => client.call(request);
const make = (recoveryDirectory) => {
  const owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null });
  owners.push(owner);
  return owner;
};
const status = (owner) => owner.request({ operation: 'status' });
const read = async (owner, range) => {
  const s = await status(owner);
  return owner.request({ operation: 'read', documentId: s.documentId, state: s.state, sheet: 'Sheet1', range });
};
function value(cell, c, calendar) {
  assert.equal(cell.formula, c.formula);
  if (c.type === 'error') {
    assert.equal(cell.calculation.status, 'error');
    assert.equal(cell.calculation.error, c.value);
  } else {
    assert.equal(cell.calculation.status, 'calculated');
    const wanted = c.type === 'date' ? (calendar === 1900 ? 36585 : 35123) : c.value;
    assert.equal(typeof cell.calculation.value, typeof wanted);
    assert.equal(cell.calculation.value, wanted);
  }
}
function oracle(source, saved, expected) {
  return JSON.parse(
    execFileSync(python, [path.join(root, 'tests/support/check-function-library.py'), source, saved, expected], {
      encoding: 'utf8',
      maxBuffer: 1024 * 1024,
    })
  );
}
try {
  for (const calendar of [1900, 1904]) {
    const directory = path.join(output, String(calendar));
    await fs.mkdir(directory);
    const original = path.join(directory, 'original.xlsx'),
      filename = path.join(directory, 'working.xlsx'),
      recovery = path.join(directory, 'recovery'),
      expected = path.join(directory, 'expected.json'),
      bytes = functionLibraryFixture({ calendar });
    await fs.writeFile(original, bytes);
    await fs.writeFile(filename, bytes);
    await fs.writeFile(
      expected,
      JSON.stringify(functionLibraryExpected(calendar, { ownerCatalog: true }), null, 2) + '\n'
    );
    let owner = make(recovery);
    await owner.open(filename);
    const initial = await status(owner);
    const accepted = await owner.request({
      operation: 'apply',
      documentId: initial.documentId,
      state: initial.state,
      commands: [
        {
          type: 'setCells',
          sheet: 'Sheet1',
          cells: functionLibraryCases.map((c) => ({ cell: c.cell, formula: c.formula })),
        },
      ],
    });
    const result = await read(owner, 'C1:D15');
    for (const c of functionLibraryCases) value(result.cells[c.cell], c, calendar);
    assert.notEqual(accepted.state, initial.state);
    assert.equal(accepted.canUndo, true);
    assert.equal(hash(await fs.readFile(filename)), hash(bytes));
    await owner.request({ operation: 'undo' });
    assert.equal((await status(owner)).state, initial.state);
    const undone = await read(owner, 'C1:D15');
    for (const c of functionLibraryCases) assert.equal(undone.cells[c.cell], undefined);
    assert.equal((await status(owner)).canRedo, true);
    await owner.request({ operation: 'redo' });
    assert.equal((await status(owner)).state, accepted.state);
    await owner.close();
    owner = make(recovery);
    await owner.open(filename);
    const recovered = await status(owner);
    assert.equal(recovered.recovered, true);
    assert.equal(recovered.state, accepted.state);
    for (const c of functionLibraryCases) value((await read(owner, c.cell)).cells[c.cell], c, calendar);
    await owner.request({ operation: 'undo' });
    assert.equal((await status(owner)).state, initial.state);
    await owner.request({ operation: 'redo' });
    assert.equal((await status(owner)).state, accepted.state);
    const copy = path.join(directory, 'copy.xlsx');
    await owner.request({ operation: 'copy', destination: copy, expectedRevision: null });
    const copyReadback = oracle(original, copy, expected);
    assert.equal(hash(await fs.readFile(filename)), hash(bytes));
    await owner.request({ operation: 'save' });
    assert.equal((await status(owner)).dirty, false);
    const saveReadback = oracle(original, filename, expected);
    assert.equal(hash(await fs.readFile(original)), hash(bytes));
    await owner.close();
    const cold = make(path.join(directory, 'cold-recovery'));
    await cold.open(filename);
    assert.equal((await status(cold)).dirty, false);
    for (const c of functionLibraryCases) value((await read(cold, c.cell)).cells[c.cell], c, calendar);
    await cold.close();
    copies.push({ original, copy, expected });
  }
  const calendar = copies[0],
    parts = unzipSync(await fs.readFile(calendar.copy)),
    worksheet = 'xl/worksheets/sheet1.xml',
    sheet = strFromU8(parts[worksheet]);
  const mutations = [
    ['boolean-as-number', /\bt="b"/, 't="n"'],
    ['text-as-error', /\bt="str"/, 't="e"'],
    [
      'date-binary-bit',
      /(<v(?:\s[^>]+)?>)36585(?:\.0)?(<\/v>)/,
      (_all, open, close) => open + '36585.00000000001' + close,
    ],
  ];
  for (const [name, pattern, replacement] of mutations) {
    assert(pattern.test(sheet), name + ' actually mutates saved output');
    const changed = { ...parts, [worksheet]: strToU8(sheet.replace(pattern, replacement)) },
      file = path.join(output, name + '.xlsx');
    await fs.writeFile(file, zipSync(changed));
    let failure;
    try {
      oracle(calendar.original, file, calendar.expected);
    } catch (error) {
      failure = String(error.stderr ?? error.message);
    }
    assert(failure, name + ' must fail strict independent cache proof');
  }
} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
