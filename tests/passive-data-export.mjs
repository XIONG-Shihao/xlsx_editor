import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { PassiveWorkbook } from '../adapters/passive-workbook.mjs';
import { dataOnlyWorkbook } from '../core/passive-data-export.mjs';
const XLSX = createRequire(import.meta.url)('xlsx'),
  sha = (value) => createHash('sha256').update(value).digest('hex'),
  fixture = path.resolve('fixtures/remaining-formats');
for (const sourceFormat of ['xlsm', 'xls', 'xlsb'])
  test('authentic ' + sourceFormat + ' explicit pure evaluation and new XLS/XLSB/XLSX scalar exports', async (t) => {
    const owner = new PassiveWorkbook(),
      source = path.join(
        fixture,
        sourceFormat === 'xlsm' ? 'vba-user-function.xlsm' : 'universal-content.' + sourceFormat
      ),
      before = sha(await fs.readFile(source));
    t.after(() => owner.close());
    await owner.open(source, { format: sourceFormat });
    let evaluation;
    if (sourceFormat === 'xlsm') {
      const modules = (await owner.vbaModules('0')).modules,
        module = modules.find((value) => value.name === 'Module1'),
        functions = await owner.vbaFunctions('0', module.id);
      assert.deepEqual(functions.functions, [{ name: 'MyFunc', parameters: [], returnType: 'numeric constant' }]);
      evaluation = await owner.evaluateVba('0', module.id, 'MyFunc', []);
      assert.equal(evaluation.value, 42);
      assert.equal(evaluation.sourceRevision, before);
      assert.equal(evaluation.execution, 'explicit-pure-numeric-only');
      assert.equal((await owner.read('0', 'Sheet1', 'A1')).cells.A1.calculation.status, 'stored');
      const document = modules.find((value) => value.kind !== 'standard');
      assert((await owner.vbaFunctions('0', document.id)).protectedReason);
      await assert.rejects(owner.evaluateVba('0', document.id, 'MyFunc', []), /document\/class/);
      await assert.rejects(owner.evaluateVba('0', module.id, 'MyFunc', [1]), /arguments/);
      await assert.rejects(owner.evaluateVba('guessed', module.id, 'MyFunc', []), /stale/);
    }
    for (const format of ['xls', 'xlsb', 'xlsx']) {
      const result = await owner.dataOnlyExport('0', { format });
      assert.equal(result.sourceRevision, before);
      assert.equal(result.format, format);
      assert(result.omissions.some((value) => value.includes('VBA')));
      assert(result.omissions.some((value) => value.includes('stored values')));
      assert.equal(sha(await fs.readFile(result.path)), result.revision);
      assert(result.path.endsWith('.' + format));
      const repeat = await owner.dataOnlyExport('0', { format });
      assert.equal(repeat.path, result.path);
      const data = XLSX.read(await fs.readFile(result.path), { type: 'buffer', bookVBA: true, cellFormula: true });
      assert(!data.vbaraw);
      for (const sheet of data.SheetNames)
        for (const [ref, cell] of Object.entries(data.Sheets[sheet]))
          if (!ref.startsWith('!')) assert.equal(cell.f, undefined);
      if (sourceFormat === 'xlsm') {
        assert.equal(data.Sheets.Sheet1.A1.v, 42);
        assert.equal(result.formulaValues, 1);
      } else {
        assert.equal(data.Sheets.Sheet1.B1.t, 's');
        assert.equal(data.Sheets.Sheet1.B1.v, 'String1');
        assert.equal(data.Sheets.Sheet1.C4.t, 'n');
        assert.equal(data.Sheets.Sheet1.C4.v, 11);
        assert.equal(result.formulaValues, 4);
      }
    }
    assert.equal(sha(await fs.readFile(source)), before);
    assert.equal((await owner.export('0')).revision, before);
    await assert.rejects(owner.dataOnlyExport('0', { format: 'xlsm' }), /format/);
    await assert.rejects(owner.dataOnlyExport('0', { format: 'xlsx', path: 'forged' }), /format/);
  });
const model = (entries) => ({
  sheets: [{ name: 'Data' }],
  records: new Map([['Data', entries.map(([ref, row, col, cell]) => ({ ref, row, col, cell }))]]),
});
test('new data-only writer preserves literal text/boolean/precision and refuses silent scalar changes', async () => {
  const ordinary = model([
    ['A1', 0, 0, { value: '00123', dataType: 'text' }],
    ['B1', 0, 1, { value: '=SUM(A1)', dataType: 'text' }],
    ['C1', 0, 2, { value: true, dataType: 'boolean' }],
    ['D1', 0, 3, { value: false, dataType: 'boolean' }],
    ['E1', 0, 4, { value: 0.1, dataType: 'number' }],
  ]);
  for (const format of ['xls', 'xlsb', 'xlsx']) {
    const output = dataOnlyWorkbook(XLSX, ordinary, format),
      read = XLSX.read(output.bytes, { type: 'buffer' });
    assert.equal(read.Sheets.Data.A1.v, '00123');
    assert.equal(read.Sheets.Data.B1.t, 's');
    assert.equal(read.Sheets.Data.B1.v, '=SUM(A1)');
    assert.equal(read.Sheets.Data.D1.v, false);
    assert(Object.is(read.Sheets.Data.E1.v, 0.1));
  }
  for (const cell of [
    { value: null, formula: 'UNKNOWN()', calculation: { status: 'unresolved' } },
    { value: null, dataType: 'error', error: 'malformed' },
    { value: 9007199254740992, dataType: 'number', rawValue: '9007199254740993' },
    { value: Infinity, dataType: 'number' },
    { value: 'x'.repeat(32768), dataType: 'text' },
  ])
    assert.throws(() => dataOnlyWorkbook(XLSX, model([['A1', 0, 0, cell]]), 'xlsx'), /refused/);
  assert.throws(
    () => dataOnlyWorkbook(XLSX, model([['IW1', 0, 256, { value: 1, dataType: 'number' }]]), 'xls'),
    /geometry/
  );
  assert.throws(
    () => dataOnlyWorkbook(XLSX, model([['A65537', 65536, 0, { value: 1, dataType: 'number' }]]), 'xls'),
    /geometry/
  );
  const signed = model([['A1', 0, 0, { value: -0, dataType: 'number' }]]),
    xls = dataOnlyWorkbook(XLSX, signed, 'xls');
  assert(Object.is(XLSX.read(xls.bytes, { type: 'buffer' }).Sheets.Data.A1.v, -0));
  for (const format of ['xlsb', 'xlsx']) assert.throws(() => dataOnlyWorkbook(XLSX, signed, format), /exact scalar/);
});
