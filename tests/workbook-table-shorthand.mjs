import { nativeWorkbookBinary } from '@office-editors/xlsx-editor/native';
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { canonicalizeTableShorthand as canonical } from '@office-editors/xlsx-editor/core/workbook-commands.mjs';
import { createWorkbook, editWorkbook, readWorkbook } from '@office-editors/xlsx-editor/core/workbook-package.mjs';
import { IndexedWorkbook } from '@office-editors/xlsx-editor/core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '@office-editors/xlsx-editor/core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '@office-editors/xlsx-editor/core/native-client.mjs';

export function tableFixture() {
  const cells = [
    ['A1', 'Value'],
    ['B1', 'Sales Amount'],
    ['C1', 'Result'],
    ['A2', 2],
    ['B2', 10],
    ['A3', 3],
    ['B3', 20],
    ['A4', 4],
    ['B4', 30],
  ];
  const populated = editWorkbook(
    createWorkbook(),
    cells.map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
  );
  return editWorkbook(populated, [{ sheet: 'Sheet1', operation: 'createTable', range: 'A1:C4', tableName: 'Prices' }]);
}
const commands = [
  { sheet: 'Sheet1', cell: 'C2', formula: 'Prices[@Value]*2' },
  { sheet: 'Sheet1', cell: 'C3', formula: '[@[Sales Amount]]+prices[@value]' },
  { sheet: 'Sheet1', cell: 'C4', formula: 'SUM(Prices[@Value],[@[Sales Amount]])' },
  { sheet: 'Sheet1', cell: 'D2', formula: 'Prices[@Value]+1' },
];
const expected = [
  'Prices[[#This Row],[Value]]*2',
  'Prices[[#This Row],[Sales Amount]]+Prices[[#This Row],[Value]]',
  'SUM(Prices[[#This Row],[Value]],Prices[[#This Row],[Sales Amount]])',
  'Prices[[#This Row],[Value]]+1',
];
test('table shorthand normalizes only proven current-row selectors and preserves quoted text', () => {
  const view = readWorkbook(tableFixture());
  commands.forEach((c, i) => assert.equal(canonical(c.formula, view, c.sheet, c.cell), expected[i]));
  const quoted = 'IF("literal ""[@Value]"""="Prices[@Value]",[@Value],\'sheet[@Value]\'!A1)';
  assert.equal(canonical(quoted, view, 'Sheet1', 'C2'), quoted.replace(',[@Value],', ',Prices[[#This Row],[Value]],'));
  assert.equal(canonical('SUM(Prices[Value])', view, 'Sheet1', 'C2'), 'SUM(Prices[Value])');
  for (const formula of [
    '[@Missing]',
    'Unknown[@Value]',
    '[@[Value]:[Sales Amount]]',
    '[@Value',
    "[@'@Value]",
    '表[@Value]',
    "'Sheet1'!Prices[@Value]",
    'Prices [@Value]',
  ])
    assert.throws(() => canonical(formula, view, 'Sheet1', 'C2'), /shorthand|serialization|qualifier|column/i, formula);
  for (const cell of ['C1', 'C5', 'E2'])
    assert.throws(() => canonical('[@Value]', view, 'Sheet1', cell), /table|data row/i, cell);
  assert.throws(() => canonical('Prices[@Value]', view, 'Other', 'C2'), /data row/);
  assert.throws(() => canonical('[@Value]', { ...view, tables: [...view.tables, ...view.tables] }, 'Sheet1', 'C2'), {
    code: 'table-bad-input',
  });
  assert.throws(
    () =>
      canonical(
        'Prices[@Value]',
        { ...view, tables: view.tables.map((t) => ({ ...t, totalsRowCount: 1 })) },
        'Sheet1',
        'C2'
      ),
    /data row/
  );
  assert.throws(() => canonical('[@Value]+'.repeat(800) + '1', view, 'Sheet1', 'C2'), { code: 'table-limit' });
});
test('shared compatibility writer serializes shorthand without changing source cells or table XML', () => {
  const source = tableFixture(),
    before = readWorkbook(source),
    saved = readWorkbook(editWorkbook(source, commands));
  commands.forEach((c, i) => assert.equal(saved.sheets[0].cells[c.cell].formula, expected[i]));
  assert.deepEqual(saved.tables, before.tables);
  for (const ref of ['A1', 'B1', 'C1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4'])
    assert.deepEqual(saved.sheets[0].cells[ref], before.sheets[0].cells[ref]);
});
test('indexed shorthand uses native calculation, canonical receipts, atomic guards and saved replay', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shorthand-')),
    source = path.join(root, 'source.xlsx'),
    original = tableFixture();
  await fs.writeFile(source, original);
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  let last;
  const make = () =>
    new IndexedWorkbook(
      client.call,
      async () => {
        throw Error('Unexpected whole-workbook writer');
      },
      async (input) => {
        last = await calculateIndexedWorkbook(input, client.call);
        return last;
      }
    );
  const book = make();
  let reopened, replayed;
  try {
    await book.open(source);
    const changed = await book.apply('0', commands),
      state = changed.view.storage.state;
    assert(last.nativeSqlite);
    assert.equal(last.inputCells, 0);
    assert.deepEqual(
      changed.edits.map((e) => e.formula),
      expected
    );
    for (const [cell, value] of [
      ['C2', 4],
      ['C3', 23],
      ['C4', 34],
      ['D2', 3],
    ])
      assert.equal((await book.read(state, 'Sheet1', cell)).cells[cell].calculation.value, value);
    assert.equal((await book.read('0', 'Sheet1', 'C2')).cells.C2, undefined);
    await assert.rejects(
      book.apply(state, [
        { sheet: 'Sheet1', cell: 'A2', value: 99 },
        { sheet: 'Sheet1', cell: 'C2', formula: '[@Missing]' },
      ]),
      /column/
    );
    assert.equal((await book.read(state, 'Sheet1', 'A2')).cells.A2.value, 2);
    const inputEdit = await book.apply(state, [{ sheet: 'Sheet1', cell: 'A2', value: 5 }]);
    assert(last.nativePlanHit);
    assert.equal((await book.read(inputEdit.view.storage.state, 'Sheet1', 'C2')).cells.C2.calculation.value, 10);
    const exported = await book.export(state),
      saved = path.join(root, 'saved.xlsx');
    await fs.copyFile(exported.path, saved);
    const disk = readWorkbook(await fs.readFile(saved));
    commands.forEach((c, i) => assert.equal(disk.sheets[0].cells[c.cell].formula, expected[i]));
    reopened = make();
    await reopened.open(saved);
    assert.equal((await reopened.read('0', 'Sheet1', 'C3')).cells.C3.calculation.value, 23);
    replayed = make();
    await replayed.open(source);
    const replay = await replayed.apply('0', changed.edits);
    assert.equal((await replayed.read(replay.view.storage.state, 'Sheet1', 'C4')).cells.C4.calculation.value, 34);
    assert.deepEqual(new Uint8Array(await fs.readFile(source)), new Uint8Array(original));
  } finally {
    await reopened?.close();
    await replayed?.close();
    await book.close();
    client.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});
