import test from 'node:test';
import assert from 'node:assert/strict';
import { indexCsv, csvCell, csvCellInfo, csvReadRange, applyCsvCommands, serializeCsv } from '../core/csv-workbook.mjs';
test('indexed CSV retains quoted fields, embedded newlines, mixed record endings and missing versus empty cells', () => {
  const text = 'id,note,empty\r\n001,"中文\nline with ""quote""",\n002,"",tail\r003,last';
  const index = indexCsv(text);
  assert.equal(index.rows, 4);
  assert.equal(index.columns, 3);
  assert.equal(csvCell(index, 1, 0), '001');
  assert.equal(csvCell(index, 1, 1), '中文\nline with "quote"');
  assert.equal(csvCell(index, 1, 2), '');
  assert.equal(csvCell(index, 2, 1), '');
  assert.equal(csvCell(index, 3, 2), null);
  assert.equal(csvCellInfo(index, 1, 0).rawValue, '001');
  assert.equal(serializeCsv(index), text);
  const edited = applyCsvCommands(index, [{ sheet: 'CSV', cell: 'B3', expectedValue: '', value: 'x,y🙂' }]);
  assert.equal(
    serializeCsv(edited.index),
    'id,note,empty\r\n001,"中文\nline with ""quote""",\n002,"x,y🙂",tail\r003,last'
  );
  assert.equal(serializeCsv(index), text);
});
test('CSV batches are atomic, sequential guards observe earlier writes, typed values become literal text', () => {
  const index = indexCsv('id,value\na,1\nb,2\n'),
    before = serializeCsv(index);
  const good = applyCsvCommands(index, [
    { sheet: 'CSV', cell: 'B2', expectedValue: '1', value: 3 },
    { sheet: 'CSV', cell: 'B2', expectedValue: '3', value: '=SUM(A1)' },
    { type: 'setCells', sheet: 'CSV', cells: [{ cell: 'B3', expectedValue: '2', value: true }] },
  ]);
  assert.equal(csvCell(good.index, 1, 1), '=SUM(A1)');
  assert.equal(csvCell(good.index, 2, 1), 'true');
  assert.equal(good.edits.length, 3);
  assert.equal(good.affected.length, 2);
  assert.throws(
    () =>
      applyCsvCommands(index, [
        { sheet: 'CSV', cell: 'B2', value: 'first' },
        { sheet: 'CSV', cell: 'B3', expectedValue: 'bad', value: 'second' },
      ]),
    /Conflict/
  );
  assert.equal(serializeCsv(index), before);
  assert.throws(() => applyCsvCommands(index, [{ sheet: 'CSV', cell: 'B2', formula: 'A1+1' }]), {
    code: 'csv-bad-input',
  });
});
test('keyed CSV updates reject missing/duplicate keys and honor exact string identity', () => {
  const index = indexCsv('id,value\n001,10\n002,20\n'),
    command = {
      type: 'updateRows',
      sheet: 'CSV',
      range: 'A2:B3',
      keyColumn: 'A',
      updates: [{ key: '001', cells: [{ column: 'B', expectedValue: '10', value: 15 }] }],
    };
  const updated = applyCsvCommands(index, [command]);
  assert.equal(csvCell(updated.index, 1, 1), '15');
  assert.equal(csvCell(index, 1, 1), '10');
  assert.throws(
    () => applyCsvCommands(index, [{ ...command, updates: [{ ...command.updates[0], key: 1 }] }]),
    { code: 'internal-missing-field' }
  );
  assert.throws(() => applyCsvCommands(indexCsv('id,value\n001,10\n001,20'), [command]), {
    code: /^(csv-duplicate|duplicates-duplicate|workbook-duplicate)$/,
  });
});
test('indexed rows and writes have no100k-row ceiling and far cells remain accessible', () => {
  const text = Array.from({ length: 100005 }, (_, row) => `${row},value-${row}`).join('\n'),
    index = indexCsv(text);
  assert.equal(index.rows, 100005);
  assert.equal(csvCell(index, 100004, 1), 'value-100004');
  const { items, total } = csvReadRange(index, 'A1:B100005', { offset: 200008, limit: 2 });
  assert.equal(total, 200010);
  assert.equal(items[1].value, 'value-100004');
  const edited = applyCsvCommands(index, [
    { sheet: 'CSV', cell: 'B100005', expectedValue: 'value-100004', value: '末尾🙂' },
  ]);
  assert(serializeCsv(edited.index).endsWith('100004,末尾🙂'));
  assert.equal(csvCell(index, 100004, 1), 'value-100004');
});
test('CSV grows rows/columns explicitly and an edit restored to source retains original lexical bytes', () => {
  const original = indexCsv('"a",b'),
    changed = applyCsvCommands(original, [{ sheet: 'CSV', cell: 'A1', value: 'x' }]).index;
  assert.equal(serializeCsv(applyCsvCommands(changed, [{ sheet: 'CSV', cell: 'A1', value: 'a' }]).index), '"a",b');
  const grown = applyCsvCommands(original, [{ sheet: 'CSV', cell: 'D3', value: 'last' }]).index;
  assert.equal(grown.rows, 3);
  assert.equal(grown.columns, 4);
  assert.equal(csvCell(grown, 1, 0), '');
  assert.equal(csvCell(grown, 2, 2), '');
  assert.equal(serializeCsv(grown), '"a",b\n\n,,,last');
});
test('CSV ignores absent optional format fields but refuses meaningful formatting', () => {
  const original = indexCsv('id;note\r\n001;"a,b"\r\n', { delimiter: 'auto' });
  assert.equal(original.delimiter, ';');
  const changed = applyCsvCommands(original, [
    {
      sheet: 'CSV',
      cell: 'B2',
      value: 'x',
      formula: undefined,
      bold: undefined,
      italic: undefined,
      numberFormat: undefined,
      styleId: undefined,
    },
  ]);
  assert.equal(serializeCsv(changed.index), 'id;note\r\n001;"x"\r\n');
  assert.throws(() => applyCsvCommands(original, [{ sheet: 'CSV', cell: 'B2', value: 'x', bold: false }]), {
    code: 'csv-bad-input',
  });
});

test('new rows use record endings rather than quoted field newlines', () => {
  const text = '"embedded\nline",x\r\nnext,y\r\n';
  const index = indexCsv(text);
  assert.equal(index.newline, '\r\n');
  const extended = applyCsvCommands(index, [{ sheet: 'CSV', cell: 'A3', value: 'new' }]).index;
  assert.equal(serializeCsv(extended), text + 'new\r\n');
  assert.equal(serializeCsv(index), text);
});
