import test from 'node:test';
import assert from 'node:assert/strict';
import { createBlobDelimitedSource } from '../core/delimited-import.mjs';
import {
  inspectLocalQuery,
  localQueryPages,
  queryCellFingerprint,
  queryCellIdentity,
  validateLocalQuery,
} from '../core/local-query.mjs';
const source = (text) => createBlobDelimitedSource(new Blob([text]), { name: 'data.csv', chunkBytes: 7 }),
  query = () => ({
    columns: [3, 1],
    filters: [{ column: 2, op: 'equals', value: 'yes', caseSensitive: false }],
    keepHeader: true,
  });
async function cells(source, inspection) {
  const output = [];
  for await (const page of localQueryPages(source, inspection)) output.push(...page.cells);
  return output;
}
test('real select/filter preserves quoted multiline/literal fields and exact output order', async () => {
  const data = source('id,flag,value\r\n00123,YES,"=SUM(A1)"\r\n00234,no,skip\r\n00345,yes,"two\nlines"\r\n'),
    inspection = await inspectLocalQuery(data, { sheet: 'Sheet1', start: 'C4', query: query() });
  assert.equal(inspection.outputRows, 3);
  assert.equal(inspection.dataRows, 2);
  assert.deepEqual(inspection.preview, [
    ['value', 'id'],
    ['=SUM(A1)', '00123'],
    ['two\nlines', '00345'],
  ]);
  assert.deepEqual(await cells(data, inspection), [
    [0, 'value'],
    [1, 'id'],
    [2, '=SUM(A1)'],
    [3, '00123'],
    [4, 'two\nlines'],
    [5, '00345'],
  ]);
  const hash = queryCellFingerprint();
  for (const [index, value] of await cells(data, inspection)) hash.add(index, 'text', value);
  assert.equal(hash.finish(), inspection.outputFingerprint);
  hash.close();
});
test('smaller and empty results replace the prior managed tail with explicit empty text', async () => {
  const data = source('id,flag,value\n01,no,z\n'),
    definition = { ...query(), keepHeader: false },
    inspection = await inspectLocalQuery(data, {
      sheet: 'Sheet1',
      start: 'A1',
      query: definition,
      previousShape: { sheet: 'Sheet1', start: 'A1', rows: 4, columns: 3 },
    });
  assert.equal(inspection.outputRows, 0);
  assert.deepEqual(inspection.shape, { sheet: 'Sheet1', start: 'A1', rows: 4, columns: 3 });
  assert.deepEqual(
    await cells(data, inspection),
    Array.from({ length: 12 }, (_, index) => [index, ''])
  );
});
test('unknown expressions, duplicate columns, missing columns and wrong prior target refuse', async () => {
  assert.throws(() => validateLocalQuery({ ...query(), sql: 'DROP TABLE' }), /Invalid/);
  assert.throws(() => validateLocalQuery({ ...query(), columns: [1, 1] }), /Invalid/);
  assert.throws(() => validateLocalQuery({ ...query(), filters: [{ ...query().filters[0], op: 'eval' }] }), /Invalid/);
  await assert.rejects(
    inspectLocalQuery(source('one,two\nx,y'), { sheet: 'Sheet1', start: 'A1', query: query() }),
    /columns/
  );
  await assert.rejects(
    inspectLocalQuery(source('a,b,c\nx,yes,z'), {
      sheet: 'Sheet1',
      start: 'A1',
      query: query(),
      previousShape: { sheet: 'Other', start: 'A1', rows: 1, columns: 2 },
    }),
    { code: 'query-refused' }
  );
});
test('preview truncates without truncating accepted data, while target fingerprints distinguish literals/formulas/numbers', async () => {
  const text = 'z'.repeat(2000),
    data = source('a,b,c\nx,yes,' + text),
    inspection = await inspectLocalQuery(data, { sheet: 'Sheet1', start: 'A1', query: query() });
  assert(inspection.previewTruncated);
  assert.equal(inspection.preview[1][0].length, 256);
  assert.equal((await cells(data, inspection))[2][1], text);
  assert.notDeepEqual(queryCellIdentity({ value: '01' }), queryCellIdentity({ value: 1 }));
  assert.notDeepEqual(queryCellIdentity({ value: '=1+2' }), queryCellIdentity({ value: 3, formula: '1+2' }));
  assert.deepEqual(queryCellIdentity({}), ['text', '']);
});
test('large filtered source remains bounded and yields independently expected dense pages', async () => {
  const count = 100000,
    data = source(
      'id,flag,value\n' +
        Array.from(
          { length: count },
          (_, index) => `${String(index).padStart(6, '0')},${index % 1000 ? 'no' : 'yes'},=${index}`
        ).join('\n')
    ),
    inspection = await inspectLocalQuery(data, { sheet: 'Sheet1', start: 'B2', query: query() });
  assert.equal(inspection.outputRows, 101);
  const output = await cells(data, inspection);
  assert.equal(output.length, 202);
  for (let index = 0; index < 100; index++) {
    assert.deepEqual(output[2 + index * 2], [2 + index * 2, '=' + index * 1000]);
    assert.deepEqual(output[3 + index * 2], [3 + index * 2, String(index * 1000).padStart(6, '0')]);
  }
});
