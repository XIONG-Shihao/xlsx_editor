import test from 'node:test';
import assert from 'node:assert/strict';
import { zipSync, unzipSync } from 'fflate';
import { normalizeDataOnlyXlsb } from '../core/xlsb-values-normalize.mjs';
const record = (id, body) => Uint8Array.from([id, body.length, ...body]),
  row = (number) => {
    const bytes = new Uint8Array(17);
    new DataView(bytes.buffer).setUint32(0, number, true);
    return record(0, bytes);
  },
  pack = (records) => zipSync({ 'xl/worksheets/sheet1.bin': Uint8Array.from(records.flatMap((bytes) => [...bytes])) }),
  full = (column) => {
    const bytes = new Uint8Array(16);
    new DataView(bytes.buffer).setUint32(0, column, true);
    new DataView(bytes.buffer).setFloat64(8, 1.5, true);
    return record(5, bytes);
  };
test('fresh XLSB normalization makes adjacent column/type/IEEE bytes explicit without changing original bytes', () => {
  const source = pack([
      row(0),
      full(0),
      record(15, Uint8Array.from([0, 0, 0, 0, 1])),
      record(18, new Uint8Array(8)),
      record(14, Uint8Array.from([0, 0, 0, 0, 7, 0, 0, 0])),
    ]),
    before = new Uint8Array(source),
    output = normalizeDataOnlyXlsb(source),
    bytes = unzipSync(output)['xl/worksheets/sheet1.bin'];
  assert.deepEqual(source, before);
  let at = 0,
    result = [];
  while (at < bytes.length) {
    const id = bytes[at++],
      length = bytes[at++],
      body = bytes.subarray(at, at + length);
    at += length;
    result.push({ id, body });
  }
  assert.deepEqual(
    result.map((value) => value.id),
    [0, 5, 4, 7, 3]
  );
  assert.equal(new DataView(result[1].body.buffer, result[1].body.byteOffset).getFloat64(8, true), 1.5);
  for (let index = 1; index < result.length; index++)
    assert.equal(new DataView(result[index].body.buffer, result[index].body.byteOffset).getUint32(0, true), index - 1);
  assert.equal(result[2].body[8], 1);
  assert.equal(result[4].body.length, 9);
  assert.equal(result[4].body[8], 7);
});
test('short cells without a predecessor, bad sizes/order/padding and unexpected formulas are refused', () => {
  for (const records of [
    [row(0), record(15, new Uint8Array(5))],
    [row(0), full(0), record(15, new Uint8Array(6))],
    [row(0), full(1), full(0)],
    [row(1), full(0), row(1)],
    [row(0), full(0), record(14, Uint8Array.from([0, 0, 0, 0, 7, 1, 0, 0]))],
    [row(0), record(9, new Uint8Array(16))],
    [row(0), full(16384)],
  ])
    assert.throws(() => normalizeDataOnlyXlsb(pack(records)), { code: 'fill-bad-input' });
  assert.throws(() => normalizeDataOnlyXlsb(zipSync({ 'xl/worksheets/sheet1.bin': new Uint8Array([5, 255]) })), {
    code: 'fill-bad-input',
  });
});
