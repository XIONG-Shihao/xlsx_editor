/** Pure syntax/React boundary checks; runtime calculation has separate owner proof. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { serializeFunction, functionDefinition } from '../ui/functionCatalog.ts';
const n = (value) => ({ kind: 'number', value }),
  s = (value) => ({ kind: 'text', value }),
  b = (value) => ({ kind: 'boolean', value }),
  r = (value) => ({ kind: 'reference', value });
test('conditional aggregation serializes typed criteria and optional same-shaped sum ranges', () => {
  assert.equal(
    serializeFunction('SUMIF', [r('a1:a3'), s('Sales'), r('$B$1:$B$3')]).formula,
    'SUMIF(A1:A3,"Sales",$B$1:$B$3)'
  );
  assert.equal(serializeFunction('SUMIF', [r('A1:A3'), s('>=2')]).formula, 'SUMIF(A1:A3,">=2")');
  assert.equal(serializeFunction('COUNTIF', [r('A1:A3'), b(true)]).formula, 'COUNTIF(A1:A3,TRUE)');
  assert.equal(serializeFunction('COUNTIF', [r('A1:A3'), r('D1')]).formula, 'COUNTIF(A1:A3,D1)');
  assert.equal(serializeFunction('COUNTIF', [r('A1:A3'), s('a"b')]).formula, 'COUNTIF(A1:A3,"a""b")');
  assert.throws(() => serializeFunction('SUMIF', [r('A1:A3'), n(1), r('B1:B2')]), /same rows/);
  assert.throws(() => serializeFunction('COUNTIF', [r('A1:A1001'), n(1)]));
});
test('VLOOKUP adds exact FALSE, preserves scalar key types and bounds return-column literals', () => {
  const call = serializeFunction('VLOOKUP', [s('sku'), r('A1:C4'), n(3)]);
  assert.equal(call.formula, 'VLOOKUP("sku",A1:C4,3,FALSE)');
  assert.equal(call.argumentCount, 4);
  assert.equal(functionDefinition('VLOOKUP').maxArgs, 3);
  assert.equal(serializeFunction('VLOOKUP', [n(2.5), r('A1:B3'), r('D1')]).formula, 'VLOOKUP(2.5,A1:B3,D1,FALSE)');
  for (const index of [0, -1, 1.5, 3])
    assert.throws(() => serializeFunction('VLOOKUP', [s('sku'), r('A1:B3'), n(index)]));
  assert.throws(() => serializeFunction('VLOOKUP', [s('sku'), r('A1:B3'), n(2), b(true)]));
});
test('XLOOKUP uses exact first-match defaults with matching vertical or horizontal vectors and typed fallback', () => {
  assert.equal(serializeFunction('XLOOKUP', [s('abc'), r('A1:A4'), r('B1:B4')]).formula, 'XLOOKUP("abc",A1:A4,B1:B4)');
  assert.equal(
    serializeFunction('XLOOKUP', [n(2.25), r('A1:D1'), r('A2:D2'), s('')]).formula,
    'XLOOKUP(2.25,A1:D1,A2:D2,"")'
  );
  assert.equal(
    serializeFunction('XLOOKUP', [b(false), r('A1:A4'), r('B1:B4'), b(true)]).formula,
    'XLOOKUP(FALSE,A1:A4,B1:B4,TRUE)'
  );
  for (const [lookup, returns] of [
    ['A1:B2', 'C1:D2'],
    ['A1:A3', 'B1:B2'],
    ['A1:A3', 'B1:D1'],
  ])
    assert.throws(() => serializeFunction('XLOOKUP', [s('sku'), r(lookup), r(returns)]));
  assert.throws(() => serializeFunction('XLOOKUP', [n(1), r('A1:A3'), r('B1:B3'), n(0), n(1)]));
});
