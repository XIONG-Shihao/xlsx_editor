/** The expected corpus is independently authored, never formatter-generated. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { formatCustomNumber, CUSTOM_NUMBER_FORMAT_LIMITS } from '../core/custom-number-format.mjs';
const corpus = JSON.parse(await fs.readFile(new URL('./support/custom-format-vectors.json', import.meta.url), 'utf8'));
assert.equal(corpus.schema, 'workbook.xlsx.custom-format-vectors.v1');
assert.equal(corpus.source, 'documentation-derived');
assert.equal(new Set(corpus.cases.map((item) => item.id)).size, corpus.cases.length);
for (const item of corpus.cases)
  test(item.id, () =>
    assert.deepEqual(formatCustomNumber(item.value, item.format, item.date1904 ?? false), item.expected)
  );
test('editor resource bounds and unsupported primitives preserve explicit null authority', () => {
  assert.equal(formatCustomNumber(7, '0'.repeat(CUSTOM_NUMBER_FORMAT_LIMITS.codeUnits + 1)), null);
  assert.equal(formatCustomNumber(7, '0.' + '0'.repeat(CUSTOM_NUMBER_FORMAT_LIMITS.fractionPlaces + 1)), null);
  assert.equal(formatCustomNumber('x'.repeat(CUSTOM_NUMBER_FORMAT_LIMITS.outputUnits + 1), '@'), null);
  for (const value of [NaN, Infinity, -Infinity, true, null, {}, []])
    assert.equal(formatCustomNumber(value, '0.00'), null);
  assert.equal(formatCustomNumber(7, 14), null);
  assert.equal(formatCustomNumber(7, '0.00', 'false'), null);
  for (const code of ['0\u0000', '0\\\ud800', '0\"\ud800\"']) assert.equal(formatCustomNumber(7, code), null);
});
test('mixed optional/required fraction positions and typed text remain pure', () => {
  assert.deepEqual(formatCustomNumber(8.9, '0.#0'), { text: '8.90', kind: 'number', supported: true });
  assert.deepEqual(formatCustomNumber('雪 😀', '0;0;0;"day: "@'), {
    text: 'day: 雪 😀',
    kind: 'text',
    supported: true,
  });
  const value = -0;
  assert.equal(Object.is(value, -0), true);
  assert.deepEqual(formatCustomNumber(value, '$#,##0.0000;[Red]($#,##0.0000)'), {
    text: '$0.0000',
    kind: 'number',
    supported: true,
  });
  assert.equal(Object.is(value, -0), true);
});

test('repeated text placeholders reject before any oversized string join', () => {
  const value = 'x'.repeat(32767),
    realJoin = Array.prototype.join,
    joinedUnits = [];
  Array.prototype.join = function (separator) {
    if (separator === '' && this.every((part) => typeof part === 'string')) {
      const units = this.reduce((total, part) => total + part.length, 0);
      if (units > CUSTOM_NUMBER_FORMAT_LIMITS.outputUnits) throw Error('Oversized display allocation attempted');
      joinedUnits.push(units);
    }
    return realJoin.call(this, separator);
  };
  try {
    // Prove the observer refuses an actual large join without allocating it.
    assert.throws(() => [value, value, value].join(''), /Oversized display allocation attempted/);
    assert.equal(formatCustomNumber(value, '@'.repeat(4096)), null);
    assert.equal(formatCustomNumber(value, '0;0;0;' + '@'.repeat(4090)), null);
    assert.deepEqual(joinedUnits, [], 'Rejected substitutions never reach the join allocation');
    const admitted = formatCustomNumber(value, '@@');
    assert.equal(admitted.text.length, 65534);
    assert.equal(admitted.kind, 'text');
    assert.deepEqual(joinedUnits, [65534], 'Only the admitted bounded output is concatenated');
    assert.equal(formatCustomNumber('x'.repeat(65535), '"!"@"!"'), null);
    assert.deepEqual(joinedUnits, [65534], 'Literal suffix is bounded before concatenation, too');
  } finally {
    Array.prototype.join = realJoin;
  }
});
