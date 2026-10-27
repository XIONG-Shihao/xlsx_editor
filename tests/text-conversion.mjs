import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  convertTextField as convert,
  TEXT_CONVERSION_POLICIES,
  TEXT_CONVERSION_LIMITS,
} from '../core/text-conversion.mjs';
function success(raw, policy, value, kind, options = {}) {
  const result = convert(raw, policy, options);
  assert.equal(result.status, 'converted', JSON.stringify(result));
  assert.equal(result.kind, kind);
  assert(Object.is(result.value, value));
  return result;
}
function refuse(raw, policy, options = {}) {
  const result = convert(raw, policy, options);
  assert.equal(result.status, 'rejected', raw);
  assert.equal(Object.hasOwn(result, 'value'), false);
  return result;
}
test('Text default preserves literal formulas, whitespace, Unicode and empty text without type guessing', () => {
  for (const raw of ['00123', '1.00', '=SUM(A1)', 'TRUE', '#N/A', '2024-02-29', '  x\t\n', '中文', '😀', '']) {
    const result = success(raw, undefined, raw, 'text');
    assert.deepEqual(result.notices, []);
    assert.equal(Object.hasOwn(result, 'numberFormat'), false);
    assert(Object.isFrozen(result) && Object.isFrozen(result.notices));
  }
  assert.deepEqual(TEXT_CONVERSION_POLICIES, ['text', 'number', 'boolean', 'iso-date']);
});
test('Explicit Number retains decimal meaning and signed zero; lexical changes are disclosed', () => {
  for (const [raw, value] of [
    ['00123', 123],
    ['1.00', 1],
    ['+1', 1],
    ['.5', 0.5],
    ['1.', 1],
    ['1e3', 1000],
    ['-0.00', -0],
    ['0.1000', 0.1],
  ]) {
    const result = success(raw, 'number', value, 'number');
    assert.equal(result.notices.length, 1);
    assert.match(result.notices[0], /spelling changes/);
  }
  assert.equal(success('1', 'number', 1, 'number').notices.length, 0);
  assert.deepEqual(success(' 1 ', 'number', 1, 'number').notices, ['Outer whitespace removed.']);
  success('9007199254740991', 'number', 9007199254740991, 'number');
  success('5e-324', 'number', 5e-324, 'number');
  success('1.' + '0'.repeat(32000), 'number', 1, 'number');
});
test('Number rejects coercion, locale syntax, overflow, unsafe integers, rounding and zero underflow', () => {
  for (const raw of [
    '0x10',
    '0b10',
    '1,234',
    '1 234',
    'NaN',
    'Infinity',
    '-Infinity',
    '=1+2',
    'TRUE',
    '1/2',
    '€1',
    '1e309',
    '-1e309',
    '9007199254740992',
    '9007199254740993',
    '0.123456789012345678901',
    '1.0000000000000001',
    '1e-324',
    '1e-999',
    '4.9406564584124654e-324',
    '   ',
  ])
    refuse(raw, 'number');
});
test('Boolean uses only explicit TRUE/FALSE grammar and discloses trim/case changes', () => {
  for (const [raw, value] of [
    ['TRUE', true],
    ['FALSE', false],
    ['true', true],
    ['False', false],
    [' true ', true],
  ])
    success(raw, 'boolean', value, 'boolean');
  assert.deepEqual(success('TRUE', 'boolean', true, 'boolean').notices, []);
  assert.equal(success(' true ', 'boolean', true, 'boolean').notices.length, 2);
  for (const raw of ['1', '0', 'yes', 'no', 'T', 'F', '=TRUE()', 'true false', '   ']) refuse(raw, 'boolean');
});
test('ISO Date validates real Gregorian days before DATE normalization and respects both calendars', () => {
  for (const [raw, date1904, serial] of [
    ['1900-01-01', false, 1],
    ['1900-02-28', false, 59],
    ['1900-03-01', false, 61],
    ['1904-01-01', false, 1462],
    ['1904-01-01', true, 0],
    ['1904-02-29', true, 59],
    ['1904-03-01', true, 60],
    ['9999-12-31', false, 2958465],
    ['9999-12-31', true, 2957003],
  ]) {
    const result = success(raw, 'iso-date', serial, 'date', { date1904 });
    assert.equal(result.numberFormat, 'date');
    assert.equal(result.notices.length, 1);
  }
  for (const raw of [
    '1900-02-29',
    '2100-02-29',
    '2023-02-29',
    '2024-04-31',
    '2024-00-01',
    '2024-13-01',
    '2024-01-00',
    '2024-01-32',
    '0000-01-01',
    '10000-01-01',
    '2024-2-01',
    '02/01/2024',
    '2024-01-01T00:00:00Z',
    '=DATE(2024,1,1)',
  ])
    refuse(raw, 'iso-date', { date1904: false });
  refuse('1899-12-31', 'iso-date', { date1904: false });
  refuse('1903-12-31', 'iso-date', { date1904: true });
  refuse('2024-02-29', 'iso-date');
  const a = success('2024-02-29', 'iso-date', 45351, 'date', { date1904: false }),
    b = success('2024-02-29', 'iso-date', 43889, 'date', { date1904: true });
  assert.equal(a.value - b.value, 1462);
});
test('Empty inputs and ragged padding stay explicit text for all selected types; whitespace is distinct', () => {
  for (const policy of TEXT_CONVERSION_POLICIES) {
    const empty = success('', policy, '', 'text');
    assert.equal(empty.empty, true);
    assert.equal(empty.padded, false);
    const padded = success('', policy, '', 'text', { padded: true });
    assert.equal(padded.empty, true);
    assert.equal(padded.padded, true);
    assert.deepEqual(padded.notices, ['Padding kept as empty text.']);
    refuse('1', policy, { padded: true });
  }
  success('   ', 'text', '   ', 'text');
});
test('Per-field failures and size/character/closed-option bounds are explicit without mutation', () => {
  assert.equal(TEXT_CONVERSION_LIMITS.textLength, 32767);
  success('x'.repeat(32767), 'text', 'x'.repeat(32767), 'text');
  refuse('x'.repeat(32768), 'text');
  for (const raw of ['\0', '\u000b', '\ud800', '\udfff', '\ufffe', '\uffff']) refuse(raw, 'text');
  refuse(1, 'number');
  refuse('1', 'auto');
  for (const options of [null, false, [], { extra: 1 }, { padded: 1 }, { date1904: 'false' }])
    refuse('1', 'number', options);
  const options = Object.freeze({ date1904: false, padded: false });
  success('1900-01-01', 'iso-date', 1, 'date', options);
  assert.deepEqual(options, { date1904: false, padded: false });
});
