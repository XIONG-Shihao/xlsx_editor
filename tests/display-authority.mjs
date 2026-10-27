/** Independent typed-authority examples; formatting cannot change result status. */
import { passiveCell } from '../core/passive-model.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { displayCell, displayCellDetails, formatDisplayDetails, formatDisplayValue } from '../core/display.mjs';
const money = '"$"#,##0.0000;[Red]("$"#,##0.0000)',
  quad = '0;0;0;"text="@';
const freeze = (value) => {
  if (value && typeof value === 'object') {
    Object.freeze(value);
    for (const child of Object.values(value)) freeze(child);
  }
  return value;
};
test('currency uses exact selected mask and named color, leaving the frozen source cell unchanged', () => {
  const c = freeze({ value: -7, rawValue: '-7', dataType: 'number', format: money }),
    before = JSON.stringify(c);
  assert.deepEqual(displayCellDetails(c), { text: '($7.0000)', formatStatus: 'supported', color: 'FF0000' });
  assert.equal(displayCell(c), '($7.0000)');
  assert.equal(JSON.stringify(c), before);
  assert.equal(displayCell({ value: 7, format: money }), '$7.0000');
  assert.equal(displayCell({ value: 0, format: money }), '$0.0000');
});
test('built-in IDs preserve the existing calendars while authored numeric-looking strings are custom codes', () => {
  assert.equal(formatDisplayValue(60, 14, false), '1900-02-29');
  assert.equal(formatDisplayValue(60, 14, true), '1904-03-01');
  assert.equal(formatDisplayValue(7, '14'), '7');
  assert.equal(formatDisplayDetails(7, '14').formatStatus, 'unsupported');
  assert.equal(formatDisplayValue(7.8, '0'), '8');
});
test('genuine literal and formula errors bypass a text section while error-looking text retains its type', () => {
  for (const c of [
    { value: '#N/A', dataType: 'error', format: quad },
    { formula: '1/0', calculation: { status: 'error', error: '#DIV/0!' }, format: quad },
  ])
    assert.equal(displayCellDetails(c).formatStatus, 'not-applicable');
  assert.equal(displayCell({ value: '#N/A', dataType: 'error', format: quad }), '#N/A');
  assert.equal(displayCell({ value: '#N/A', dataType: 'text', format: quad }), 'text=#N/A');
  assert.equal(
    displayCell({ formula: '"#N/A"', calculation: { status: 'calculated', value: '#N/A' }, format: quad }),
    'text=#N/A'
  );
});
test('unsafe raw literal digits are preserved before formatting and text-section interpolation', () => {
  const c = freeze({ value: 1e20, rawValue: '100000000000000000001', dataType: 'number', format: quad });
  assert.deepEqual(displayCellDetails(c), {
    text: '100000000000000000001',
    formatStatus: 'not-applicable',
    reason: 'precision-preserved',
  });
});
test('pending cache formatting retains Stored labels and never grants current result authority', () => {
  const c = freeze({
      formula: 'A1',
      format: money,
      calculation: { status: 'stored', pending: true },
      cachedResult: { status: 'stored', value: -7, rawValue: '-7' },
    }),
    before = JSON.stringify(c);
  assert.deepEqual(displayCellDetails(c), { text: 'Stored: ($7.0000)', formatStatus: 'supported', color: 'FF0000' });
  assert.equal(c.calculation.pending, true);
  assert.equal(c.calculation.value, undefined);
  assert.equal(JSON.stringify(c), before);
});
test('pending unsafe caches and typed errors are not decorated by text masks', () => {
  const base = { formula: 'A1', format: quad, calculation: { status: 'stored', pending: true } };
  assert.equal(
    displayCell({ ...base, cachedResult: { status: 'stored', value: 1e20, rawValue: '100000000000000000001' } }),
    'Stored: 100000000000000000001'
  );
  assert.equal(displayCell({ ...base, cachedResult: { status: 'error', error: '#N/A' } }), 'Stored: #N/A');
});
test('pending-without-cache, unresolved, circular and not-yet-calculated keep their own labels', () => {
  const base = { formula: 'A1', format: quad };
  assert.equal(displayCell({ ...base, calculation: { status: 'stored', pending: true } }), 'Pending calculation');
  assert.equal(displayCell({ ...base, calculation: { status: 'unresolved' } }), 'Unresolved');
  assert.equal(displayCell({ ...base, calculation: { status: 'circular' } }), '#CYCLE!');
  assert.equal(displayCell(base), 'Calculating…');
});
test('blank absence/null and explicit empty Text stay distinct without changing stored values', () => {
  assert.equal(displayCell({ value: null, format: quad }), '');
  assert.equal(displayCell({ format: quad }), '');
  assert.equal(displayCell({ value: '', dataType: 'text', format: quad }), 'text=');
  assert.equal(displayCell({ formula: 'A1', calculation: { status: 'calculated', value: null }, format: quad }), '');
  assert.equal(displayCell({ value: false, dataType: 'boolean', format: quad }), 'false');
});
test('unsupported patterns return the raw value and explicit status instead of guessed dates or precision', () => {
  for (const code of ['[h]:mm', '*d0.00', '[$USD-9999]0.00', '[Color1]0.00', '0.00E+00', '0 "unterminated']) {
    assert.deepEqual(
      formatDisplayDetails(7, code),
      { text: '7', formatStatus: 'unsupported', reason: 'unsupported-custom-format' },
      code
    );
    assert.equal(displayCell({ value: 7, rawValue: '7', format: code }), '7');
  }
});

test('real passive stored errors and exact digits stay authoritative, while calculated results ignore stale source types', () => {
  const error = freeze(passiveCell({ v: 7, t: 'e', f: '1/0', z: quad }, 'stored-error')),
    before = JSON.stringify(error);
  assert.deepEqual(displayCellDetails(error), { text: '#DIV/0!', formatStatus: 'not-applicable' });
  assert.equal(JSON.stringify(error), before);
  const text = passiveCell({ v: '#DIV/0!', t: 's', f: '"#DIV/0!"', z: quad }, 'stored-text');
  assert.equal(displayCell(text), 'text=#DIV/0!');
  const numeric = freeze(
    passiveCell({ v: 9007199254740992, t: 'n', f: 'A1', z: '0.00' }, 'stored-number', {
      type: 'n',
      hasValue: true,
      raw: '9007199254740993',
      formula: 'A1',
    })
  );
  assert.deepEqual(displayCellDetails(numeric), {
    text: '9007199254740993',
    formatStatus: 'not-applicable',
    reason: 'precision-preserved',
  });
  assert.equal(displayCell({ ...numeric, calculation: { status: 'calculated', value: 8 } }), '8.00');
  assert.equal(displayCell({ ...error, calculation: { status: 'calculated', value: '#DIV/0!' } }), 'text=#DIV/0!');
  assert.equal(
    displayCell({ ...numeric, format: quad, calculation: { status: 'calculated', value: 'fresh' } }),
    'text=fresh'
  );
  const invalid = passiveCell({ v: 7, t: 'n', f: '', z: '0.00' }, 'empty-formula');
  assert.equal(invalid.calculation.status, 'unresolved');
  assert.deepEqual(displayCellDetails(invalid), { text: 'Unresolved', formatStatus: 'not-applicable' });
});
