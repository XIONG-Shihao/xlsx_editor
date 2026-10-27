import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  validateConditionalRules,
  validateValidationRules,
  conditionalFill,
  containsFormula,
  assertValidationEdits,
} from '../core/rules-model.mjs';
const cf = (kind, extra = {}) => ({ range: 'B2:B9', kind, fillColor: '#FF0000', stopIfTrue: false, ...extra });
const dv = (kind, extra = {}) => ({
  range: 'B2:B9',
  kind,
  allowBlank: true,
  showDropdown: false,
  promptTitle: '',
  prompt: '',
  errorTitle: 'Invalid',
  error: 'Use permitted input',
  ...extra,
});
test('closed rule forms normalize criteria and refuse invalid bounds, overlapping areas and extra fields', () => {
  assert.equal(validateConditionalRules([cf('between', { values: [2, 5] })])[0].range, 'B2:B9');
  assert.throws(() => validateConditionalRules([cf('between', { values: [5, 2] })]), {
    code: /^(fill-limit|rules-limit)$/,
  });
  assert.throws(() => validateConditionalRules([cf('equal', { values: [1], formula: 'A1' })]), /fields/);
  assert.throws(() => validateConditionalRules([cf('equal', { values: [Infinity] })]), /finite/);
  assert.throws(() => validateConditionalRules([{ ...cf('duplicateValues'), range: 'A1:A10001' }]), /10,000/);
  assert.throws(() => validateValidationRules([dv('list', { values: ['Yes,No'] })]), { code: 'rules-unsupported' });
  assert.throws(
    () =>
      validateValidationRules([
        dv('decimal', { minimum: 0, maximum: 2 }),
        dv('wholeNumber', { minimum: 0, maximum: 2 }),
      ]),
    /overlap/
  );
  assert.throws(() => validateValidationRules([dv('textLength', { minimum: -1, maximum: 2 })]), /bounds/);
});
test('conditional priorities, exact typed duplicates and literal case-insensitive text are portable', () => {
  const rules = validateConditionalRules([
    cf('greaterThan', { values: [10], fillColor: '#FF0000' }),
    cf('between', { values: [2, 20], fillColor: '#00FF00' }),
    cf('textContains', { text: 'A*?~"', fillColor: '#0000FF' }),
  ]);
  assert.equal(conditionalFill(rules, 'B3', { value: 12, rawValue: '12' }), '#FF0000');
  assert.equal(conditionalFill(rules, 'B3', { value: 'a*?~"' }), '#0000FF');
  assert.equal(
    conditionalFill(rules, 'B3', { value: 99, formula: 'UNKNOWN()', calculation: { status: 'unresolved' } }),
    undefined
  );
  assert.equal(conditionalFill(rules, 'B3', { value: 12, rawValue: '12.0000000000000001' }), undefined);
  assert.equal(
    conditionalFill([cf('duplicateValues')], 'B3', { value: '1' }, { 'B2:B9': { '"1"': 2, 1: 1 } }),
    '#FF0000'
  );
  assert.equal(
    conditionalFill([cf('duplicateValues')], 'B3', { value: 1, rawValue: '1' }, { 'B2:B9': { '"1"': 2, 1: 1 } }),
    undefined
  );
  assert.equal(containsFormula('B2:B9', 'A*?~"'), 'NOT(ISERROR(SEARCH("A~*~?~~""",B2)))');
});
test('validation checks exact literal lists, blank policy, integer/decimal/text length and refuses formulas', () => {
  const list = validateValidationRules([dv('list', { values: ['Yes', 'No'], showDropdown: true, allowBlank: false })]);
  assertValidationEdits(list, [{ cell: 'B2', value: 'Yes' }]);
  assert.throws(() => assertValidationEdits(list, [{ cell: 'B2', value: 'yes' }]), { code: 'rules-refused' });
  assert.throws(() => assertValidationEdits(list, [{ cell: 'B2', value: null }]), { code: 'rules-refused' });
  assert.throws(() => assertValidationEdits(list, [{ cell: 'B2', formula: '"Yes"' }]), /Formula/);
  assertValidationEdits(list, [
    { cell: 'A2', value: 'any' },
    { cell: 'B2', bold: true },
  ]);
  for (const kind of ['wholeNumber', 'decimal', 'textLength']) {
    const rules = validateValidationRules([dv(kind, { minimum: 2, maximum: 5 })]);
    assertValidationEdits(rules, [{ cell: 'B2', value: kind === 'textLength' ? 'four' : 3 }]);
    assert.throws(() => assertValidationEdits(rules, [{ cell: 'B2', value: kind === 'textLength' ? 'a' : 6 }]), {
      code: 'rules-refused',
    });
  }
});
