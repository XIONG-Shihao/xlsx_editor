/** Small, portable rule model. No owner paths or source XML cross this boundary. */
import { rangeBounds, cellPosition, columnName, assertXmlText } from './workbook-commands.mjs';
import { hasExactNumericValue } from './calculation.mjs';
import { refusal } from './refusal-codes.mjs';
export const RULE_LIMITS = { rules: 64, positions: 10000, listCharacters: 255 };
const own = (value, key) => Object.hasOwn(value, key);
function exact(value, keys) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid rule fields');
}
function text(value, max, label) {
  if (typeof value !== 'string' || value.length > max) throw refusal('internal-invalid-field', 'Invalid ' + label);
  assertXmlText(value);
  return value;
}
function bit(value) {
  if (typeof value !== 'boolean') throw refusal('internal-invalid-field', 'Invalid rule boolean');
  return value;
}
function number(value) {
  if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1e15)
    throw refusal('rules-bad-input', 'Rule numbers must be finite and within 15 digits');
  return Object.is(value, -0) ? 0 : value;
}
function range(value) {
  if (typeof value !== 'string' || !/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/.test(value))
    throw refusal('internal-invalid-field', 'Invalid rule range');
  const b = rangeBounds(value);
  if (b.size > RULE_LIMITS.positions) throw refusal('rules-limit', 'Rule area is limited to 10,000 positions');
  return columnName(b.first.col) + (b.first.row + 1) + ':' + columnName(b.last.col) + (b.last.row + 1);
}
export function validateConditionalRules(input) {
  if (!Array.isArray(input) || input.length > RULE_LIMITS.rules)
    throw refusal('rules-limit', 'At most 64 conditional rules are supported');
  const rules = input.map((value) => {
    exact(value, ['range', 'kind', 'values', 'text', 'fillColor', 'stopIfTrue']);
    const result = {
      range: range(value.range),
      kind: value.kind,
      fillColor: text(value.fillColor, 7, 'fill color').toUpperCase(),
      stopIfTrue: bit(value.stopIfTrue),
    };
    if (!/^#[0-9A-F]{6}$/.test(result.fillColor)) throw refusal('internal-invalid-field', 'Invalid fill color');
    if (['greaterThan', 'lessThan', 'equal', 'between'].includes(value.kind)) {
      if (
        !Array.isArray(value.values) ||
        value.values.length !== (value.kind === 'between' ? 2 : 1) ||
        own(value, 'text')
      )
        throw refusal('internal-invalid-field', 'Invalid rule criteria');
      result.values = value.values.map(number);
      if (value.kind === 'between' && result.values[0] > result.values[1])
        throw refusal('rules-limit', 'Rule minimum exceeds maximum');
    } else if (value.kind === 'textContains') {
      if (own(value, 'values')) throw refusal('internal-invalid-field', 'Invalid text criteria');
      result.text = text(value.text, 128, 'search text');
      if (!result.text) throw refusal('rules-bad-input', 'Search text must not be empty');
    } else if (value.kind === 'duplicateValues') {
      if (own(value, 'text') || own(value, 'values'))
        throw refusal('internal-invalid-field', 'Invalid duplicate criteria');
    } else throw refusal('rules-unsupported', 'Unsupported conditional rule');
    return result;
  });
  const duplicateRanges = [
    ...new Set(rules.filter((rule) => rule.kind === 'duplicateValues').map((rule) => rule.range)),
  ];
  if (duplicateRanges.reduce((total, range) => total + rangeBounds(range).size, 0) > 10000)
    throw refusal('rules-duplicate', 'Duplicate rules are limited to 10,000 total source positions');
  return rules;
}
export function validateValidationRules(input) {
  if (!Array.isArray(input) || input.length > RULE_LIMITS.rules)
    throw refusal('rules-limit', 'At most 64 validation rules are supported');
  const rules = input.map((value) => {
    exact(value, [
      'range',
      'kind',
      'values',
      'minimum',
      'maximum',
      'allowBlank',
      'showDropdown',
      'promptTitle',
      'prompt',
      'errorTitle',
      'error',
    ]);
    const result = {
      range: range(value.range),
      kind: value.kind,
      allowBlank: bit(value.allowBlank),
      showDropdown: bit(value.showDropdown),
      promptTitle: text(value.promptTitle, 32, 'prompt title'),
      prompt: text(value.prompt, 255, 'prompt'),
      errorTitle: text(value.errorTitle, 32, 'error title'),
      error: text(value.error, 225, 'error message'),
    };
    if (value.kind === 'list') {
      if (
        own(value, 'minimum') ||
        own(value, 'maximum') ||
        !Array.isArray(value.values) ||
        !value.values.length ||
        value.values.length > 128
      )
        throw refusal('internal-invalid-field', 'Invalid validation list');
      result.values = value.values.map((item) => {
        text(item, 255, 'list item');
        if (!item || /[,\r\n"]/u.test(item))
          throw refusal('rules-unsupported', 'List items cannot contain commas, quotes or line breaks');
        return item;
      });
      if (new Set(result.values).size !== result.values.length || result.values.join(',').length > 255)
        throw refusal('rules-limit', 'Validation list is duplicate or exceeds 255 characters');
    } else if (['wholeNumber', 'decimal', 'textLength'].includes(value.kind)) {
      if (own(value, 'values')) throw refusal('internal-invalid-field', 'Invalid numeric validation');
      result.minimum = number(value.minimum);
      result.maximum = number(value.maximum);
      if (
        result.minimum > result.maximum ||
        (value.kind !== 'decimal' &&
          (!Number.isSafeInteger(result.minimum) || !Number.isSafeInteger(result.maximum))) ||
        (value.kind === 'textLength' && (result.minimum < 0 || result.maximum > 32767))
      )
        throw refusal('internal-invalid-field', 'Invalid validation bounds');
      if (result.showDropdown) throw refusal('rules-unsupported', 'Dropdown only applies to lists');
    } else throw refusal('rules-unsupported', 'Unsupported validation rule');
    return result;
  });
  for (let i = 0; i < rules.length; i++)
    for (let j = 0; j < i; j++)
      if (rangesOverlap(rules[i].range, rules[j].range))
        throw refusal('rules-conflict', 'Validation ranges cannot overlap');
  return rules;
}
export function rangesOverlap(a, b) {
  const x = rangeBounds(a),
    y = rangeBounds(b);
  return (
    x.first.row <= y.last.row && x.last.row >= y.first.row && x.first.col <= y.last.col && x.last.col >= y.first.col
  );
}
export function ruleContains(rule, ref) {
  const b = rangeBounds(rule.range),
    p = cellPosition(ref);
  return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
}
export function acceptedRuleValue(cell) {
  if (cell?.calculation?.pending) return { qualified: false };
  if ((cell?.formulaType && cell.formulaType !== 'shared') || cell?.scenario || cell?.dataType === 'error')
    return { qualified: false };
  let value = cell?.value ?? null;
  if (cell?.formula) {
    if (cell.calculation?.status !== 'calculated') return { qualified: false };
    value = cell.calculation.value;
  }
  if (
    typeof value === 'number' &&
    (!Number.isFinite(value) || (!cell?.formula && !hasExactNumericValue(value, cell.rawValue)))
  )
    return { qualified: false };
  return { qualified: true, value };
}
export function duplicateRuleKey(value) {
  return JSON.stringify(value === null ? '' : value);
}
export function conditionalRuleMatches(rule, cell, duplicateCounts) {
  const record = acceptedRuleValue(cell);
  if (!record.qualified) return false;
  const value = record.value;
  if (rule.kind === 'duplicateValues') return (duplicateCounts?.[duplicateRuleKey(value)] ?? 0) > 1;
  if (rule.kind === 'textContains')
    return typeof value === 'string' && value.toLocaleLowerCase('en-US').includes(rule.text.toLocaleLowerCase('en-US'));
  if (typeof value !== 'number') return false;
  return rule.kind === 'greaterThan'
    ? value > rule.values[0]
    : rule.kind === 'lessThan'
      ? value < rule.values[0]
      : rule.kind === 'equal'
        ? value === rule.values[0]
        : value >= rule.values[0] && value <= rule.values[1];
}
export function conditionalFill(rules, ref, cell, duplicates = {}) {
  let fill;
  for (const rule of rules) {
    if (!ruleContains(rule, ref) || !conditionalRuleMatches(rule, cell, duplicates[rule.range])) continue;
    fill ??= rule.fillColor;
    if (rule.stopIfTrue) break;
  }
  return fill;
}
export function assertValidationEdits(rules, edits) {
  for (const edit of edits) {
    if (!own(edit, 'value') && !own(edit, 'formula')) continue;
    for (const rule of rules) {
      if (!ruleContains(rule, edit.cell)) continue;
      if (own(edit, 'formula') && edit.formula)
        throw refusal('rules-unsupported', 'Formula input in a validated area is not qualified; use a literal value');
      const value = edit.value ?? null,
        blank = value === null || value === '';
      let valid = blank && rule.allowBlank;
      if (!blank) {
        if (rule.kind === 'list') valid = typeof value === 'string' && rule.values.includes(value);
        else if (rule.kind === 'textLength') {
          const length = String(value).length;
          valid = length >= rule.minimum && length <= rule.maximum;
        } else
          valid =
            typeof value === 'number' &&
            Number.isFinite(value) &&
            (rule.kind !== 'wholeNumber' || Number.isSafeInteger(value)) &&
            value >= rule.minimum &&
            value <= rule.maximum;
      }
      if (!valid)
        throw refusal(
          'rules-refused',
          'Validation rejected ' +
            edit.cell +
            ': ' +
            (rule.error || 'Use ' + (rule.kind === 'list' ? 'a listed value' : rule.minimum + ' to ' + rule.maximum))
        );
    }
  }
}
export function containsFormula(range, text) {
  const start = range.split(':')[0],
    literal = text.replaceAll('~', '~~').replaceAll('*', '~*').replaceAll('?', '~?').replaceAll('"', '""');
  return 'NOT(ISERROR(SEARCH("' + literal + '",' + start + ')))';
}
