/** Closed ordered sort keys. The first differing key decides; equal rows retain
 * source order. This validates command shape only, not source protection. */
import { refusal } from './refusal-codes.mjs';

const own = (value, key) => Object.hasOwn(value, key);
function columnIndex(column) {
  if (typeof column !== 'string' || !/^[A-Z]{1,3}$/i.test(column))
    throw refusal('internal-invalid-field', 'Invalid sort column');
  const index = [...column.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  if (index >= 16384) throw refusal('internal-invalid-field', 'Invalid sort column');
  return index;
}
function position(reference) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(reference);
  if (!match || Number(match[2]) > 1048576) throw refusal('internal-invalid-field', 'Invalid sort range');
  return { col: columnIndex(match[1]), row: Number(match[2]) - 1 };
}
export function normalizeSortKeys(edit) {
  if (!edit || typeof edit !== 'object' || Array.isArray(edit) || typeof edit.range !== 'string')
    throw refusal('internal-invalid-field', 'Invalid sort command');
  const refs = edit.range.split(':');
  if (refs.length > 2) throw refusal('internal-invalid-field', 'Invalid sort range');
  const first = position(refs[0]),
    last = position(refs[1] ?? refs[0]);
  if (first.row > last.row || first.col > last.col) throw refusal('internal-invalid-field', 'Invalid sort range');
  const multiple = own(edit, 'keys');
  if (multiple && (own(edit, 'column') || own(edit, 'direction')))
    throw refusal('sort-unsupported', 'Sort keys cannot be mixed with legacy column or direction');
  const keys = multiple ? edit.keys : [{ column: edit.column, direction: edit.direction }];
  if (!Array.isArray(keys) || keys.length < 1 || keys.length > 16)
    throw refusal('sort-bad-input', 'Sort requires 1–16 ordered keys');
  const seen = new Set();
  return Object.freeze(
    keys.map((key) => {
      if (
        !key ||
        typeof key !== 'object' ||
        Array.isArray(key) ||
        Object.keys(key).length !== 2 ||
        !own(key, 'column') ||
        !own(key, 'direction')
      )
        throw refusal('internal-invalid-field', 'Invalid sort key fields');
      const index = columnIndex(key.column);
      if (index < first.col || index > last.col)
        throw refusal('sort-unsupported', 'Sort key column is outside the requested range');
      if (seen.has(index)) throw refusal('sort-duplicate', 'Duplicate sort key column');
      seen.add(index);
      if (!['asc', 'desc'].includes(key.direction)) throw refusal('internal-invalid-field', 'Invalid sort direction');
      return Object.freeze({ column: key.column.toUpperCase(), direction: key.direction });
    })
  );
}
