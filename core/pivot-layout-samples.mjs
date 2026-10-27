/** Finite owned pivot presentation; cache values are independent of layout. */
import { refusal } from './refusal-codes.mjs';

export const PIVOT_LAYOUT_SAMPLES = Object.freeze([
  Object.freeze({
    id: 'tabular-light',
    label: 'Tabular Light',
    compact: 0,
    outline: 0,
    style: 'PivotStyleLight16',
    stripes: 0,
  }),
  Object.freeze({
    id: 'compact-blue',
    label: 'Compact Blue',
    compact: 1,
    outline: 1,
    style: 'PivotStyleMedium9',
    stripes: 1,
  }),
  Object.freeze({
    id: 'outline-dark',
    label: 'Outline Dark',
    compact: 0,
    outline: 1,
    style: 'PivotStyleDark2',
    stripes: 1,
  }),
]);
export function pivotLayoutSample(id) {
  const sample = PIVOT_LAYOUT_SAMPLES.find((s) => s.id === id);
  if (!sample) throw refusal('pivot-bad-input', 'Choose a supported pivot layout sample');
  return sample;
}
export function pivotPresentation(root) {
  const styles = Array.from(root.childNodes ?? []).filter(
    (n) => n.nodeType === 1 && n.localName === 'pivotTableStyleInfo'
  );
  const sample = PIVOT_LAYOUT_SAMPLES.find(
    (s) =>
      styles.length === 1 &&
      styles[0].getAttribute('name') === s.style &&
      root.getAttribute('compact') === String(s.compact) &&
      root.getAttribute('outline') === String(s.outline)
  );
  const total = root.getAttribute('colGrandTotals');
  if (!sample || !['0', '1'].includes(total) || root.getAttribute('rowGrandTotals') !== total)
    throw refusal('pivot-unsupported', 'Unsupported owned pivot layout or totals');
  return { sample: sample.id, grandTotals: total === '1' };
}
export function pivotOutputData(model, { sample = 'tabular-light', grandTotals = true } = {}) {
  const layout = pivotLayoutSample(sample);
  if (typeof grandTotals !== 'boolean')
    throw refusal('pivot-bad-input', 'Pivot grand totals require an explicit Boolean');
  return [
    [
      layout.compact ? 'Row Labels' : model.fields[model.rowField].name,
      (model.aggregate === 'sum' ? 'Sum of ' : 'Count of ') + model.fields[model.valueField].name,
    ],
    ...model.groups.map((g) => [g.label, g.value]),
    ...(grandTotals ? [['Grand Total', model.total]] : []),
  ];
}
