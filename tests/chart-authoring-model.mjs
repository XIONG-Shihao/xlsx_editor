import { test } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeChartModel, readChartData, createChartRecommendations } from '../core/chart-authoring-model.mjs';
import { CHART_STYLES, chartSeriesAppearance, chartSeriesPropertiesXml } from '../core/chart-styles.mjs';
const view = { sheets: [{ name: 'Data', hiddenRows: [] }] },
  cells = {
    A1: { value: 'Category', dataType: 'text' },
    B1: { value: 'Sales', dataType: 'text' },
    C1: { value: 'Costs', dataType: 'text' },
    A2: { value: 'Alpha', dataType: 'text' },
    A3: { value: 'Beta', dataType: 'text' },
    A4: { value: 'Gamma', dataType: 'text' },
    B2: { value: 2, dataType: 'number', rawValue: '2' },
    B3: { value: 5, dataType: 'number', rawValue: '5' },
    B4: { value: 8, dataType: 'number', rawValue: '8' },
    C2: { value: 1, dataType: 'number', rawValue: '1' },
    C3: { value: 3, dataType: 'number', rawValue: '3' },
    C4: { value: 6, dataType: 'number', rawValue: '6' },
  };
const recommendations = (more) =>
  createChartRecommendations({
    sheet: 'Data',
    range: 'A1:C4',
    view,
    readRange: async () => ({ cells: structuredClone(cells) }),
    ...more,
  });
test('three concrete series presets and closed normalized models preserve the legacy unstyled profile', () => {
  assert.equal(normalizeChartModel({ kind: 'column', range: 'A1:C4' }).style, undefined);
  for (const style of CHART_STYLES) {
    const c = normalizeChartModel({ kind: 'bar', range: 'A1:C4', style: style.id });
    assert.equal(c.style, style.id);
    assert(chartSeriesPropertiesXml(style.id, 0).includes('<c:spPr>'));
    assert.equal(chartSeriesAppearance(style.id, 1).width, 2);
  }
  assert.equal(new Set(CHART_STYLES.map((s) => chartSeriesPropertiesXml(s.id, 0))).size, 3);
  assert.throws(() => normalizeChartModel({ kind: 'pie', range: 'A1:C4' }));
  assert.throws(() => normalizeChartModel({ kind: 'column', range: 'A1:C4', style: 'foreign' }));
  assert.throws(() => normalizeChartModel({ kind: 'column', range: 'A1:C501' }));
  assert.throws(() => chartSeriesAppearance('outline', 8));
  assert.equal(chartSeriesPropertiesXml(undefined, 0), '');
});
test('ordinary selection yields three editable data-backed column/bar/line proposals without modifying source', async () => {
  const before = structuredClone(cells),
    r = await recommendations();
  assert.equal(r.proposals.length, 3);
  assert.deepEqual(
    r.proposals.map((p) => p.chart.kind),
    ['column', 'bar', 'line']
  );
  assert.deepEqual(
    r.proposals.map((p) => p.chart.style),
    ['colorful', 'monochrome', 'outline']
  );
  assert.equal(r.proposals[0].chart.title, 'Sales / Costs');
  for (const p of r.proposals) {
    assert.deepEqual(p.data.categories, ['Alpha', 'Beta', 'Gamma']);
    assert.deepEqual(
      p.data.series.map((s) => s.values),
      [
        [2, 5, 8],
        [1, 3, 6],
      ]
    );
    assert.equal(p.chart.range, 'A1:C4');
  }
  assert.deepEqual(cells, before);
});
test('long labels and increasing numeric categories change recommendation order from actual selected data', async () => {
  const long = structuredClone(cells);
  long.A2.value = 'A longer category label';
  let r = await recommendations({ readRange: async () => long });
  assert.equal(r.proposals[0].chart.kind, 'bar');
  const numeric = structuredClone(cells);
  for (const [i, ref] of ['A2', 'A3', 'A4'].entries())
    numeric[ref] = { value: (i + 1) * 10, dataType: 'number', rawValue: String((i + 1) * 10) };
  r = await recommendations({ readRange: async () => numeric });
  assert.equal(r.proposals[0].chart.kind, 'line');
  numeric.A3.value = 5;
  numeric.A3.rawValue = '5';
  r = await recommendations({ readRange: async () => numeric });
  assert.equal(r.proposals[0].chart.kind, 'column');
});
test('pending/error/unsafe/text-only or oversized selections are explained refusals, not cached-result guesses', async () => {
  for (const replacement of [
    { value: 8, formula: '1+1', dataType: 'number', calculation: { status: 'calculated', pending: true, value: 2 } },
    { value: 8, formula: '1+1', dataType: 'number', calculation: { status: 'stored', value: 2 } },
    { value: 2, dataType: 'number', rawValue: '2.00000000000000001' },
    { value: '2', dataType: 'text' },
    { value: '#DIV/0!', dataType: 'error' },
  ]) {
    const bad = structuredClone(cells);
    bad.B2 = replacement;
    await assert.rejects(recommendations({ readRange: async () => bad }));
  }
  await assert.rejects(recommendations({ range: 'A1:C501' }));
  const blank = structuredClone(cells);
  for (const col of ['B', 'C']) for (const row of [2, 3, 4]) delete blank[col + row];
  await assert.rejects(recommendations({ readRange: async () => blank }), /numeric/);
});
test('current calculated values, blank gaps, signed zero and declared hidden rows retain typed source truth', async () => {
  const input = structuredClone(cells);
  input.B2 = {
    value: 999,
    formula: '3+4',
    dataType: 'number',
    calculation: { status: 'calculated', pending: false, value: 7 },
  };
  delete input.B3;
  input.B4 = { value: -0, dataType: 'number', rawValue: '-0' };
  const data = await readChartData(
    async () => ({ cells: input }),
    'Data',
    normalizeChartModel({ kind: 'line', range: 'A1:C4' }),
    { sheets: [{ name: 'Data', hiddenRows: [2] }] }
  );
  assert.deepEqual(data.series[0].values.slice(0, 2), [7, null]);
  assert(Object.is(data.series[0].values[2], -0));
  assert.deepEqual(data.hidden, [false, true, false]);
});
