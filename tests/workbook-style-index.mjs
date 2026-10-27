import assert from 'node:assert/strict';
import test from 'node:test';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyWorkbookStyle, createWorkbookStyleWriter } from '../core/workbook-styles.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const font = '<font><name val="Calibri"/><sz val="11"/></font>',
  fill = '<fill><patternFill patternType="none"/></fill>',
  xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
const source = `<styleSheet xmlns="${NS}"><fonts count="2">${font}${font}</fonts><fills count="2">${fill}${fill}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1">${xf}</cellStyleXfs><cellXfs count="2">${xf}${xf}</cellXfs></styleSheet>`;
const serialize = (node) => new XMLSerializer().serializeToString(node);
const operations = [
  { cell: 'A1', base: 1, edit: {} },
  ...Array.from({ length: 96 }, (_, i) => ({
    cell: `B${i % 23}`,
    base: i % 2,
    edit: {
      bold: i % 2 === 0,
      italic: i % 3 === 0,
      fontFamily: i % 4 === 0 ? '宋体 & "A"' : 'Calibri',
      fontSize: 10 + (i % 4),
      fontColor: (0xaabb00 + (i % 17)).toString(16),
      fillColor: (0x110000 + (i % 11)).toString(16),
      border: i % 2 ? 'all' : 'none',
      numberFormat: ['general', 'decimal1', 'currency', 'percent', 'date'][i % 5],
      horizontalAlign: i % 2 ? 'left' : 'center',
      wrapText: i % 2 === 0,
    },
  })),
];
function run(cached, options, xml = source) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml'),
    cells = new Map(),
    ids = [],
    write = cached ? createWorkbookStyleWriter(doc, options) : (cell, edit) => applyWorkbookStyle(doc, cell, edit);
  for (const op of operations) {
    let cell = cells.get(op.cell);
    if (!cell) {
      cell = doc.createElementNS(NS, 'c');
      cell.setAttribute('s', String(op.base));
      cells.set(op.cell, cell);
    }
    write(cell, op.edit);
    ids.push(Number(cell.getAttribute('s')));
  }
  return { xml: serialize(doc), ids, stats: cached ? write.diagnostics() : undefined };
}
test('batch style index preserves exact XML, cumulative edits and first duplicate IDs', () => {
  const expected = run(false),
    actual = run(true);
  assert.equal(actual.xml, expected.xml);
  assert.deepEqual(actual.ids, expected.ids);
  assert.equal(actual.ids[0], 0, 'The first identical source XF must win');
  assert.equal(actual.stats.disabled, false);
  assert(actual.stats.indexedEntries > 10);
  assert(actual.stats.estimatedIndexBytes <= 4 * 1024 * 1024);
});
test('style index budget exhaustion falls back without changing styles or IDs', () => {
  const expected = run(false);
  for (const options of [{ maxBytes: 0 }, { maxBytes: 512 }, { maxEntries: 12 }, { maxEntries: 128 }]) {
    const actual = run(true, options);
    assert.equal(actual.xml, expected.xml);
    assert.deepEqual(actual.ids, expected.ids);
    assert.equal(actual.stats.disabled, true);
    assert.equal(actual.stats.indexedEntries, 0);
    assert.equal(actual.stats.estimatedIndexBytes, 0);
    assert(actual.stats.peakEstimatedIndexBytes <= (options.maxBytes ?? 4 * 1024 * 1024));
  }
});
test('style indexes remain scoped to each parsed document and reject invalid budgets', () => {
  const changed = source.replaceAll('Calibri', 'Other Font');
  assert.equal(run(true, {}, changed).xml, run(false, {}, changed).xml);
  assert.equal(run(true).xml, run(false).xml);
  for (const options of [
    { maxBytes: -1 },
    { maxBytes: Infinity },
    { maxBytes: 4 * 1024 * 1024 + 1 },
    { maxEntries: -1 },
    { maxEntries: 16385 },
  ])
    assert.throws(
      () => createWorkbookStyleWriter(new DOMParser().parseFromString(source, 'application/xml'), options),
      { code: 'style-malformed' }
    );
});

test('a small style batch does not eagerly index a large existing table', () => {
  const xml = source.replace(/<cellXfs[^>]*>[\s\S]*?<\/cellXfs>/, `<cellXfs count="1024">${xf.repeat(1024)}</cellXfs>`),
    doc = new DOMParser().parseFromString(xml, 'application/xml'),
    cell = doc.createElementNS(NS, 'c');
  cell.setAttribute('s', '0');
  const write = createWorkbookStyleWriter(doc);
  write(cell, { bold: true });
  const stats = write.diagnostics();
  assert.equal(stats.styleWrites, 1);
  assert.equal(stats.indexedEntries, 0);
  assert.equal(stats.estimatedIndexBytes, 0);
});

test('indexed metadata accepts styles appended earlier in the same batch', async () => {
  const { createWorkbook, indexedStyleMetadata } = await import('../core/workbook-package.mjs');
  const result = indexedStyleMetadata(
    createWorkbook(),
    [
      { sheet: 's', cell: 'A1', bold: true },
      { sheet: 's', cell: 'A2', styleId: 1, italic: true },
      { sheet: 's', cell: 'A3', styleId: 2 },
    ],
    {}
  );
  assert.deepEqual(
    result.changes.map((c) => c.styleId),
    [1, 2, 2]
  );
  const doc = new DOMParser().parseFromString(result.replacements['xl/styles.xml'], 'application/xml');
  assert.equal(doc.getElementsByTagNameNS(NS, 'cellXfs')[0].getAttribute('count'), '3');
});
