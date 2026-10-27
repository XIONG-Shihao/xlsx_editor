import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  accessibilityFields,
  accessibilityReport,
  bindAccessibility,
  validateAccessibility,
} from '../core/accessibility.mjs';
import { indexedAccessibility } from '../core/indexed-accessibility.mjs';
import { accessibilityFixture } from './support/accessibility-fixture.mjs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { columnName } from '../core/workbook-commands.mjs';
const receipt = { documentId: 'doc', ownerId: 'owner', state: 'head', revision: 'revision' },
  read = async (sheet) => ({
    cells:
      sheet === 'Sheet1'
        ? { A1: { value: 'Label' }, B1: { value: ' label ' }, C1: { value: '   ' } }
        : { A1: { value: 'Good' }, B1: { value: 'Distinct' }, C1: { value: 'Third' } },
  });
async function report(options = {}) {
  const f = accessibilityFixture(options);
  return indexedAccessibility(f.bytes, f.view, { readRange: read });
}
test('literal cell headers rather than table definition names produce exact blank and duplicate locations across sheets', async () => {
  const v = await report();
  assert.equal(v.scope, 'workbook');
  assert.equal(v.status, 'complete');
  assert.deepEqual(v.coverage.tables, { total: 2, checked: 2, skipped: 0 });
  assert.equal(v.coverage.headerCellsRead, 6);
  assert.deepEqual(
    v.findings.map((f) => [f.rule, f.sheet, f.cell, f.tableName, f.relatedCells]),
    [
      ['blank-table-header', 'Sheet1', 'C1', 'Data1', []],
      ['duplicate-table-header', 'Sheet1', 'A1', 'Data1', ['A1', 'B1']],
    ]
  );
  assert.deepEqual(v.coverage.objects, { total: 0, checked: 0, skipped: 0 });
});
test('readonly worksheet still allows findings without a repair grant', async () => {
  const v = await report({ protectedSheet: true });
  assert.equal(v.findings.length, 2);
  assert.equal(v.status, 'complete');
  assert.equal(Object.hasOwn(v, 'canEdit'), false);
});
test('extended, no-header and undeclared table graphs are skipped explicitly while another sheet remains checked', async () => {
  for (const options of [{ extended: true }, { headerless: true }, { orphan: true }]) {
    const v = await report(options);
    assert.equal(v.status, 'partial');
    assert.deepEqual(v.coverage.tables, { total: 2, checked: 1, skipped: 1 });
    assert.equal(v.findings.length, 0);
    assert.equal(v.coverage.skipped[0].sheet, 'Sheet1');
    assert.equal(v.coverage.skipped[0].name, 'Data1');
  }
});
test('formula/error/special headers and truncated native replies are never treated as ordinary cached labels or absent cells', async () => {
  for (const result of [
    { cells: { B1: { value: 'label', formula: '"label"' } } },
    { cells: { B1: { value: '#N/A', dataType: 'error' } } },
    { cells: {}, nextOffset: 5 },
    {},
  ]) {
    const f = accessibilityFixture({ twoSheets: false }),
      v = await indexedAccessibility(f.bytes, f.view, { readRange: async () => result });
    assert.equal(v.status, 'partial');
    assert.equal(v.findings.length, 0);
    assert.equal(v.coverage.tables.checked, 0);
    assert.equal(v.coverage.headerCellsRead, result.cells && result.nextOffset == null ? 3 : 0);
  }
});
test('original native head cancellation escapes inspection even when a range-read catch handles an ordinary failure', async () => {
  const f = accessibilityFixture();
  let retired = false;
  await assert.rejects(
    indexedAccessibility(f.bytes, f.view, {
      readRange: async () => {
        retired = true;
        return { cells: {} };
      },
      check: () => {
        if (retired) throw Error('Original owner retired');
      },
    }),
    /Original owner retired/
  );
});
test('reply validator rejects stale owners, hidden coverage, duplicate locations, malformed rules and caller paths', async () => {
  const v = await report(),
    bound = bindAccessibility(v, receipt);
  validateAccessibility(bound, receipt);
  v.findings.length = 0;
  assert.equal(bound.findings.length, 2);
  for (const key of Object.keys(receipt))
    assert.throws(() => validateAccessibility(bound, { ...receipt, [key]: 'other' }));
  for (const change of [
    (v) => v.coverage.tables.checked--,
    (v) => (v.coverage.objects.total = null),
    (v) => v.rules.reverse(),
    (v) => v.findings.push(v.findings[0]),
    (v) => (v.findings[1].relatedCells = ['A1']),
    (v) => (v.findings[0].cell = 'a1'),
    (v) => (v.filename = '/external'),
  ]) {
    const candidate = structuredClone(bound);
    change(candidate);
    assert.throws(() => validateAccessibility(candidate, receipt));
  }
  const partial = {
    ...bound,
    status: 'partial',
    coverage: {
      ...bound.coverage,
      skipped: [{ kind: 'table', sheet: 'Sheet1', name: 'Other', reason: 'Extended model' }],
    },
  };
  validateAccessibility(partial, receipt);
});
test('exact read transport fields exclude arbitrary ranges, sheets, repair commands and wrong indexed owner kinds', () => {
  const v = { operation: 'accessibility', documentId: 'doc', state: 'head' };
  accessibilityFields(v);
  accessibilityFields({ ...v, kind: 'indexed', sessionId: 'owner' }, { transport: true });
  for (const patch of [
    { sheet: 'Sheet1' },
    { range: 'A1:C3' },
    { state: '' },
    { operation: 'repairAccessibility' },
    { path: '/external' },
  ])
    assert.throws(() => accessibilityFields({ ...v, ...patch }));
  assert.throws(() => accessibilityFields({ ...v, kind: 'csv', sessionId: 'owner' }, { transport: true }));
});
test('table-count and aggregate header budgets give explicit skipped coverage and bounded reads/findings', async () => {
  for (const [count, width, checked] of [
    [18, 1000, 16],
    [65, 3, 64],
    [1, 1001, 0],
  ]) {
    const f = accessibilityFixture({ twoSheets: false }),
      parts = unzipSync(f.bytes),
      ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    f.view.tables = [];
    const links = [],
      rels = [];
    for (let i = 0; i < count; i++) {
      const name = 'Bounded' + i,
        row = i * 5 + 1,
        range = `A${row}:${columnName(width - 1)}${row + 2}`,
        part = `xl/tables/bounded${i}.xml`;
      f.view.tables.push({ name, sheet: 'Sheet1', ref: range, path: part });
      parts[part] = strToU8(
        `<table xmlns="${ns}" id="${i + 10}" name="${name}" displayName="${name}" ref="${range}"><tableColumns count="${width}">${Array.from({ length: width }, (_, c) => `<tableColumn id="${c + 1}" name="Label${c}"/>`).join('')}</tableColumns></table>`
      );
      links.push(`<tablePart r:id="b${i}"/>`);
      rels.push(`<Relationship Id="b${i}" Type="${rel}/table" Target="../tables/bounded${i}.xml"/>`);
    }
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${ns}" xmlns:r="${rel}"><sheetData/><tableParts count="${count}">${links.join('')}</tableParts></worksheet>`
    );
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`
    );
    let reads = 0;
    const v = await indexedAccessibility(zipSync(parts), f.view, {
      readRange: async () => {
        reads++;
        return { cells: {} };
      },
    });
    assert.equal(reads, checked);
    assert.equal(v.coverage.tables.checked, checked);
    assert.equal(v.coverage.tables.skipped, count - checked);
    assert.equal(v.coverage.headerCellsRead, checked * width);
    assert.equal(v.findings.length, checked * width);
    assert.equal(v.status, 'partial');
  }
});

test('an unreadable drawing inventory is unknown coverage, never an empty clean list, while table findings survive', async () => {
  const f = accessibilityFixture(),
    parts = unzipSync(f.bytes);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<tableParts', '<drawing r:id="missingDrawing"/><tableParts')
  );
  const v = await indexedAccessibility(zipSync(parts), f.view, { readRange: read });
  assert.equal(v.status, 'partial');
  assert.equal(v.coverage.objects.total, null);
  assert.equal(v.coverage.skipped[0].kind, 'objects');
  assert.match(v.coverage.skipped[0].reason, /inventory unavailable/);
  assert.equal(v.findings.length, 2);
  assert.equal(v.coverage.tables.checked, 2);
});
