import test from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { calculateWorkbook, transformWorkbook } from '../core/calculation-node.mjs';
test('invalid local name scopes stay unavailable through JSON instead of becoming workbook names', () => {
  const parts = unzipSync(createWorkbook()),
    definition = (name, scope) =>
      `<definedName name="${name}"${scope === null ? '' : ` localSheetId="${scope}"`}>Sheet1!$A$1</definedName>`;
  const names = [
    ['Global', null],
    ['Local', '0'],
    ['LeadingZero', '00'],
    ['Empty', ''],
    ['Outside', '1'],
    ['Negative', '-1'],
    ['Fraction', '0.0'],
    ['NotNumber', 'NaN'],
    ['Overflow', '4294967296'],
  ];
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames>' + names.map(([name, scope]) => definition(name, scope)).join('') + '</definedNames></workbook>'
    )
  );
  const bytes = zipSync(parts),
    before = bytes.slice(),
    view = readWorkbook(bytes),
    roundtrip = JSON.parse(JSON.stringify(view)).namedRanges;
  assert.deepEqual(roundtrip[0], { name: 'Global', formula: 'Sheet1!$A$1' });
  for (const name of ['Local', 'LeadingZero'])
    assert.deepEqual(
      roundtrip.find((r) => r.name === name),
      { name, formula: 'Sheet1!$A$1', sheet: 'Sheet1' }
    );
  for (const name of ['Empty', 'Outside', 'Negative', 'Fraction', 'NotNumber', 'Overflow']) {
    const value = roundtrip.find((r) => r.name === name);
    assert.equal(value.scopeInvalid, true, name + ' must retain an explicit unavailable scope');
    assert.equal(Object.hasOwn(value, 'sheet'), false);
  }
  assert.deepEqual(bytes, before, 'Metadata read must preserve the package bytes');
});
test('an unavailable name cannot supply a numeric result or a dependent result in the JS engine', async () => {
  const view = {
    revision: 'scope-source',
    sheets: [
      {
        name: 'Main',
        rows: 10,
        columns: 4,
        cells: {
          A1: { value: 13 },
          B1: { formula: 'BadScope', value: 999, cachedValue: 999 },
          B2: { formula: 'B1+1' },
          C1: { formula: 'Good*2' },
        },
      },
    ],
    namedRanges: [
      { name: 'BadScope', formula: 'Main!$A$1', scopeInvalid: true },
      { name: 'Good', formula: 'Main!$A$1' },
    ],
  };
  const before = structuredClone(view),
    output = await calculateWorkbook(view);
  for (const ref of ['B1', 'B2']) {
    assert.notEqual(output.sheets[0].cells[ref].calculation.status, 'calculated');
    assert.equal(output.sheets[0].cells[ref].calculation.value, undefined);
  }
  assert.equal(output.sheets[0].cells.C1.calculation.value, 26);
  assert.deepEqual(view, before);
  await assert.rejects(transformWorkbook(view, [{ sheet: 'Main', operation: 'insertRow', index: 1 }]), {
    code: 'internal-invalid-field',
  });
  assert.deepEqual(view, before);
});
