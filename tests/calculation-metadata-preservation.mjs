import test from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, indexedCalculationMetadata } from '../core/workbook-package.mjs';

test('unchanged calculation policy retains the exact source metadata envelope', () => {
  const parts = unzipSync(createWorkbook());
  const calc = '<calcPr calcId="162194" calcMode="manual" calcOnSave="0" fullCalcOnLoad="1" forceFullCalc="1"/>';
  const raw = strFromU8(parts['xl/workbook.xml']).replace('</workbook>', calc + '</workbook>');
  parts['xl/workbook.xml'] = strToU8(raw);
  const result = indexedCalculationMetadata(zipSync(parts), true, { mode: 'manual', recalculateBeforeSave: false });
  assert.equal(result.replacements['xl/workbook.xml'], raw);
  assert.deepEqual(result.removed, []);
  const changed = indexedCalculationMetadata(zipSync(parts), false, { mode: 'auto', recalculateBeforeSave: true });
  const next = changed.replacements['xl/workbook.xml'];
  for (const attribute of [
    'calcId="162194"',
    'calcMode="auto"',
    'calcOnSave="1"',
    'fullCalcOnLoad="0"',
    'forceFullCalc="0"',
  ])
    assert(next.includes(attribute));
  assert.equal(next.replace(/<calcPr\b[^>]*\/>/, ''), raw.replace(calc, ''));
});
