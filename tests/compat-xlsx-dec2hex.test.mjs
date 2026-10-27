import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readWorkbook } from '../core/workbook-package.mjs';
import { calculateWorkbook } from '../core/calculation-node.mjs';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');

test('portable DEC2HEX retains every actual Excel scalar result and protects all array/ANCHORARRAY cases', async () => {
  const oracle = JSON.parse(
    await fs.readFile(new URL('../fixtures/excel-dec2hex-reference.json', import.meta.url), 'utf8')
  );
  assert.equal(oracle.excel, '16.113.3');
  assert.equal(Object.keys(oracle.authoredFormulas).length, 78);
  const arrays = new Set(oracle.sourceArrayMasters);
  for (const [name, digest] of [
    [oracle.source, oracle.sourceSha256],
    [oracle.oracleWorkbook, oracle.oracleWorkbookSha256],
  ]) {
    const bytes = await fs.readFile(new URL(`../fixtures/${name}`, import.meta.url));
    assert.equal(hash(bytes), digest);
    const input = readWorkbook(bytes),
      original = structuredClone(input),
      calculated = await calculateWorkbook(input),
      cells = calculated.sheets[0].cells;
    let qualified = 0,
      protectedCases = 0;
    for (const ref of Object.keys(oracle.authoredFormulas)) {
      const expected = oracle.cells[ref],
        result = cells[ref].calculation;
      if (arrays.has(ref)) {
        assert.equal(result.status, 'unresolved', ref);
        protectedCases++;
        continue;
      }
      assert.equal(result.status, expected.dataType === 'error' ? 'error' : 'calculated', ref);
      assert.equal(expected.dataType === 'error' ? result.error : result.value, expected.expected, ref);
      if (expected.dataType === 'text') assert.equal(typeof result.value, 'string', ref);
      assert.equal(cells[ref].formula, input.sheets[0].cells[ref].formula, ref);
      qualified++;
    }
    assert.equal(qualified, 64);
    assert.equal(protectedCases, 14);
    assert.deepEqual(input, original);
  }
  assert.deepEqual(oracle.formulaNormalizations, {
    E106: { authored: 'DEC2HEX(-0)', saved: 'DEC2HEX(0)' },
    E109: { authored: 'DEC2HEX(+1)', saved: 'DEC2HEX(1)' },
  });
});
