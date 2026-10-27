/** Replay selected retained scripted Excel arithmetic, never execute VBA in Excel. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { evaluatePureVba } from '../core/pure-vba.mjs';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
test('eleven selected scripted Excel16.113 retained arithmetic cases agree with explicit Double UDF AST', async () => {
  const source = await fs.readFile('fixtures/excel-b6-runtime.xlsx'),
    oracleBytes = await fs.readFile('fixtures/excel-b6-runtime.json');
  assert.equal(hash(source), 'cd93179908e44e2f0ffea6b482ae9a967e8a61f5c528323939f256cc29bca4d5');
  assert.equal(hash(oracleBytes), 'aca596f77385a5f0438c7ee9dc088a1f9fc5b5d4e9a7c0aa35e30bc77cd2f55a');
  const oracle = JSON.parse(oracleBytes),
    xml = strFromU8(unzipSync(source)['xl/worksheets/sheet1.xml']),
    doc = new DOMParser().parseFromString(xml, 'application/xml'),
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    cells = new Map(Array.from(doc.getElementsByTagNameNS(NS, 'c')).map((cell) => [cell.getAttribute('r'), cell])),
    number = (ref) => Number(cells.get(ref).getElementsByTagNameNS(NS, 'v')[0].textContent),
    results = [];
  for (const index of [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12]) {
    const item = oracle.cases.find((value) => value.ref === 'D' + index),
      cell = cells.get(item.ref);
    assert.equal(cell.getElementsByTagNameNS(NS, 'f')[0].textContent, item.formula);
    const references = [...new Set(item.formula.match(/[A-Z]+[1-9]\d*/g) ?? [])],
      parameters = references.map((ref) => 'ByVal ' + ref.toLowerCase() + ' As Double').join(', '),
      formula = item.formula
        .replace(/[A-Z]+[1-9]\d*/g, (ref) => ref.toLowerCase())
        .replace(/(?<![A-Za-z_\d.])(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?(?![A-Za-z_\d.])/g, (value) => value + '#'),
      code =
        'Function Case' +
        item.ref +
        '(' +
        parameters +
        ') As Double\nCase' +
        item.ref +
        ' = ' +
        formula +
        '\nEnd Function';
    const result = evaluatePureVba(code, 'Case' + item.ref, references.map(number));
    assert(Object.is(result.value, item.expected), item.ref);
    assert(Object.is(result.value, number(item.ref)), item.ref + ' stored cache');
    results.push({
      ref: item.ref,
      worksheetFormula: item.formula,
      pureCode: code,
      args: references.map(number),
      value: result.value,
      expected: item.expected,
    });
  }
  // Excel's worksheet subtraction zero adjustment differs from plain Double
  // arithmetic; this engine deliberately has no worksheet recalculation binding.
  const simple = evaluatePureVba(
    'Function Difference() As Double\nDifference = 0.1# + 0.2# - 0.3#\nEnd Function',
    'Difference',
    []
  ).value;
  assert.notEqual(simple, oracle.cases.find((value) => value.ref === 'D9').expected);
});
