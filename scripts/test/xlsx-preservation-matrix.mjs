/** Bounded preserving-writer qualification; supplied and generated files stay distinct. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { WorkbookSession } from '@office-editors/xlsx-editor/core/workbook-service.mjs';
import { readWorkbook } from '@office-editors/xlsx-editor/core/workbook-package.mjs';
import { checkXlsxPackage } from './package-oracle.mjs';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { componentPath } from '../support/benchmark-paths.mjs';
const hash = (file) => createHash('sha256').update(fs.readFileSync(file)).digest('hex');

const output = componentPath('.build/runs/preservation-matrix');
fs.mkdirSync(output, { recursive: true });
const corpus = path.join(output, 'generated');
const generated = spawnSync('python3', [componentPath('fixtures/build-corpus.py'), '--output', corpus], {
  encoding: 'utf8',
});
assert.equal(generated.status, 0, generated.stderr);
const cases = [
  {
    id: 'pricing-schedule',
    sheet: 'Pricing',
    cell: 'B2',
    value: 3,
    expected: {
      'Pricing!B2': { value: 3 },
      'Pricing!D2': { formula: 'B2*C2', calculated: 375 },
      'Pricing!F4': { formula: 'SUM(F2:F3)', calculated: 495.95 },
    },
  },
  {
    id: 'budget',
    sheet: 'Budget',
    cell: 'B2',
    value: 3,
    expected: {
      'Budget!B2': { value: 3 },
      'Budget!D4': { formula: 'SUM(D2:D3)', calculated: 300 },
      'Budget!F2': { formula: "'Rates'!B1*D4", calculated: 30 },
    },
  },
  {
    id: 'response-matrix',
    sheet: 'Responses',
    cell: 'C2',
    value: 'Compliant',
    expected: { 'Responses!C2': { value: 'Compliant' } },
  },
  {
    id: 'dates-1900',
    sheet: 'Dates',
    cell: 'B5',
    value: 'Reviewed',
    expected: { 'Dates!B5': { value: 'Reviewed' }, 'Dates!A2': { value: 60 } },
  },
  {
    id: 'dates-1904',
    sheet: 'Dates',
    cell: 'B5',
    value: 'Reviewed',
    expected: { 'Dates!B5': { value: 'Reviewed' }, 'Dates!B1': { formula: 'A1+1', calculated: 1 } },
  },
  {
    id: 'preservation-rich',
    sheet: 'Data',
    cell: 'D2',
    value: 'Reviewed',
    expected: { 'Data!D2': { value: 'Reviewed' } },
  },
];
const report = {
  date: new Date().toISOString(),
  runnerSha256: hash(import.meta.filename),
  cases: [],
  passed: false,
  realWorkbookCoverage:
    'Six independently generated OOXML cases. Optional supplied financial workbook requires XLSX_EDITOR_FINANCIAL_WORKBOOK; omitted means explicitly untested.',
};
try {
  for (const item of cases) {
    const source = path.join(corpus, `${item.id}.xlsx`),
      before = hash(source);
    const session = new WorkbookSession(fs.readFileSync(source)),
      view = await session.read();
    const result = await session.apply({
      operationId: `preservation-${item.id}`,
      expectedRevision: view.revision,
      commands: [{ sheet: item.sheet, cell: item.cell, value: item.value }],
    });
    const saved = path.join(output, `${item.id}-delivered.xlsx`);
    fs.writeFileSync(saved, result.bytes);
    const independent = checkXlsxPackage(source, saved, {
      targets: { [item.sheet]: [item.cell] },
      expected: item.expected,
    });
    assert.equal(hash(source), before);
    report.cases.push({
      id: item.id,
      origin: 'synthetic: independent stdlib OOXML generator',
      source,
      sourceSha256: before,
      saved,
      savedSha256: hash(saved),
      independent,
      passed: true,
    });
  }
  if (process.env.XLSX_EDITOR_FINANCIAL_WORKBOOK) {
    const source = path.resolve(process.env.XLSX_EDITOR_FINANCIAL_WORKBOOK),
      before = hash(source),
      bytes = fs.readFileSync(source);
    const view = readWorkbook(bytes);
    assert.equal(view.readOnly, false);
    const session = new WorkbookSession(bytes),
      accepted = await session.read();
    const edited = await session.apply({
      operationId: 'supplied-financial-literal',
      expectedRevision: accepted.revision,
      commands: [{ sheet: '项目成本', cell: 'D4', expectedValue: 254063.62, value: 254064.62 }],
    });
    const saved = path.join(output, 'financial-delivered.xlsx');
    fs.writeFileSync(saved, edited.bytes);
    const independent = checkXlsxPackage(source, saved, {
      targets: { 项目成本: ['D4'] },
      expected: { '项目成本!D4': { value: 254064.62 } },
    });
    await assert.rejects(
      session.apply({
        operationId: 'supplied-financial-structure',
        expectedRevision: edited.view.revision,
        commands: [{ sheet: '项目成本', operation: 'insertRow', index: 3 }],
      }),
      /unsupported|unavailable|cannot|requires|refus|unqualified|preserv/i
    );
    assert.equal(hash(source), before);
    report.cases.push({
      id: 'supplied-financial',
      origin: 'supplied real WPS workbook',
      source,
      sourceSha256: before,
      sheets: view.sheets.length,
      outcome: 'guarded literal update preserved; unsafe row insertion refused',
      saved,
      savedSha256: hash(saved),
      independent,
      limitations: view.limitations,
      passed: true,
    });
  } else
    report.cases.push({
      id: 'supplied-financial',
      status: 'not-run',
      reason: 'No explicit XLSX_EDITOR_FINANCIAL_WORKBOOK supplied',
    });
  report.passed = true;
  console.log(
    JSON.stringify({ passed: true, cases: report.cases.map(({ id, passed, status }) => ({ id, passed, status })) })
  );
} catch (error) {
  report.error = String(error);
  process.exitCode = 1;
  console.error(error);
} finally {
  fs.writeFileSync(path.join(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
}
