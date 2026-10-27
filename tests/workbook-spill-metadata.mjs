import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { buildSpillMetadata } from '../core/spill-metadata.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const sheet = 'xl/worksheets/sheet1.xml',
  groups = [
    { schema: 1, sheet, master: 'A1', range: 'A1:B3', status: 'calculated' },
    { schema: 1, sheet, master: 'A90', range: 'A90', intendedRange: 'A90:C93', status: 'error', error: '#SPILL!' },
  ];
async function owner(t, bytes) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-native-spill-meta-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  await fs.writeFile(path.join(root, 'source.xlsx'), bytes);
  await native({ action: 'workbook', operation: 'index', root });
  return { root, check: (plan) => native({ action: 'workbook', operation: 'spill-metadata-check', root, plan }) };
}

test('native independent metadata admission validates JS append/reuse/new package closures', async (t) => {
  for (const bytes of [
    createWorkbook(),
    await fs.readFile(new URL('../fixtures/excel-dynamic-source.xlsx', import.meta.url)),
    await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
  ]) {
    const { root, check } = await owner(t, bytes),
      plan = buildSpillMetadata(bytes, groups),
      before = await fs.readFile(path.join(root, 'cells.sqlite'));
    assert.deepEqual(await check(plan), { valid: true });
    assert.deepEqual(await fs.readFile(path.join(root, 'source.xlsx')), Buffer.from(bytes));
    assert.deepEqual(await fs.readFile(path.join(root, 'cells.sqlite')), before);
  }
});
test('native metadata proof refuses exact graph/index/ancestry/closure/offset mutants atomically', async (t) => {
  const bytes = createWorkbook(),
    { root, check } = await owner(t, bytes),
    plan = buildSpillMetadata(bytes, groups),
    before = await fs.readFile(path.join(root, 'cells.sqlite'));
  const replacements = (part, from, to) => (mutant) => {
    assert(mutant.replacements[part].includes(from));
    mutant.replacements[part] = mutant.replacements[part].replace(from, to);
  };
  const mutants = [
    replacements('xl/metadata.xml', '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}', '{wrong-uri}'),
    replacements('xl/metadata.xml', '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}', '{wrong-uri}'),
    replacements('xl/metadata.xml', '</metadata>', '<futureMetadata name="OTHER" count="0"/></metadata>'),
    replacements('xl/_rels/workbook.xml.rels', 'relationships/sheetMetadata', 'relationships/worksheet'),
    replacements('[Content_Types].xml', 'application/vnd.ms-excel.rdrichvalue+xml', 'application/x-wrong'),
    replacements('xl/richData/rdrichvaluestructure.xml', 'n="errorType" t="i"', 'n="errorType" t="s"'),
    replacements('xl/richData/rdRichValueTypes.xml', 'value="1"/>', 'value="1"><extra/></flag>'),
    (mutant) => {
      mutant.masters[`${sheet}\0A1`].cm = 0;
    },
    (mutant) => {
      mutant.masters[`${sheet}\0A90`].vm = 0;
    },
    (mutant) => {
      mutant.masters[`${sheet}\0A90`].intendedRange = 'A90:D93';
    },
  ];
  for (const mutate of mutants) {
    const mutant = structuredClone(plan);
    mutate(mutant);
    await assert.rejects(check(mutant));
    assert.deepEqual(await fs.readFile(path.join(root, 'cells.sqlite')), before);
  }
});
