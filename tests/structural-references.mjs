// Inserting or deleting rows and columns moves the formulas, names and references that point at them.
// The engine does this itself (native/src/package_edit/structural_parts.rs and
// rewrite_admission.rs); these cases go through the real indexed session and
// check the saved file with an independent reader (openpyxl).
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
import { pythonOnPath } from '../scripts/audit/python.mjs';

pythonOnPath();
const describe = new URL('../scripts/audit/describe-workbook.py', import.meta.url).pathname;
const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-structural-references-'));
after(async () => {
  closeNativeWorkbookWorker();
  await fs.rm(root, { recursive: true, force: true });
});

let run = 0;
/** A workbook with sheet "Data" (and "Other Sheet"), its cells and defined names. */
async function workbook(cells, names = {}, other = {}) {
  const file = path.join(root, `source-${++run}.xlsx`);
  const script = `
import json, sys
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
spec = json.loads(sys.argv[2])
book = Workbook()
data = book.active
data.title = "Data"
for ref, value in spec["cells"].items():
    data[ref] = value
other = book.create_sheet("Other Sheet")
for ref, value in spec["other"].items():
    other[ref] = value
for name, text in spec["names"].items():
    book.defined_names[name] = DefinedName(name, attr_text=text)
book.save(sys.argv[1])
`;
  execFileSync('python3', ['-c', script, file, JSON.stringify({ cells, names, other })]);
  return file;
}
async function structural(file, command) {
  const directory = path.join(root, `run-${++run}`);
  await fs.mkdir(directory);
  const book = new IndexedWorkbook(
    native,
    async (source, edits, out) => {
      const result = await editCalculatedWorkbookResult(await fs.readFile(source), edits);
      await fs.writeFile(out, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(file, { storageDirectory: path.join(directory, 'index') });
    const applied = await book.apply('0', [{ sheet: 'Data', ...command }]);
    const saved = path.join(directory, 'saved.xlsx');
    await fs.copyFile((await book.export(applied.view.storage.state)).path, saved);
    const read = (sheet) => JSON.parse(execFileSync('python3', [describe, saved, sheet], { encoding: 'utf8' }));
    return { data: read('Data'), other: read('Other Sheet') };
  } finally {
    await book.close();
  }
}
const insertRows = (file, index, count = 1) => structural(file, { operation: 'insertRow', index, count });

test('names keep their range, a single-cell name moves, and references from other sheets follow', async () => {
  const file = await workbook(
    { A1: 2, B1: '=$A1+A$1', A2: 3, B2: '=$A2+A$1', B3: '=SUM(B1:B2)' },
    { Region: 'Data!$A$1:$A$2', Scalar: 'Data!$A$3' },
    { A1: '=Data!B3' }
  );
  const { data, other } = await insertRows(file, 2);
  // A row inserted just below a named range does not join it, as in Excel.
  assert.deepEqual(data.names, ['Region=Data!$A$1:$A$2', 'Scalar=Data!$A$4']);
  assert.equal(data.formulas.B2, '=$A2+A$1', 'a row above the insertion is untouched');
  assert.equal(other.formulas.A1, '=Data!B4', 'another sheet follows the moved cell');
});

test('a total just above inserted rows expands; a name around a total keeps its range', async () => {
  const file = await workbook(
    { A1: 2, A2: 3, B3: '=SUM(A1:A2)+MAX(A1:A2)*2' },
    { Compound: 'SUM(Data!$A$1:$A$2)+MAX(Data!$A$1:$A$2)', Scalar: 'Data!$A$3' }
  );
  const { data } = await insertRows(file, 2, 2);
  assert.equal(data.formulas.B5, '=SUM(A1:A4)+MAX(A1:A4)*2');
  assert.deepEqual(data.names, ['Compound=SUM(Data!$A$1:$A$2)+MAX(Data!$A$1:$A$2)', 'Scalar=Data!$A$5']);
});

test('a formula the engine cannot adjust stops the edit instead of being left wrong', async () => {
  for (const cells of [{ B3: '=INDIRECT("A2")' }, { B3: "='Data:Other Sheet'!A1" }])
    await assert.rejects(insertRows(await workbook(cells), 2), (error) => typeof error.code === 'string');
  // A workbook-wide name with no sheet could mean any sheet.
  await assert.rejects(insertRows(await workbook({}, { Global: '$A$1:$A$2' }), 2), (error) => typeof error.code === 'string');
});

test('columns: mixed and cross-sheet references move; a deleted column shrinks ranges or becomes #REF!', async () => {
  const file = await workbook(
    { A1: 2, B1: 3, C1: '=SUM($A$1:B1)', D1: '=B$1' },
    { Region: 'Data!$A$1:$B$1' },
    { A1: '=Data!C1' }
  );
  const inserted = await structural(file, { operation: 'insertColumn', index: 1 });
  assert.equal(inserted.data.formulas.D1, '=SUM($A$1:C1)');
  assert.equal(inserted.other.formulas.A1, '=Data!D1');
  assert.deepEqual(inserted.data.names, ['Region=Data!$A$1:$C$1']);
  const deleted = await structural(file, { operation: 'deleteColumn', index: 1 });
  assert.equal(deleted.data.formulas.B1, '=SUM($A$1:A1)');
  assert.equal(deleted.data.formulas.C1, '=#REF!');
  assert.equal(deleted.other.formulas.A1, '=Data!B1');
});
