// A workbook whose formulas the editor cannot calculate opens in
// preserve-calculation mode. Changing how the sheet is laid out touches no
// formula and no stored result, so it is allowed there.
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { preservedFixture } from './support/preserved-calculation-fixture.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

after(() => closeNativeWorkbookWorker());
const cellsOf = (xml) => [...xml.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((match) => match[0]).join('\n');

for (const kind of ['shared', 'cache', 'iterative'])
  test(`column width, row height and frozen panes change on a ${kind} workbook without touching a cell`, async (t) => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserved-layout-')),
      file = path.join(directory, 'source.xlsx'),
      bytes = preservedFixture(kind);
    await fs.writeFile(file, bytes);
    const book = new IndexedWorkbook(
      native,
      () => assert.fail('No fallback writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
    t.after(async () => {
      await book.close();
      await fs.rm(directory, { recursive: true, force: true });
    });
    const view = await book.open(file);
    assert.equal(view.calculationPolicy.preserveSource, true);
    let state = '0';
    for (const command of [
      { sheet: 'Sheet1', operation: 'resizeColumn', index: 0, width: 24 },
      { sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 30 },
      { sheet: 'Sheet1', operation: 'freeze', rows: 1, columns: 0 },
    ])
      state = (await book.apply(state, [command])).view.storage.state;
    const before = unzipSync(bytes),
      saved = unzipSync(await fs.readFile((await book.export(state)).path)),
      sheet = strFromU8(saved['xl/worksheets/sheet1.xml']);
    assert.equal(
      cellsOf(sheet),
      cellsOf(strFromU8(before['xl/worksheets/sheet1.xml'])),
      'every cell is byte-identical'
    );
    assert.match(sheet, /<col\b[^>]*\bwidth="24"/);
    assert.match(sheet, /<row\b(?=[^>]*\br="1")[^>]*\bht="30"/);
    assert.match(sheet, /<pane\b[^>]*\bySplit="1"[^>]*\bstate="frozen"/);
    for (const part of Object.keys(before))
      if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(part))
        assert.deepEqual(new Uint8Array(saved[part]), new Uint8Array(before[part]), part);
    // Merging empty cells changes no formula either.
    const merged = await book.apply(state, [{ sheet: 'Sheet1', operation: 'mergeCells', range: 'F5:G5' }]),
      shown = strFromU8(
        unzipSync(await fs.readFile((await book.export(merged.view.storage.state)).path))['xl/worksheets/sheet1.xml']
      );
    assert.match(shown, /<mergeCell ref="F5:G5"/);
    assert.equal(cellsOf(shown), cellsOf(strFromU8(before['xl/worksheets/sheet1.xml'])), 'cells unchanged');
    // Formulas and structure stay protected.
    await assert.rejects(book.apply(state, [{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]), {
      code: /^(formula-protected|package-protected|spill-protected)$/,
    });
    await assert.rejects(book.apply(state, [{ sheet: 'Sheet1', cell: 'Z9', formula: 'A1+1' }]), {
      code: /^(formula-protected|package-protected|spill-protected)$/,
    });
  });
