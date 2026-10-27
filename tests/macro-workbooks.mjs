import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const MACRO_PARTS = ['xl/vbaProject.bin', 'xl/vbaProjectSignature.bin', 'xl/_rels/vbaProject.bin.rels'];

// A small macro workbook: a macro part with its signature part, the macro-enabled main
// type, and a code name on the workbook and the sheet, as Excel writes them.
function macroPackage() {
  const parts = unzipSync(createWorkbook());
  const bytes = (seed, length) => Uint8Array.from({ length }, (_, i) => (i * 31 + seed) & 255);
  parts['xl/vbaProject.bin'] = bytes(7, 4096);
  parts['xl/vbaProjectSignature.bin'] = bytes(11, 512);
  parts['xl/_rels/vbaProject.bin.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>'
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml'])
      .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', MACRO_MAIN)
      .replace(
        '</Types>',
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
          '<Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>'
      )
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      `<Relationship Id="rId9" Type="${VBA}" Target="vbaProject.bin"/></Relationships>`
    )
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      /<workbook([^>]*)>/,
      '<workbook$1><workbookPr codeName="ThisWorkbook"/>'
    )
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<sheetData', '<sheetPr codeName="Sheet1"/><sheetData')
  );
  return parts;
}

async function withBook(t, run) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-macro-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'macros.xlsm'),
    parts = macroPackage();
  await fs.writeFile(source, zipSync(parts));
  const book = new IndexedWorkbook(native, () => assert.fail('No fallback'));
  try {
    await run({ book, source, directory, parts });
  } finally {
    await book.close();
  }
}
// The macro part, its signature, the main type and the relationship to the project.
function assertMacroKept(saved, original) {
  for (const name of MACRO_PARTS) assert.equal(hash(saved[name]), hash(original[name]), name + ' changed');
  const types = strFromU8(saved['[Content_Types].xml']);
  assert(types.includes(MACRO_MAIN), 'macro-enabled type lost');
  assert(types.includes('Extension="bin" ContentType="application/vnd.ms-office.vbaProject"'));
  assert(strFromU8(saved['xl/_rels/workbook.xml.rels']).includes(VBA), 'macro relationship lost');
  assert(strFromU8(saved['xl/workbook.xml']).includes('codeName="ThisWorkbook"'));
}

test('a macro workbook opens for editing and every kind of edit leaves the macro part, its signature and its type untouched', async (t) => {
  await withBook(t, async ({ book, source, parts }) => {
    const view = await book.open(source);
    assert.equal(view.readOnly, false);
    let state = view.storage.state;
    const apply = async (commands) => (state = (await book.apply(state, commands)).view.storage.state);
    await apply([
      { sheet: 'Sheet1', cell: 'A1', value: 'Name' },
      { sheet: 'Sheet1', cell: 'B1', value: 'Amount' },
      { sheet: 'Sheet1', cell: 'A2', value: 'pear' },
      { sheet: 'Sheet1', cell: 'B2', value: 3 },
      { sheet: 'Sheet1', cell: 'A3', value: 'apple' },
      { sheet: 'Sheet1', cell: 'B3', value: 9 },
    ]);
    await apply([{ type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true } }]);
    await apply([{ sheet: 'Sheet1', operation: 'sortRows', range: 'A2:B3', column: 'A', direction: 'asc' }]);
    await apply([{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]);
    await apply([{ sheet: 'Added', operation: 'addSheet' }]);
    const saved = unzipSync(await fs.readFile((await book.export(state)).path));
    assertMacroKept(saved, parts);
    assert.equal((await book.read(state, 'Sheet1', 'A3')).cells.A3.value, 'apple');
  });
});

test('a saved macro workbook reopens, takes more edits and saves again with the macro part still byte for byte the same', async (t) => {
  await withBook(t, async ({ book, source, directory, parts }) => {
    let view = await book.open(source);
    const first = await book.apply(view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 'one' }]);
    const copy = path.join(directory, 'saved.xlsm');
    await fs.copyFile((await book.export(first.view.storage.state)).path, copy);
    await book.close();
    const again = new IndexedWorkbook(native, () => assert.fail('No fallback'));
    try {
      view = await again.open(copy);
      assert.equal(view.readOnly, false);
      const second = await again.apply(view.storage.state, [{ sheet: 'Sheet1', cell: 'A2', value: 'two' }]);
      const saved = unzipSync(await fs.readFile((await again.export(second.view.storage.state)).path));
      assertMacroKept(saved, parts);
    } finally {
      await again.close();
    }
    assert.deepEqual(unzipSync(await fs.readFile(source))['xl/vbaProject.bin'], parts['xl/vbaProject.bin']);
  });
});

test('ActiveX controls and package signatures still keep a workbook read-only', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-macro-guard-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  for (const [label, extra] of [
    ['activex', { 'xl/activeX/activeX1.xml': strToU8('<ocx/>') }],
    ['signature', { '_xmlsignatures/sig1.xml': strToU8('<Signature/>') }],
  ]) {
    const source = path.join(directory, label + '.xlsm');
    await fs.writeFile(source, zipSync({ ...macroPackage(), ...extra }));
    const book = new IndexedWorkbook(native, () => assert.fail('No fallback'));
    try {
      assert.equal((await book.open(source)).readOnly, true, label);
    } finally {
      await book.close();
    }
  }
});

test('copying a sheet is refused with a code, because Excel would add a code module to the macro project', async (t) => {
  await withBook(t, async ({ book, source }) => {
    const view = await book.open(source);
    assert.equal(view.macroProject, true);
    await assert.rejects(
      book.apply(view.storage.state, [{ sheet: 'Sheet1', operation: 'duplicateSheet', newName: 'Copy' }]),
      { code: 'macro-protected' }
    );
  });
});
