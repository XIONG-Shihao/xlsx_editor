/** Negative tests ensure the independent Freeze oracle cannot bless other edits. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
const run = promisify(execFile),
  checker = fileURLToPath(new URL('./support/check-excel-view.py', import.meta.url));
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function packageBytes(split, mutate = (x) => x) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${NS}"><sheetViews><sheetView workbookViewId="0"><pane xSplit="${split}" ySplit="${split}" topLeftCell="${split === 1 ? 'B2' : 'C3'}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><cols><col min="4" max="4" hidden="1"/></cols><sheetData><row r="1"><c r="C1"><f>A1+B1</f><v>30</v></c></row><row r="4" hidden="1"><c r="C4"><v>444</v></c></row><row r="8"><c r="C8"><v>17</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="E2:F3"/></mergeCells></worksheet>`
  );
  parts['custom/view-opaque.bin'] = new Uint8Array([1, 2, 3, 255]);
  mutate(parts);
  return zipSync(parts, { mtime: new Date(2000, 0, 1) });
}
async function check(t, mutation) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-view-oracle-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const before = path.join(directory, 'before.xlsx'),
    saved = path.join(directory, 'saved.xlsx'),
    undo = path.join(directory, 'undo.xlsx');
  await fs.writeFile(before, packageBytes(1));
  await fs.writeFile(saved, packageBytes(2, mutation));
  await fs.writeFile(undo, packageBytes(1));
  return run(process.env.XLSX_DESKTOP_PYTHON ?? 'python3', [checker, before, saved, undo], { maxBuffer: 1024 * 1024 });
}
test('independent Freeze oracle accepts only the declared saved and Undo pane delta', async (t) => {
  const result = JSON.parse((await check(t)).stdout);
  assert.equal(result.passed, true);
  assert.deepEqual(
    result.outputs.map((r) => r.freeze),
    [
      { rows: 2, columns: 2 },
      { rows: 1, columns: 1 },
    ]
  );
});
for (const [name, mutation] of [
  [
    'cached formula changed to a text-typed scalar',
    (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<c r="C1">', '<c r="C1" t="str">')
      );
    },
  ],
  [
    'unrelated hidden cell content changed',
    (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<v>444</v>', '<v>445</v>')
      );
    },
  ],
  [
    'view preferences leaked into saved sheet metadata',
    (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace('workbookViewId="0"', 'workbookViewId="0" zoomScale="125"')
      );
    },
  ],
  [
    'opaque member altered',
    (parts) => {
      parts['custom/view-opaque.bin'] = new Uint8Array([1, 2, 4, 255]);
    },
  ],
  [
    'unrelated workbook relationship changed',
    (parts) => {
      parts['xl/_rels/workbook.xml.rels'] = strToU8(
        strFromU8(parts['xl/_rels/workbook.xml.rels']).replace('styles.xml', 'other.xml')
      );
    },
  ],
  [
    'undeclared calculation mode',
    (parts) => {
      parts['xl/workbook.xml'] = strToU8(
        strFromU8(parts['xl/workbook.xml']).replace('</workbook>', '<calcPr calcMode="manual"/></workbook>')
      );
    },
  ],
  [
    'valid-looking but unrelated calcPr auto/calcId addition',
    (parts) => {
      parts['xl/workbook.xml'] = strToU8(
        strFromU8(parts['xl/workbook.xml']).replace(
          '</workbook>',
          '<calcPr calcMode="auto" calcId="191029"/></workbook>'
        )
      );
    },
  ],
  [
    'duplicate identical pane',
    (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
          '</sheetView>',
          '<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/></sheetView>'
        )
      );
    },
  ],
  [
    'duplicate conflicting pane',
    (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
          '</sheetView>',
          '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView>'
        )
      );
    },
  ],
])
  test('independent Freeze oracle refuses ' + name, async (t) => {
    await assert.rejects(
      check(t, mutation),
      (error) => error.code === 1 && String(error.stderr).includes('AssertionError')
    );
  });
