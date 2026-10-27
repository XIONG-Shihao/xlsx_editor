import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { SpreadsheetDocument } from '../../adapters/document-session.mjs';
import { createWorkbook } from '../../core/workbook-package.mjs';
import { closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
const rows = Number(process.argv[2] ?? 10000),
  ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
let body = '';
for (let r = 1; r <= rows; r++) {
  body += `<row r="${r}">`;
  for (let c = 0; c < 7; c++) body += `<c r="${String.fromCharCode(65 + c)}${r}"><v>${r * 10 + c}</v></c>`;
  body += `<c r="H${r}"><f>A${r}+B${r}</f><v>${r * 20 + 1}</v></c></row>`;
}
const parts = unzipSync(createWorkbook());
parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
const root = await mkdtemp(path.join(os.tmpdir(), 'bench-doc-')),
  filename = path.join(root, 'book.xlsx');
await writeFile(filename, zipSync(parts));
const owner = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery') });
const time = async (name, fn) => {
  const t = performance.now();
  const v = await fn();
  console.log(name.padEnd(6), (performance.now() - t).toFixed(0).padStart(7), 'ms');
  return v;
};
await time('open', () => owner.open(filename));
await time('edit', () =>
  owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: 'Sheet1', cell: 'A1', value: 5 }],
  })
);
await time('save', () => owner.request({ operation: 'save' }));
await time('edit2', () =>
  owner.request({
    operation: 'apply',
    state: owner.status().state,
    commands: [{ sheet: 'Sheet1', cell: 'A2', value: 6 }],
  })
);
await time('save2', () => owner.request({ operation: 'save' }));
await owner.close();
closeNativeWorkbookWorker();
await rm(root, { recursive: true, force: true });
process.exit(0);
