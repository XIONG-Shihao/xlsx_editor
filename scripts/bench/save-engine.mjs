import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../../core/native-runtime.mjs';
import { createWorkbook, editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
const rows = Number(process.argv[2] ?? 10000),
  formulas = process.argv[3] !== 'plain';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
let body = '';
for (let r = 1; r <= rows; r++) {
  body += `<row r="${r}">`;
  for (let c = 0; c < 7; c++) body += `<c r="${String.fromCharCode(65 + c)}${r}"><v>${r * 10 + c}</v></c>`;
  body += formulas ? `<c r="H${r}"><f>A${r}+B${r}</f><v>${r * 20 + 1}</v></c>` : `<c r="H${r}"><v>${r}</v></c>`;
  body += '</row>';
}
const parts = unzipSync(createWorkbook());
parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
const root = await fs.mkdtemp(path.join(os.tmpdir(), 'bench-save-')),
  source = path.join(root, 'source.xlsx');
await fs.writeFile(source, zipSync(parts));
const calls = [];
const timed = async (input) => {
  const t = performance.now();
  try {
    return await native(input);
  } finally {
    calls.push([input.operation, performance.now() - t]);
  }
};
const calc = async (input) => {
  const t = performance.now();
  try {
    return await calculateIndexedWorkbook(input, timed);
  } finally {
    calls.push(['JS-calculate', performance.now() - t]);
  }
};
const book = new IndexedWorkbook(
  timed,
  async (s, edits, out) => {
    const t = performance.now();
    const result = await editCalculatedWorkbookResult(await fs.readFile(s), edits);
    await fs.writeFile(out, result.bytes);
    calls.push(['JS-legacy-writer', performance.now() - t]);
    return { revision: result.view.revision, edits: result.edits };
  },
  calc,
  {}
);
const phase = async (name, fn) => {
  const mark = calls.length,
    t = performance.now();
  const value = await fn();
  const total = performance.now() - t;
  const by = {};
  for (const [op, ms] of calls.slice(mark)) by[op] = (by[op] ?? 0) + ms;
  console.log(
    name.padEnd(8),
    total.toFixed(0).padStart(7) + ' ms  ',
    Object.entries(by)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 7)
      .map(([k, v]) => k + ' ' + v.toFixed(0))
      .join(', ')
  );
  return value;
};
await phase('open', () => book.open(source));
const edited = await phase('edit', () => book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]));
const out = await phase('export', () => book.export(edited.view.storage.state));
const edited2 = await phase('edit2', () =>
  book.apply(edited.view.storage.state, [{ sheet: 'Sheet1', cell: 'A2', value: 6 }])
);
await phase('export2', () => book.export(edited2.view.storage.state));
console.log('size', (await fs.stat(out.path)).size);
await book.close();
closeNativeWorkbookWorker();
await fs.rm(root, { recursive: true, force: true });
