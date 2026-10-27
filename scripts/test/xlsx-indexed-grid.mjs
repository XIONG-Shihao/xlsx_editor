import { fileURLToPath } from 'node:url';
import { mkdir as ensureDirectory } from 'node:fs/promises';
/** Isolated browser acceptance for worksheet tabs and the real renderer worker. */
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { build } from 'vite';
import { chromium } from '@playwright/test';
import { createWorkbook, editWorkbookResult, readWorkbook } from '@office-editors/xlsx-editor/core/workbook-package.mjs';

const root = fileURLToPath(new URL('../../', import.meta.url));
await ensureDirectory(path.join(root, '.build/runs/browser'), { recursive: true });
const temporary = await mkdtemp(path.join(root, '.build/runs/browser/xlsx-sheet-grid-'));
const input = path.join(temporary, 'index.html'),
  output = path.join(temporary, 'out');
await writeFile(
  input,
  `<!doctype html><style>html,body,#root{height:100%;margin:0}#root{display:flex;flex-direction:column}.spreadsheet-editor{height:100%;min-height:0}</style><div id="root"></div><script type="module" src="${path.join(root, 'scripts/test/xlsx-differential-entry.tsx')}"></script>`
);
await build({
  configFile: false,
  root: temporary,
  logLevel: 'warn',
  worker: { format: 'es' },
  build: { outDir: output, target: 'esnext', rollupOptions: { input } },
});
const server = createServer(async (request, response) => {
  try {
    const file = path.join(output, request.url === '/' ? 'index.html' : decodeURIComponent(request.url.split('?')[0]));
    response.setHeader(
      'content-type',
      file.endsWith('.js') ? 'application/javascript' : file.endsWith('.css') ? 'text/css' : 'text/html'
    );
    response.end(await readFile(file));
  } catch {
    response.statusCode = 404;
    response.end();
  }
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const url = `http://127.0.0.1:${server.address().port}`;
let browser, book;
try {
  const { IndexedWorkbook } = await import('@office-editors/xlsx-editor/core/indexed-session.mjs');
  const { nativeWorkbookOperationAsync: native } = await import('@office-editors/xlsx-editor/native');
  const { calculateIndexedWorkbook } = await import('@office-editors/xlsx-editor/core/indexed-calculation.mjs');
  const { editCalculatedWorkbookResult } = await import('@office-editors/xlsx-editor/engine');
  const { zipSync, unzipSync, strToU8 } = await import('fflate');
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      Array.from(
        { length: 4200 },
        (_, row) =>
          `<row r="${row + 1}"><c r="A${row + 1}"><v>${row + 1}</v></c><c r="B${row + 1}" t="inlineStr"><is><t>${row === 3999 ? 'offscreen needle' : row === 0 || row === 2999 ? 'duplicate offscreen' : 'row ' + row}</t></is></c></row>`
      ).join('') +
      '</sheetData></worksheet>'
  );
  const source = path.join(temporary, 'source.xlsx');
  await writeFile(source, zipSync(parts));
  const sourceHash = createHash('sha256')
    .update(await readFile(source))
    .digest('hex');
  book = new IndexedWorkbook(
    native,
    async (source, commands, output) => {
      const result = await editCalculatedWorkbookResult(await readFile(source), commands);
      await writeFile(output, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  const view = await book.open(source);
  view.sheets[0].conditionalRules = [{ rule: 'duplicateValues', range: 'B1:B4200', fillColor: 'FFF2CC' }];
  browser = await chromium.launch({
    executablePath: process.env.XLSX_EDITOR_BROWSER_BIN ?? '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } }),
    errors = [],
    reads = [],
    queries = [];
  page.on('pageerror', (error) => errors.push(error.message));
  let blockQuery = false,
    releaseQuery,
    blockFilter = false,
    releaseFilter;
  await page.exposeFunction('__indexedCall', async (request) => {
    if (request.operation === 'read') {
      reads.push(request.range);
      return book.read(request.state, request.sheet, request.range);
    }
    if (request.operation === 'apply') return book.apply(request.state, request.commands);
    if (request.operation === 'scan') return book.scan(request.state, request.sheet, request.after);
    if (request.operation === 'filterValues') {
      if (blockFilter) {
        blockFilter = false;
        await new Promise((resolve) => (releaseFilter = resolve));
      }
      return book.filterValues(
        request.state,
        request.sheet,
        request.range,
        request.column,
        request.query,
        request.offset
      );
    }
    if (request.operation === 'query') {
      queries.push(request);
      if (blockQuery) {
        blockQuery = false;
        await new Promise((resolve) => (releaseQuery = resolve));
      }
      return book.query(request.state, request.sheet, request.query, request.mode, request.after, {
        range: request.range,
        matchLimit: request.matchLimit,
      });
    }
    throw Error('Unknown test request');
  });
  await page.goto(url);
  await page.waitForFunction(() => globalThis.__xlsxDifferential);
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), view);
  const ready = () =>
    page.waitForFunction(() => document.querySelector('.spreadsheet-editor')?.getAttribute('aria-busy') === 'false');
  await ready();
  assert.equal(await page.getByLabel('Sheet1 A1', { exact: true }).inputValue(), '1');
  assert.equal(
    await page.getByLabel('Sheet1 B1', { exact: true }).evaluate((input) => input.closest('td').style.backgroundColor),
    'rgb(255, 242, 204)',
    'Duplicate highlighting must include a matching offscreen cell'
  );
  await page.getByLabel('Go to cell', { exact: true }).fill('A4000');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 A4000', { exact: true }).inputValue(), '4000');
  const target = page.getByLabel('Sheet1 A4000', { exact: true });
  await target.fill('8000');
  await target.press('Tab');
  await ready();
  assert.equal(
    (
      await book.read(
        (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.storage.state,
        'Sheet1',
        'A4000'
      )
    ).cells.A4000.value,
    8000
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 A4000', { exact: true }).inputValue(), '4000');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 A4000', { exact: true }).inputValue(), '8000');
  await page.getByLabel('Go to cell', { exact: true }).fill('A1');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await ready();
  await page.getByLabel('Sheet1 A1', { exact: true }).press('Control+ArrowDown');
  await ready();
  await page
    .waitForFunction(() => document.querySelector('input[aria-label="Sheet1 A4200"]')?.value === '4200')
    .catch(async (error) => {
      console.log(JSON.stringify({ reads, errors, body: await page.locator('body').innerText() }));
      throw error;
    });
  await page.keyboard.press('Control+f');
  await page.getByLabel('Find', { exact: true }).fill('offscreen needle');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await page.waitForFunction(
    () => document.querySelector('input[aria-label="Sheet1 B4000"]')?.value === 'offscreen needle'
  );
  await ready();
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await ready();
  const formatted = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.deepEqual(formatted.edits.at(-1), {
    type: 'formatRange',
    sheet: 'Sheet1',
    range: 'B4000',
    format: { bold: true },
  });
  assert.equal((await book.read(formatted.view.storage.state, 'Sheet1', 'B4000')).cells.B4000.bold, true);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(await page.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'), 'false');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await ready();
  assert.equal(await page.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'), 'true');
  await page.getByLabel('Search content', { exact: true }).selectOption('raw');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await ready();
  await page.waitForFunction(
    () => document.querySelector('input[aria-label="Sheet1 B4000"]')?.value === 'offscreen needle'
  );
  assert(queries.length > 0, 'Raw search must reach the native query path');
  await page.getByLabel('Sheet1 B4000', { exact: true }).click();
  await page.keyboard.press('Delete');
  await ready();
  const cleared = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.deepEqual(cleared.edits.at(-1), { type: 'clearRange', sheet: 'Sheet1', range: 'B4000' });
  assert.equal((await book.read(cleared.view.storage.state, 'Sheet1', 'B4000')).cells.B4000.value, null);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 B4000', { exact: true }).inputValue(), 'offscreen needle');
  await page.getByLabel('Find', { exact: true }).fill('row ');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await ready();
  await page.waitForFunction(() => document.querySelector('input[aria-label="Sheet1 B2"]')?.value === 'row 1');
  await page.getByRole('button', { name: 'More results', exact: true }).click();
  await ready();
  await page.waitForFunction(() => document.querySelector('input[aria-label="Sheet1 B1002"]')?.value === 'row 1001');
  await page.getByRole('button', { name: 'Previous results', exact: true }).click();
  await ready();
  await page.waitForFunction(() => document.querySelector('input[aria-label="Sheet1 B2"]')?.value === 'row 1');
  await page.getByRole('button', { name: 'More results', exact: true }).click();
  await ready();
  await page.waitForFunction(() => document.querySelector('input[aria-label="Sheet1 B1002"]')?.value === 'row 1001');
  const beforeCancel = (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).frame;
  blockQuery = true;
  await page.getByRole('button', { name: 'More results', exact: true }).click();
  await page.getByRole('button', { name: 'Cancel search', exact: true }).click();
  assert.equal(typeof releaseQuery, 'function');
  releaseQuery();
  await page.waitForFunction(() => document.querySelector('.sheet-find')?.textContent.includes('Replace this page'));
  assert.equal((await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).frame, beforeCancel);
  assert.equal(await page.getByLabel('Sheet1 B1002', { exact: true }).inputValue(), 'row 1001');
  await page.getByLabel('Replace with', { exact: true }).fill('changed ');
  await page.getByRole('button', { name: 'Replace this page', exact: true }).click();
  await ready();
  const replaced = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.equal(replaced.lastCommands.length, 1);
  assert.equal(replaced.lastCommands[0].cells.length, 1000);
  let replacedCells = (await book.read(replaced.view.storage.state, 'Sheet1', 'B1001:B2002')).cells;
  assert.equal(replacedCells.B1001.value, 'row 1000');
  assert.equal(replacedCells.B1002.value, 'changed 1001');
  assert.equal(replacedCells.B2001.value, 'changed 2000');
  assert.equal(replacedCells.B2002.value, 'row 2001');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 B1002', { exact: true }).inputValue(), 'row 1001');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 B1002', { exact: true }).inputValue(), 'changed 1001');
  await page.getByLabel('Find', { exact: true }).fill('changed ');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await ready();
  await page.waitForFunction(
    () => document.querySelector('input[aria-label="Sheet1 B1002"]')?.value === 'changed 1001'
  );
  await page.getByRole('button', { name: 'More results', exact: true }).click();
  await page.getByRole('button', { name: 'Replace all', exact: true }).waitFor();
  assert.equal(await page.getByRole('button', { name: 'More results', exact: true }).isDisabled(), true);
  assert.equal(
    await page.getByLabel('Sheet1 B1002', { exact: true }).inputValue(),
    'changed 1001',
    'An exactly-full terminal page remains visible after EOF look-ahead'
  );
  const exported = await book.export(
    (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.storage.state
  );
  const saved = JSON.parse(
    execFileSync(
      'python3',
      [
        '-c',
        `import sys,json,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 cells={c.attrib['r']:c for c in E.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(n+'c')}
 strings=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  strings=[''.join(si.itertext()) for si in E.fromstring(z.read('xl/sharedStrings.xml'))]
 def text(ref):
  c=cells[ref]
  return strings[int(c.findtext(n+'v'))] if c.attrib.get('t')=='s' else ''.join(c.find(n+'is').itertext()) if c.attrib.get('t')=='inlineStr' else c.findtext(n+'v')
 print(json.dumps({r:text(r) for r in ['B1001','B1002','B2001','B2002','B4000']}))`,
        exported.path,
      ],
      { encoding: 'utf8' }
    )
  );
  assert.deepEqual(saved, {
    B1001: 'row 1000',
    B1002: 'changed 1001',
    B2001: 'changed 2000',
    B2002: 'row 2001',
    B4000: 'offscreen needle',
  });
  assert.equal(
    createHash('sha256')
      .update(await readFile(source))
      .digest('hex'),
    sourceHash,
    'Paged replacement never publishes over the input source'
  );
  await page.getByRole('button', { name: 'Close find', exact: true }).click();
  await page.locator('.sheet-grid thead th').filter({ hasText: /^B$/ }).click();
  await ready();
  await page.locator('.sheet-ribbon-menu > summary').filter({ hasText: 'Sort & Filter' }).click();
  await page.getByRole('button', { name: 'Load filter values', exact: true }).click();
  await page.waitForFunction(
    () => document.querySelectorAll('.sheet-filter-values input[type="checkbox"]').length === 100
  );
  const firstChoices = await page
    .locator('.sheet-filter-values input[type="checkbox"]')
    .evaluateAll((inputs) => inputs.map((input) => input.getAttribute('data-filter-key')));
  await page.getByRole('button', { name: 'More values', exact: true }).click();
  await page.waitForFunction(
    (first) =>
      document.querySelector('.sheet-filter-values input[type="checkbox"]')?.getAttribute('data-filter-key') !== first,
    firstChoices[0]
  );
  const secondChoices = await page
    .locator('.sheet-filter-values input[type="checkbox"]')
    .evaluateAll((inputs) => inputs.map((input) => input.getAttribute('data-filter-key')));
  assert.equal(secondChoices.length, 100);
  assert(!secondChoices.some((key) => firstChoices.includes(key)));
  await page.getByRole('button', { name: 'Previous values', exact: true }).click();
  await page.waitForFunction(
    (first) =>
      document.querySelector('.sheet-filter-values input[type="checkbox"]')?.getAttribute('data-filter-key') === first,
    firstChoices[0]
  );
  blockFilter = true;
  await page.getByRole('button', { name: 'More values', exact: true }).click();
  await page.getByRole('button', { name: 'Cancel value loading', exact: true }).click();
  assert.equal(typeof releaseFilter, 'function');
  releaseFilter();
  assert.deepEqual(
    await page
      .locator('.sheet-filter-values input[type="checkbox"]')
      .evaluateAll((inputs) => inputs.map((input) => input.getAttribute('data-filter-key'))),
    firstChoices
  );
  await page.getByLabel('Search filter values', { exact: true }).fill('offscreen needle');
  await page.getByRole('button', { name: 'Load filter values', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Filter value "offscreen needle" (text)', exact: true }).check();
  await page.getByRole('button', { name: 'Apply selected values', exact: true }).click();
  await ready();
  const filtered = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.deepEqual(filtered.lastCommands, [
    { sheet: 'Sheet1', operation: 'filterRows', range: 'B1:B4200', column: 'B', values: ['offscreen needle'] },
  ]);
  assert.equal(filtered.view.sheets[0].hiddenRows.length, 4198);
  assert(!filtered.view.sheets[0].hiddenRows.includes(3999));
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(
    (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.sheets[0].hiddenRows.length,
    0
  );
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await ready();
  const filteredExport = await book.export(
    (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.storage.state
  );
  const hidden = JSON.parse(
    execFileSync(
      'python3',
      [
        '-c',
        `import sys,json,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z:
 n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
 sheet=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
 print(json.dumps({'hidden':[int(row.attrib['r']) for row in sheet.iter(n+'row') if row.attrib.get('hidden')=='1'],'filter':[(f.attrib) for f in sheet.iter(n+'filter')]}))`,
        filteredExport.path,
      ],
      { encoding: 'utf8' }
    )
  );
  assert.equal(hidden.hidden.length, 4198);
  assert(!hidden.hidden.includes(4000));
  assert.deepEqual(hidden.filter, [{ val: 'offscreen needle' }]);
  assert.equal(
    createHash('sha256')
      .update(await readFile(source))
      .digest('hex'),
    sourceHash
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  await page.getByLabel('Go to cell', { exact: true }).fill('A1');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await ready();
  await page.getByRole('button', { name: 'Names and tables', exact: true }).click();
  await page.getByLabel('Range name', { exact: true }).fill('TopAmount');
  await page.getByRole('button', { name: 'Create range name', exact: true }).click();
  await ready();
  await page.getByRole('button', { name: 'Go to name TopAmount', exact: true }).waitFor();
  let namedSnapshot = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.deepEqual(namedSnapshot.lastCommands, [
    {
      operation: 'defineName',
      sheet: 'Sheet1',
      name: 'TopAmount',
      range: 'A1',
      nameScope: 'workbook',
      expectedNameFormula: null,
    },
  ]);
  assert.deepEqual(namedSnapshot.view.namedRanges, [{ name: 'TopAmount', formula: "'Sheet1'!$A$1:$A$1" }]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal((await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.namedRanges.length, 0);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await ready();
  await page.getByRole('button', { name: 'Go to name TopAmount', exact: true }).click();
  await ready();
  assert.equal(await page.getByLabel('Sheet1 A1', { exact: true }).inputValue(), '1');
  const namedExport = await book.export(
    (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.storage.state
  );
  const savedNames = JSON.parse(
    execFileSync(
      'python3',
      [
        '-c',
        `import sys,json,zipfile,xml.etree.ElementTree as E
with zipfile.ZipFile(sys.argv[1]) as z:
 n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
 print(json.dumps([{'name':v.attrib['name'],'formula':v.text} for v in E.fromstring(z.read('xl/workbook.xml')).iter(n+'definedName')]))`,
        namedExport.path,
      ],
      { encoding: 'utf8' }
    )
  );
  assert.deepEqual(savedNames, [{ name: 'TopAmount', formula: "'Sheet1'!$A$1:$A$1" }]);
  await page.getByRole('button', { name: 'Remove name TopAmount', exact: true }).click();
  await ready();
  assert.equal((await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.namedRanges.length, 0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await ready();
  assert.equal(
    (await page.evaluate(() => globalThis.__xlsxDifferential.snapshot())).view.namedRanges[0].name,
    'TopAmount'
  );
  assert.equal(
    createHash('sha256')
      .update(await readFile(source))
      .digest('hex'),
    sourceHash
  );
  assert(
    queries.every(
      (request) => Number.isInteger(request.matchLimit) && request.matchLimit >= 1 && request.matchLimit <= 1000
    ),
    'All real grid searches must bound returned matches'
  );
  const snapshot = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.equal(Object.keys(snapshot.view.sheets[0].cells).length, 0, 'History remains metadata-only');
  assert(reads.length > 4);
  assert.deepEqual(errors, []);
  console.log(
    'PASS Rust-backed real grid: distant edit/undo/format/navigation; bounded Find/filter pages; named-range create/navigate/remove/undo; independent saved XML; cell-free history'
  );
} finally {
  await browser?.close();
  await book?.close();
  const { closeNativeWorkbookWorker } = await import('@office-editors/xlsx-editor/native');
  closeNativeWorkbookWorker();
  await new Promise((resolve) => server.close(resolve));
  await rm(temporary, { recursive: true, force: true });
}
