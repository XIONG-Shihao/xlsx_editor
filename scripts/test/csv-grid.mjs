import { fileURLToPath } from 'node:url';
import { mkdir as ensureDirectory } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { build } from 'vite';
import { chromium } from '@playwright/test';
const root = fileURLToPath(new URL('../../', import.meta.url));
await ensureDirectory(path.join(root, '.build/runs/browser'), { recursive: true });
const temporary = await mkdtemp(path.join(root, '.build/runs/browser/csv-grid-')),
  output = path.join(temporary, 'out');
await writeFile(
  path.join(temporary, 'index.html'),
  `<style>html,body,#root{height:100%;margin:0}#root{display:flex;flex-direction:column}.spreadsheet-editor{height:100%;min-height:0}</style><div id="root"></div><script type="module" src="${path.join(root, 'scripts/test/csv-grid-entry.tsx')}"></script>`
);
await build({
  configFile: false,
  root: temporary,
  logLevel: 'warn',
  worker: { format: 'es' },
  build: { outDir: output, target: 'esnext', rollupOptions: { input: path.join(temporary, 'index.html') } },
});
const server = createServer(async (req, res) => {
  try {
    const file = path.join(output, req.url === '/' ? 'index.html' : decodeURIComponent(req.url.split('?')[0]));
    res.setHeader(
      'content-type',
      file.endsWith('.js') ? 'application/javascript' : file.endsWith('.css') ? 'text/css' : 'text/html'
    );
    res.end(await readFile(file));
  } catch {
    res.statusCode = 404;
    res.end();
  }
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const url = `http://127.0.0.1:${server.address().port}`;
const browser = await chromium.launch({
    executablePath: process.env.XLSX_EDITOR_BROWSER_BIN ?? '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    headless: true,
  }),
  errors = [];
try {
  const page = await browser.newPage();
  page.on('pageerror', (e) => errors.push(e.message));
  await page.route('**/*', (route) => (route.request().url().startsWith(url) ? route.continue() : route.abort()));
  await page.goto(url);
  await page.getByLabel('CSV A1', { exact: true }).waitFor();
  await mkdir(path.join(root, '.build/runs/browser/csv'), { recursive: true });
  await page.screenshot({ path: path.join(root, '.build/runs/browser/csv/editor-ribbon-wide.png') });
  await page.getByText('Sort & Filter', { exact: false }).click();
  await page.screenshot({ path: path.join(root, '.build/runs/browser/csv/editor-sort-menu.png') });
  await page.keyboard.press('Escape');
  await page.getByLabel('Go to cell', { exact: true }).fill('B130001');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.screenshot({ path: path.join(root, '.build/runs/browser/csv/editor-ribbon-jump.png') });
  assert.equal(await page.getByLabel('CSV B130001', { exact: true }).inputValue(), 'original-129999');
  await page.getByLabel('CSV B130001', { exact: true }).fill('末尾🙂,quoted');
  await page.screenshot({ path: path.join(root, '.build/runs/browser/csv/editor-ribbon-filled.png') });
  assert.equal(
    (await page.evaluate(() => globalThis.__csv())).edits.length,
    0,
    'draft text must not enter history before commit'
  );
  await page.getByLabel('CSV B130001', { exact: true }).press('Tab');
  assert.equal((await page.evaluate(() => globalThis.__csv())).tail, '末尾🙂,quoted');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  assert.equal(await page.getByLabel('CSV B130001', { exact: true }).inputValue(), 'original-129999');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  assert.equal(await page.getByLabel('CSV B130001', { exact: true }).inputValue(), '末尾🙂,quoted');
  const state = await page.evaluate(() => globalThis.__csv());
  assert.equal(state.rows, 130001);
  assert(state.serializedTail.endsWith('129999,"末尾🙂,quoted"\r\n'));
  assert.equal(state.edits.length, 1);
  assert.equal(state.error, '');
  await page.getByLabel('Go to cell', { exact: true }).fill('B65001');
  await page.getByLabel('Go to cell', { exact: true }).press('Enter');
  assert.equal(await page.getByLabel('CSV B65001', { exact: true }).inputValue(), 'original-64999');
  await page.getByLabel('CSV B65001', { exact: true }).press('Meta+ArrowDown');
  await page.waitForFunction(() => document.activeElement?.getAttribute('aria-label') === 'CSV B130001');
  await page.screenshot({ path: path.join(root, '.build/runs/browser/csv/editor-ribbon-distant.png') });
  await page.locator('.sheet-ribbon-menu > summary').filter({ hasText: 'Find & Select' }).click();
  await page.getByRole('button', { name: 'Find and replace', exact: true }).click();
  await page.getByLabel('Find', { exact: true }).fill('original-64999');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await page.getByRole('status').filter({ hasText: '1 / 1' }).waitFor();
  await page.getByRole('button', { name: 'Preview replacements' }).click();
  assert.match(await page.getByLabel('Replacement preview').textContent(), /CSV!B65001/);
  const worker = await page.evaluate(() => globalThis.__workerProbe());
  assert.equal(worker.rows, 50000);
  assert.equal(worker.originalTail, 49999);
  assert.equal(worker.firstFormula.value, 9);
  assert.equal(worker.dependent.value, 12);
  assert.equal(worker.historyResync.value, 10);
  assert.equal(worker.errorPreserved.dataType, 'error');
  assert.equal(worker.unchangedIdentity, true);
  assert.deepEqual(errors, []);
  const files = [
    'scripts/test/csv-grid-entry.tsx',
    'scripts/test/csv-grid.mjs',
    'ui/SpreadsheetGrid.tsx',
    'ui/csvModel.ts',
    'ui/workbook.worker.ts',
    'ui/workbookWorker.ts',
    'ui/sheetModel.ts',
    'core/calculation.mjs',
    'core/csv-workbook.mjs',
  ];
  const report = {
    date: new Date().toISOString(),
    passed: true,
    scope:
      'Headless actual grid, indexed CSV core and actual browser calculation worker; no native app or publication IPC',
    csv: state,
    worker,
    browser: browser.version(),
    sources: Object.fromEntries(
      await Promise.all(
        files.map(async (file) => [
          file,
          createHash('sha256')
            .update(await readFile(path.join(root, file)))
            .digest('hex'),
        ])
      )
    ),
  };
  await mkdir(path.join(root, '.build/runs/browser/csv'), { recursive: true });
  await writeFile(path.join(root, '.build/runs/browser/csv/csv-grid.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(
    'PASS CSV 130001-row grid direct navigation, exact Unicode edit, undo/redo and serialize; real browser worker 50000 literals, first formula, dependent edit, error preservation and history resync'
  );
} finally {
  await browser.close();
  await new Promise((resolve) => server.close(resolve));
  await rm(temporary, { recursive: true, force: true });
}
