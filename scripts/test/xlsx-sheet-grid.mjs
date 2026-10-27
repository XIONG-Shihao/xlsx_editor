import { fileURLToPath } from 'node:url';
import { mkdir as ensureDirectory } from 'node:fs/promises';
/** Isolated browser acceptance for worksheet tabs and the real renderer worker. */
import assert from 'node:assert/strict';
import { readFile, writeFile, mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import path from 'node:path';
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
let browser;
try {
  browser = await chromium.launch({
    executablePath: process.env.XLSX_EDITOR_BROWSER_BIN ?? '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 2200, height: 900 } }),
    errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.route('**/*', (route) => (route.request().url().startsWith(url) ? route.continue() : route.abort()));
  await page.goto(url);
  await page.waitForFunction(() => globalThis.__xlsxDifferential);
  const seeded = editWorkbookResult(createWorkbook(), [{ sheet: 'Sheet1', cell: 'A1', value: 'Original' }]);
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), readWorkbook(seeded.bytes));
  await page.getByLabel('Sheet1 A1', { exact: true }).waitFor();
  const displayFixture = editWorkbookResult(createWorkbook(), [
    { sheet: 'Sheet1', cell: 'A1', value: 125 },
    { sheet: 'Sheet1', cell: 'B1', formula: 'A1*2' },
    { sheet: 'Sheet1', cell: 'C1', value: 1234.5, numberFormat: 'currency' },
    { sheet: 'Sheet1', cell: 'D1', value: '00123' },
    { sheet: 'Sheet1', cell: 'E1', value: 0.125, numberFormat: 'percent' },
    { sheet: 'Sheet1', cell: 'F1', value: 45292, numberFormat: 'date' },
  ]);
  const displayView = readWorkbook(displayFixture.bytes);
  displayView.sheets[0].widths = [
    { min: 1, max: 1, width: 23, hidden: false },
    { min: 2, max: 2, width: 31, hidden: false },
  ];
  displayView.sheets[0].rowHeights = [{ row: 0, height: 24 }];
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), displayView);
  const formulaCell = page.getByLabel('Sheet1 B1', { exact: true });
  await page.waitForFunction(() => document.querySelector('input[aria-label="Sheet1 B1"]')?.value === '250');
  await formulaCell.click();
  assert.equal(await page.getByLabel('Formula or value', { exact: true }).inputValue(), '=A1*2');
  await formulaCell.press('F2');
  assert.equal(await formulaCell.inputValue(), '=A1*2');
  await formulaCell.press('Escape');
  assert.equal(await formulaCell.inputValue(), '250');
  for (const [ref, formatted] of [
    ['C1', '$1,234.50'],
    ['D1', '00123'],
    ['E1', '12.50%'],
    ['F1', '2024-01-01'],
  ]) {
    const cell = page.getByLabel(`Sheet1 ${ref}`, { exact: true });
    await cell.click();
    assert.equal(await cell.inputValue(), formatted);
  }
  await page.getByLabel('Sheet1 A1', { exact: true }).click();
  assert.equal(
    await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().frame),
    0,
    'Selecting formatted cells must not commit display text'
  );
  await page.locator('.sheet-ribbon-cells summary').filter({ hasText: 'Format' }).click();
  const width = page.getByLabel('Column width', { exact: true }),
    height = page.getByLabel('Row height', { exact: true });
  assert.equal(await width.inputValue(), '23');
  assert.equal(await height.inputValue(), '24');
  await width.focus();
  await width.press('Tab');
  await height.press('Tab');
  assert.equal(
    await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().frame),
    0,
    'Untouched dimensions must not create history'
  );
  await width.fill('29');
  await width.press('Escape');
  await width.press('Tab');
  assert.equal(await width.inputValue(), '23');
  await width.fill('');
  await width.press('Tab');
  assert.equal(await width.inputValue(), '23');
  assert.equal(
    await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().frame),
    0,
    'Canceled and invalid sizes must leave the workbook unchanged'
  );
  await width.fill('27');
  await width.press('Enter');
  await page.waitForFunction(() => globalThis.__xlsxDifferential.snapshot().frame === 1);
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.getByLabel('Sheet1 B1', { exact: true }).click();
  await page.locator('.sheet-ribbon-cells summary').filter({ hasText: 'Format' }).click();
  assert.equal(await width.inputValue(), '31', 'Dimensions must follow the selected column');
  const beforeClipboard = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().frame);
  await width.evaluate((element) => {
    const data = new DataTransfer();
    data.setData('text/plain', '99');
    element.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }));
  });
  assert.equal(
    await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().frame),
    beforeClipboard,
    'Pasting into toolbar input must not mutate the selected cell'
  );
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const original = window.requestAnimationFrame,
      callbacks = [];
    window.requestAnimationFrame = (callback) => {
      callbacks.push(callback);
      return -callbacks.length;
    };
    globalThis.__flushFocusFrames = () => {
      window.requestAnimationFrame = original;
      for (const callback of callbacks) callback(performance.now());
    };
    globalThis.__focusFrames = callbacks;
    globalThis.__xlsxDifferential.restoreFocus({ sheet: 'Sheet1', cell: 'A1' });
  });
  await page.waitForFunction(() => globalThis.__focusFrames.length > 0);
  await page.getByLabel('Go to cell', { exact: true }).fill('B20');
  await page.evaluate(() => globalThis.__flushFocusFrames());
  assert.equal(
    await page.getByLabel('Go to cell', { exact: true }).evaluate((element) => element === document.activeElement),
    true,
    'Delayed history focus must not interrupt typing in the name box'
  );
  assert.equal(await page.getByLabel('Go to cell', { exact: true }).inputValue(), 'B20');
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), readWorkbook(seeded.bytes));
  await page.getByRole('button', { name: 'Add worksheet' }).click();
  await page.getByLabel('New worksheet name').fill('Notes');
  await page.getByRole('button', { name: 'Apply' }).click();
  await page.getByRole('button', { name: 'Notes', exact: true }).waitFor();
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await page.getByRole('button', { name: 'Duplicate', exact: true }).click();
  await page.getByLabel('Duplicate worksheet as').fill('Copy');
  await page.getByRole('button', { name: 'Apply' }).click();
  await page.getByRole('button', { name: 'Copy', exact: true }).waitFor();
  await page.getByRole('button', { name: 'Rename', exact: true }).click();
  await page.getByLabel('Rename worksheet').fill('Archive');
  await page.getByRole('button', { name: 'Apply' }).click();
  await page.getByRole('button', { name: 'Archive', exact: true }).waitFor();
  const anchor = await page.getByLabel('Archive A1', { exact: true }).boundingBox();
  const corner = await page.getByLabel('Archive B2', { exact: true }).boundingBox();
  assert(anchor && corner);
  await page.mouse.move(anchor.x + anchor.width / 2, anchor.y + anchor.height / 2);
  await page.mouse.down();
  await page.mouse.move(corner.x + corner.width / 2, corner.y + corner.height / 2);
  await page.mouse.up();
  await page.locator('.sheet-ribbon-alignment summary').click();
  await page
    .locator('.sheet-ribbon-alignment .sheet-ribbon-menu > div')
    .getByRole('button', { name: 'Merge & Center' })
    .click();
  await page.waitForFunction(() =>
    globalThis.__xlsxDifferential
      .snapshot()
      .view.sheets.find((sheet) => sheet.name === 'Archive')
      ?.merges.includes('A1:B2')
  );
  await page.locator('.sheet-ribbon-alignment summary').click();
  await page
    .locator('.sheet-ribbon-alignment .sheet-ribbon-menu > div')
    .getByRole('button', { name: 'Unmerge cells' })
    .click();
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.merges.length ===
      0
  );
  for (const [ref, value] of [
    ['B1', 'Qty'],
    ['A2', 'Paper'],
    ['B2', '2'],
  ]) {
    await page.getByLabel(`Archive ${ref}`, { exact: true }).fill(value);
    await page.getByLabel(`Archive ${ref}`, { exact: true }).press('Tab');
    await page.waitForFunction(
      () => document.querySelector('.spreadsheet-editor')?.getAttribute('aria-busy') === 'false'
    );
  }
  const header = await page.getByLabel('Archive A1', { exact: true }).boundingBox();
  const data = await page.getByLabel('Archive B2', { exact: true }).boundingBox();
  assert(header && data);
  await page.mouse.move(header.x + header.width / 2, header.y + header.height / 2);
  await page.mouse.down();
  await page.mouse.move(data.x + data.width / 2, data.y + data.height / 2);
  await page.mouse.up();
  await page.locator('.sheet-ribbon-styles summary').filter({ hasText: 'Format as Table' }).click();
  await page.getByLabel('Table name').fill('Items');
  await page.getByRole('button', { name: 'Create table' }).click();
  await page.waitForFunction(() =>
    globalThis.__xlsxDifferential.snapshot().view.tables?.some((table) => table.name === 'Items')
  );
  for (const [ref, value] of [
    ['D1', '1'],
    ['D2', '3'],
  ]) {
    await page.getByLabel(`Archive ${ref}`, { exact: true }).fill(value);
    await page.getByLabel(`Archive ${ref}`, { exact: true }).press('Tab');
    await page.waitForFunction(
      () => document.querySelector('.spreadsheet-editor')?.getAttribute('aria-busy') === 'false'
    );
  }
  const seriesStart = await page.getByLabel('Archive D1', { exact: true }).boundingBox();
  const seriesEnd = await page.getByLabel('Archive D5', { exact: true }).boundingBox();
  assert(seriesStart && seriesEnd);
  await page.mouse.move(seriesStart.x + seriesStart.width / 2, seriesStart.y + seriesStart.height / 2);
  await page.mouse.down();
  await page.mouse.move(seriesEnd.x + seriesEnd.width / 2, seriesEnd.y + seriesEnd.height / 2);
  await page.mouse.up();
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Fill' }).click();
  await page.getByRole('button', { name: 'Fill series' }).click();
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.D5
        ?.value === 9
  );
  await page.getByLabel('Archive D1', { exact: true }).click();
  const handle = await page
    .getByLabel('Archive D1', { exact: true })
    .locator('xpath=../..')
    .getByRole('button', { name: 'Drag to fill selected range' })
    .boundingBox();
  const fillTarget = await page.getByLabel('Archive E1', { exact: true }).boundingBox();
  assert(handle && fillTarget);
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
  await page.mouse.down();
  await page.mouse.move(fillTarget.x + fillTarget.width / 2, fillTarget.y + fillTarget.height / 2);
  await page.mouse.up();
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.E1
        ?.value === 1
  );
  const beforeComposition = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().edits.length);
  const composed = page.getByLabel('Archive E2', { exact: true });
  await composed.click();
  await composed.dispatchEvent('compositionstart', { data: '日' });
  await composed.fill('日本語');
  await composed.dispatchEvent('keydown', { key: 'Enter', isComposing: true });
  assert.equal(
    await page.evaluate(() => globalThis.__xlsxDifferential.snapshot().edits.length),
    beforeComposition,
    'IME Enter must not publish an unfinished cell draft'
  );
  await composed.dispatchEvent('compositionend', { data: '日本語' });
  await composed.press('Tab');
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.E2
        ?.value === '日本語'
  );
  const selectionStart = await page.getByLabel('Archive A1', { exact: true }).boundingBox();
  const selectionEnd = await page.getByLabel('Archive E2', { exact: true }).boundingBox();
  assert(selectionStart && selectionEnd);
  await page.mouse.move(selectionStart.x + selectionStart.width / 2, selectionStart.y + selectionStart.height / 2);
  await page.mouse.down();
  await page.mouse.move(selectionEnd.x + selectionEnd.width / 2, selectionEnd.y + selectionEnd.height / 2);
  await page.mouse.up();
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Find & Select' }).click();
  await page.getByRole('button', { name: 'Select blanks' }).click();
  await page.getByRole('status').filter({ hasText: '2 matching cells selected' }).waitFor();
  assert.equal(
    await page.getByLabel('Archive C1', { exact: true }).locator('xpath=../..').getAttribute('aria-selected'),
    'true'
  );
  assert.equal(
    await page.getByLabel('Archive A1', { exact: true }).locator('xpath=../..').getAttribute('aria-selected'),
    'false'
  );
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await page.waitForFunction(() => {
    const cells = globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells;
    return cells?.C1?.bold === true && cells?.C2?.bold === true;
  });
  assert.equal(await page.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'), 'true');
  assert.notEqual(
    await page.evaluate(
      () =>
        globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.A1?.bold
    ),
    true
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.waitForFunction(() => {
    const cells = globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells;
    return cells?.C1?.bold !== true && cells?.C2?.bold !== true;
  });
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.waitForFunction(() => {
    const cells = globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells;
    return cells?.C1?.bold === true && cells?.C2?.bold === true;
  });
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Find & Select' }).click();
  await page.getByRole('button', { name: 'Clear special selection' }).click();
  await page.getByLabel('Archive F1', { exact: true }).fill('=D1+1');
  await page.getByLabel('Archive F1', { exact: true }).press('Tab');
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.F1
        ?.formula === 'D1+1'
  );
  const formulaStart = await page.getByLabel('Archive D1', { exact: true }).boundingBox();
  const formulaEnd = await page.getByLabel('Archive F2', { exact: true }).boundingBox();
  assert(formulaStart && formulaEnd);
  await page.mouse.move(formulaStart.x + formulaStart.width / 2, formulaStart.y + formulaStart.height / 2);
  await page.mouse.down();
  await page.mouse.move(formulaEnd.x + formulaEnd.width / 2, formulaEnd.y + formulaEnd.height / 2);
  await page.mouse.up();
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Find & Select' }).click();
  await page.getByRole('button', { name: 'Select formulas' }).click();
  await page.getByRole('status').filter({ hasText: '1 matching cells selected' }).waitFor();
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Clear' }).click();
  await page.getByRole('button', { name: 'Clear contents' }).click();
  await page.waitForFunction(
    () =>
      !globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.F1?.formula
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.waitForFunction(
    () =>
      globalThis.__xlsxDifferential.snapshot().view.sheets.find((sheet) => sheet.name === 'Archive')?.cells.F1
        ?.formula === 'D1+1'
  );
  await page.locator('.sheet-ribbon-editing summary').filter({ hasText: 'Find & Select' }).click();
  await page.getByRole('button', { name: 'Clear special selection' }).click();
  await page.getByLabel('Archive B2', { exact: true }).click();
  await page.locator('.sheet-ribbon-cells summary').filter({ hasText: 'Format' }).click();
  await page.getByRole('button', { name: 'Freeze here' }).click();
  await page.waitForFunction(() => {
    const freeze = globalThis.__xlsxDifferential
      .snapshot()
      .view.sheets.find((sheet) => sheet.name === 'Archive')?.freeze;
    return freeze?.rows === 1 && freeze?.columns === 1;
  });
  await page.getByLabel('Go to cell').fill('E100');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.getByLabel('Archive E100', { exact: true }).waitFor();
  await page.waitForFunction(() => document.querySelector('.sheet-grid-scroll')?.scrollTop > 0);
  const scrollBox = await page.locator('.sheet-grid-scroll').boundingBox();
  const frozenBox = await page.getByLabel('Archive A1', { exact: true }).boundingBox();
  const targetBox = await page.getByLabel('Archive E100', { exact: true }).boundingBox();
  const scrollState = await page.locator('.sheet-grid-scroll').evaluate((element) => ({
    top: element.scrollTop,
    left: element.scrollLeft,
    height: element.scrollHeight,
    width: element.scrollWidth,
  }));
  assert(scrollBox && frozenBox && targetBox);
  assert(
    frozenBox.x >= scrollBox.x && frozenBox.y >= scrollBox.y && frozenBox.y < scrollBox.y + scrollBox.height,
    'Frozen A1 must stay inside the visible grid after distant navigation'
  );
  assert(
    targetBox.x >= scrollBox.x &&
      targetBox.x < scrollBox.x + scrollBox.width &&
      targetBox.y >= scrollBox.y &&
      targetBox.y < scrollBox.y + scrollBox.height,
    `Distant E100 target must be visible in the grid: ${JSON.stringify({ scrollBox, frozenBox, targetBox, scrollState })}`
  );
  const snapshot = await page.evaluate(() => globalThis.__xlsxDifferential.snapshot());
  assert.deepEqual(
    snapshot.view.sheets.map((sheet) => sheet.name),
    ['Sheet1', 'Notes', 'Archive']
  );
  assert.equal(snapshot.view.sheets[2].cells.A1.value, 'Original');
  assert.equal(snapshot.edits.length, 21);
  assert.deepEqual(snapshot.view.sheets[2].tableRanges, [{ name: 'Items', ref: 'A1:B2' }]);
  assert.deepEqual(
    ['D3', 'D4', 'D5'].map((ref) => snapshot.view.sheets[2].cells[ref].value),
    [5, 7, 9]
  );
  assert.equal(snapshot.view.sheets[2].cells.E1.value, 1);
  assert.equal(snapshot.view.sheets[2].cells.E2.value, '日本語');
  const navigationView = readWorkbook(createWorkbook());
  navigationView.sheets[0].rows = 40;
  navigationView.sheets[0].columns = 10;
  navigationView.sheets[0].hiddenRows = [1];
  navigationView.sheets[0].widths = [{ min: 2, max: 2, width: 12, hidden: true }];
  navigationView.sheets[0].merges = ['C1:D2'];
  navigationView.sheets[0].freeze = { rows: 1, columns: 1 };
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), navigationView);
  const active = async (label) =>
    page.waitForFunction((expected) => document.activeElement?.getAttribute('aria-label') === expected, label);
  await page.getByLabel('Sheet1 A1', { exact: true }).click();
  await page.getByLabel('Sheet1 A1', { exact: true }).press('ArrowDown');
  await active('Sheet1 A3');
  await page.getByLabel('Sheet1 A3', { exact: true }).press('ArrowRight');
  await active('Sheet1 C3');
  await page.getByLabel('Sheet1 C3', { exact: true }).press('ArrowUp');
  await active('Sheet1 C1');
  await page.getByLabel('Sheet1 C1', { exact: true }).press('ArrowRight');
  await active('Sheet1 E1');
  await page.getByLabel('Sheet1 E1', { exact: true }).press('ArrowLeft');
  await active('Sheet1 C1');
  assert.match(await page.locator('.sheet-a11y-status').textContent(), /Sheet1 C1\. Cell selected/);
  await page.getByLabel('Sheet1 C1', { exact: true }).press('F2');
  assert.match(await page.getByLabel('Sheet1 C1', { exact: true }).getAttribute('aria-description'), /Editing cell/);
  await page.getByLabel('Sheet1 C1', { exact: true }).press('Escape');
  const mixedView = readWorkbook(createWorkbook());
  mixedView.sheets[0].cells.A1 = {
    value: 'Left',
    bold: true,
    italic: true,
    wrapText: true,
    horizontalAlign: 'left',
    verticalAlign: 'top',
    format: 10,
  };
  mixedView.sheets[0].cells.B1 = {
    value: 'Right',
    bold: false,
    italic: false,
    wrapText: false,
    horizontalAlign: 'right',
    verticalAlign: 'bottom',
    format: 0,
  };
  mixedView.sheets[0].columns = 2;
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), mixedView);
  await page.setViewportSize({ width: 2200, height: 900 });
  await page.getByLabel('Sheet1 A1', { exact: true }).waitFor();
  await page.getByLabel('Sheet1 B1', { exact: true }).waitFor();
  const mixedStart = await page.getByLabel('Sheet1 A1', { exact: true }).boundingBox();
  const mixedEnd = await page.getByLabel('Sheet1 B1', { exact: true }).boundingBox();
  assert(mixedStart && mixedEnd);
  await page.mouse.move(mixedStart.x + mixedStart.width / 2, mixedStart.y + mixedStart.height / 2);
  await page.mouse.down();
  await page.mouse.move(mixedEnd.x + mixedEnd.width / 2, mixedEnd.y + mixedEnd.height / 2);
  await page.mouse.up();
  const fullRibbon = page.getByRole('toolbar', { name: 'Spreadsheet Home', exact: true });
  for (const name of ['Bold', 'Italic', 'Wrap Text', 'Align left', 'Align top'])
    assert.equal(
      await fullRibbon.getByRole('button', { name, exact: true }).getAttribute('aria-pressed'),
      'mixed',
      `${name} must report a mixed selection`
    );
  assert.equal(await fullRibbon.getByRole('combobox', { name: 'Number format' }).inputValue(), 'mixed');
  await page.setViewportSize({ width: 850, height: 680 });
  const compactRibbon = page.getByRole('toolbar', { name: 'Spreadsheet Home', exact: true });
  await compactRibbon.waitFor();
  assert.equal(
    await compactRibbon.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
    'mixed'
  );
  assert.equal(await compactRibbon.getByRole('combobox', { name: 'Number format' }).inputValue(), 'mixed');
  const spanningView = readWorkbook(createWorkbook());
  spanningView.sheets[0].rows = 120;
  spanningView.sheets[0].columns = 6;
  spanningView.sheets[0].cells.C50 = { value: 'Span' };
  spanningView.sheets[0].merges = ['C50:D100'];
  spanningView.sheets[0].freeze = { rows: 1, columns: 1 };
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), spanningView);
  await page.setViewportSize({ width: 2200, height: 900 });
  await page.getByLabel('Go to cell').fill('E90');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.getByLabel('Sheet1 E90', { exact: true }).waitFor();
  assert.equal(
    await page.getByLabel('Sheet1 C90', { exact: true }).count(),
    0,
    'a covered cell must not become an independent editable input when the merge anchor scrolls out'
  );
  await page.getByRole('cell', { name: 'Merged continuation of C50:D100' }).waitFor();
  await page.getByLabel('Sheet1 E90', { exact: true }).press('ArrowLeft');
  await active('Sheet1 C50');
  spanningView.sheets[0].columns = 56;
  spanningView.sheets[0].merges = ['C50:BC100'];
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), spanningView);
  await page.getByLabel('Go to cell').fill('BD90');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.getByLabel('Sheet1 BD90', { exact: true }).waitFor();
  await page.waitForFunction(() => {
    const scroll = document.querySelector('.sheet-grid-scroll')?.getBoundingClientRect();
    const target = document.querySelector('input[aria-label="Sheet1 BD90"]')?.getBoundingClientRect();
    return (
      scroll &&
      target &&
      target.x < scroll.right &&
      target.right > scroll.x &&
      target.y < scroll.bottom &&
      target.bottom > scroll.y
    );
  });
  assert.equal(
    await page.getByLabel('Sheet1 BB90', { exact: true }).count(),
    0,
    'a merge spanning both virtual axes must not expose a covered cell as an input'
  );
  const continuation = page.getByRole('cell', { name: 'Merged continuation of C50:BC100' });
  await continuation.waitFor();
  const continuationBox = await continuation.boundingBox();
  const gridBox = await page.locator('.sheet-grid-scroll').boundingBox();
  const distantBox = await page.getByLabel('Sheet1 BD90', { exact: true }).boundingBox();
  const offset = await page
    .locator('.sheet-grid-scroll')
    .evaluate((element) => ({ top: element.scrollTop, left: element.scrollLeft }));
  assert(
    continuationBox &&
      gridBox &&
      continuationBox.x < gridBox.x + gridBox.width &&
      continuationBox.x + continuationBox.width > gridBox.x &&
      continuationBox.y < gridBox.y + gridBox.height &&
      continuationBox.y + continuationBox.height > gridBox.y,
    `the merged continuation must intersect the visible viewport: ${JSON.stringify({ continuationBox, gridBox, distantBox, offset })}`
  );
  await page.waitForFunction(
    () => {
      const cell = document.querySelector('input[aria-label="Sheet1 BD90"]')?.closest('td')?.getBoundingClientRect();
      const header = [...document.querySelectorAll('.sheet-grid thead th')]
        .find((item) => item.textContent === 'BD')
        ?.getBoundingClientRect();
      return cell && header && Math.abs(cell.x - header.x) < 3;
    },
    undefined,
    { timeout: 5000 }
  );
  await continuation.click();
  assert.match(await page.locator('.sheet-a11y-status').textContent(), /Sheet1 C50\. Cell selected/);
  await page.getByLabel('Go to cell').fill('BD90');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.getByLabel('Sheet1 BD90', { exact: true }).waitFor();
  await page.getByLabel('Sheet1 BD90', { exact: true }).press('ArrowLeft');
  await active('Sheet1 C50');
  await page.getByLabel('Go to cell').fill('BB90');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await active('Sheet1 C50');
  spanningView.sheets[0].merges = ['A1:BC100'];
  spanningView.sheets[0].cells.A1 = { value: 'Frozen span' };
  delete spanningView.sheets[0].cells.C50;
  await page.evaluate((view) => globalThis.__xlsxDifferential.initialize(view), spanningView);
  await page.getByLabel('Go to cell').fill('BD90');
  await page.getByRole('button', { name: 'Go', exact: true }).click();
  await page.getByLabel('Sheet1 BD90', { exact: true }).waitFor();
  await page.waitForFunction(() => {
    const scroll = document.querySelector('.sheet-grid-scroll')?.getBoundingClientRect();
    const target = document.querySelector('input[aria-label="Sheet1 BD90"]')?.getBoundingClientRect();
    return (
      scroll &&
      target &&
      target.x < scroll.right &&
      target.right > scroll.x &&
      target.y < scroll.bottom &&
      target.bottom > scroll.y
    );
  });
  assert.equal(await page.getByLabel('Sheet1 A90', { exact: true }).count(), 0);
  assert.equal(await page.getByLabel('Sheet1 BB90', { exact: true }).count(), 0);
  assert(
    (await page.getByRole('cell', { name: 'Merged continuation of A1:BC100' }).count()) >= 2,
    'frozen and scrolled panes must render noneditable merge continuations'
  );
  await page.waitForFunction(
    () => {
      const cell = document.querySelector('input[aria-label="Sheet1 BD90"]')?.closest('td')?.getBoundingClientRect();
      const header = [...document.querySelectorAll('.sheet-grid thead th')]
        .find((item) => item.textContent === 'BD')
        ?.getBoundingClientRect();
      return cell && header && Math.abs(cell.x - header.x) < 3;
    },
    undefined,
    { timeout: 5000 }
  );
  assert.deepEqual(errors, []);
  console.log(
    'PASS XLSX worksheet, merge, table, numeric series, drag fill, special selections, composition commit, frozen distant navigation and hidden/merged keyboard movement through the real grid, worker and history'
  );
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
  await rm(temporary, { recursive: true, force: true });
}
