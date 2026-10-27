import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { createHash } from 'node:crypto';
import * as physical from '../core/worksheet-geometry.mjs';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { indexedPrintGeometry } from '../core/indexed-print.mjs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/source-geometry-checks-'));
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/viewGeometry.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/worksheetFontMetrics.ts'))};export {SpreadsheetGrid} from ${JSON.stringify(path.join(component, 'ui/SpreadsheetGrid.tsx'))};export {useAcceptedPrintPreview,isPrintPreviewRetired} from ${JSON.stringify(path.join(component, 'ui/useAcceptedPrintPreview.ts'))};export {configureXlsxHost} from ${JSON.stringify(path.join(component, 'ui/host.tsx'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: bundle,
  external: ['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'],
  logLevel: 'silent',
});
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  HTMLDetailsElement: dom.window.HTMLDetailsElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
Object.defineProperties(window.HTMLElement.prototype, {
  clientWidth: {
    get() {
      return 600;
    },
  },
  clientHeight: {
    get() {
      return 300;
    },
  },
});
globalThis.ResizeObserver = class {
  constructor(fn) {
    this.fn = fn;
  }
  observe() {
    this.fn([{ contentRect: { width: 600, height: 300 } }]);
  }
  disconnect() {}
};
globalThis.requestAnimationFrame = () => 1;
globalThis.cancelAnimationFrame = () => {};
const fontsReady = defer();
Object.defineProperty(document, 'fonts', {
  configurable: true,
  value: { ready: fontsReady.promise, addEventListener() {}, removeEventListener() {} },
});
const measurements = [];
function canvas() {
  return {
    font: '',
    measureText(text) {
      measurements.push({ font: this.font, text });
      if (text.length > 1)
        return { width: this.font.includes('Missing') ? (this.font.endsWith('monospace') ? 100 : 80) : 90 };
      return {
        width: this.font.includes('Arial')
          ? 9.6
          : this.font.includes('sans-serif') && !this.font.includes('Calibri')
            ? 8.2
            : 7.3,
      };
    },
  };
}
globalThis.CanvasRenderingContext2D = class {};
window.HTMLCanvasElement.prototype.getContext = () => canvas();
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
ui.configureXlsxHost({
  api: { copyText: async () => {}, readClipboardText: async () => '' },
  Icon: () => React.createElement('svg'),
  useI18n: () => ({ locale: 'en' }),
});
function defer() {
  let resolve, reject;
  const promise = new Promise((a, b) => {
    resolve = a;
    reject = b;
  });
  return { promise, resolve, reject };
}
function check(name, fn) {
  test(name, async () => {
    await fn();
  });
}
const font = { fontFamily: 'Calibri', fontSize: 11, bold: false, italic: false };
const profile = (width) =>
  physical.validateGeometryProfile({
    schema: 1,
    font,
    maximumDigitWidth: width,
    source: 'local-font',
    availability: 'unverified',
  });
const sheet = (patch = {}) => ({
  name: 'Sheet1',
  state: 'visible',
  path: 'xl/worksheets/sheet1.xml',
  rows: 10,
  columns: 6,
  cells: { A1: { value: 1 }, B1: { value: 2 } },
  merges: [],
  widths: [],
  hiddenRows: [],
  visibleRows: [],
  rowHeights: [],
  normalFont: { ...font },
  sheetFormat: { rowsHiddenByDefault: false },
  freeze: { rows: 0, columns: 0 },
  ...patch,
});
const close = (a, b) => assert(Math.abs(a - b) < 1e-8, a + ' != ' + b);
after(async () => {
  const sources = [];
  for (const relative of [
    'core/worksheet-geometry.mjs',
    'core/worksheet-geometry.d.mts',
    'ui/viewGeometry.ts',
    'ui/worksheetFontMetrics.ts',
    'ui/SpreadsheetGrid.tsx',
    'ui/styles/01.css',
    'ui/useAcceptedPrintPreview.ts',
    'ui/WorkbookPrint.tsx',
    'tests/source-geometry-ui.mjs',
  ]) {
    const p = path.join(component, relative),
      bytes = await fs.readFile(p);
    sources.push({ relative, path: p, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length });
  }
  dom.window.close();
});

check('OOXML defaults are 15pt rows and base8 plus padding; exact published raw-width example gives61px', () => {
  const source = sheet(),
    before = JSON.stringify(source),
    model = physical.createWorksheetDimensions(source, profile(7));
  assert.equal(model.defaultRowHeight, 15);
  assert.equal(model.defaultColumnWidth, 8.7109375);
  assert.equal(model.defaultColumnPixels, 61);
  assert.equal(physical.xmlColumnWidthPixels(8.7109375, 7), 61);
  assert.equal(physical.xmlColumnWidthPixels(9.140625, 7), 64);
  assert.equal(physical.rowHeightPoints(source, 0), 15);
  assert.equal(ui.rowHeight(source, 0), 20);
  assert.equal(ui.columnWidth(source, 0, 100, { geometryProfile: profile(7) }), 61);
  assert.equal(
    physical.createWorksheetDimensions(
      sheet({ sheetFormat: { rowsHiddenByDefault: false, baseColumnWidth: 255 } }),
      profile(1)
    ).defaultColumnPixels,
    260
  );
  assert.equal(JSON.stringify(source), before);
});
check('defaultColWidth is already encoded and measured digit widths change source axes without a60px minimum', () => {
  const source = sheet({
    sheetFormat: { rowsHiddenByDefault: false, defaultColumnWidth: 8.43, baseColumnWidth: 30 },
    widths: [
      { min: 1, max: 1, width: 0.5, hidden: false },
      { min: 2, max: 2, width: 0, hidden: false },
      { min: 3, max: 3, width: 20, hidden: true },
    ],
  });
  assert.deepEqual([...physical.worksheetColumnOffsets(source, 4, profile(7))], [0, 3, 3, 3, 62]);
  assert.deepEqual([...physical.worksheetColumnOffsets(source, 4, profile(10))], [0, 5, 5, 5, 89]);
  assert.equal(physical.columnOffsetPixels(source, 4, profile(7)), 62);
  assert.equal(physical.columnWidthPixels(source, 3, profile(10)), 84);
});
check(
  'hidden-by-default uses explicit visible false overrides; explicit hidden/zero heights win and sparse distances agree',
  () => {
    const source = sheet({
      sheetFormat: { rowsHiddenByDefault: true, defaultRowHeight: 22.5 },
      visibleRows: [0, 2, 1000],
      hiddenRows: [2],
      rowHeights: [{ row: 1000, height: 30 }],
    });
    const model = physical.createWorksheetDimensions(source, profile(7));
    assert.equal(model.rowHeightPt(0), 22.5);
    assert.equal(model.rowHeightPt(1), 0);
    assert.equal(model.rowHeightPt(2), 0);
    assert.equal(model.rowHeightPt(1000), 30);
    close(physical.rowOffsetPixels(source, 1001), 70);
    const offsets = physical.worksheetRowOffsets(source, 1001, profile(7));
    close(offsets[1001], 70);
    assert.equal(offsets[1000], 30);
    const zero = sheet({
      sheetFormat: { rowsHiddenByDefault: false, defaultRowHeight: 0 },
      rowHeights: [{ row: 2, height: 15 }],
    });
    assert.deepEqual([...physical.worksheetRowOffsets(zero, 4)], [0, 0, 0, 20, 20]);
  }
);
check(
  'source prefixes, hidden runs, Freeze, zoom, pointer and distant reveal share the same physical coordinates',
  () => {
    const source = sheet({
      rowHeights: [{ row: 1, height: 30 }],
      hiddenRows: [2],
      widths: [
        { min: 1, max: 1, width: 0.5, hidden: false },
        { min: 2, max: 2, width: 10, hidden: true },
      ],
      freeze: { rows: 1, columns: 1 },
    });
    const g = ui.createViewGeometry(source, { rows: 10, columns: 6, zoom: 125, geometryProfile: profile(7) });
    assert.deepEqual([...g.rowOffsets].slice(0, 5), [0, 25, 75, 75, 100]);
    assert.deepEqual([...g.columnOffsets].slice(0, 4), [0, 3.75, 3.75, 80]);
    assert.equal(g.frozenHeight, 25);
    assert.equal(g.frozenWidth, 3.75);
    assert.deepEqual(ui.hitTestCell(g, 62, 42, { scrollLeft: 200, scrollTop: 300 }), { row: 0, col: 0 });
    assert.deepEqual(ui.hitTestCell(g, 70, 80, { scrollLeft: 0, scrollTop: 0 }), { row: 1, col: 2 });
    assert.deepEqual(
      ui.virtualWindow(g.rowOffsets, 0, 110, { headerExtent: 40, frozenCount: 1, overscan: 0 }).indices,
      [1]
    );
    const far = ui.createViewGeometry(source, { rows: 1000, columns: 200, zoom: 100, geometryProfile: profile(7) }),
      scroll = ui.revealCellScroll(far, { width: 600, height: 300, scrollLeft: 0, scrollTop: 0 }, 999, 199);
    assert(scroll.scrollTop > 19000);
    assert(scroll.scrollLeft > 11000);
    assert.equal(
      ui.zoomToSelection(far, { first: { row: 0, col: 0 }, last: { row: 4, col: 4 } }, { width: 600, height: 300 }),
      Math.min(400, Math.floor((600 / (48 + 3 + 183)) * 100), Math.floor((300 / (32 + 20 + 80)) * 100))
    );
  }
);
check(
  'profile validator binds exact Normal font and closes provenance/ranges/fields rather than trusting font check',
  () => {
    const p = profile(7);
    assert(Object.isFrozen(p) && Object.isFrozen(p.font));
    assert.equal(p.availability, 'unverified');
    for (const bad of [
      { ...p, extra: true },
      { ...p, maximumDigitWidth: 0 },
      { ...p, maximumDigitWidth: 7.5 },
      { ...p, maximumDigitWidth: 4097 },
      { ...p, font: { ...font, fontSize: 0 } },
      { ...p, font: { ...font, extra: 1 } },
      { ...p, font: { ...font, fontFamily: 'bad\nfont' } },
      { ...p, source: 'reference-fallback', availability: 'confirmed' },
    ])
      assert.throws(() => physical.validateGeometryProfile(bad));
    assert.throws(() => physical.validateGeometryProfile(p, { ...font, fontFamily: 'Arial' }));
    assert.equal(physical.referenceGeometryProfile(font).source, 'reference-fallback');
  }
);
check(
  'local canvas measures all ten source-font digits unzoomed; substitution and missing canvas remain explicit',
  () => {
    const start = measurements.length,
      local = ui.measureGeometryProfile({ ...font, fontSize: 12, bold: true, italic: true }, canvas());
    assert.equal(local.maximumDigitWidth, 7);
    assert.equal(local.source, 'local-font');
    assert.equal(local.availability, 'unverified');
    const samples = measurements.slice(start);
    assert(samples.some((x) => x.font.startsWith('italic bold 16px')));
    assert.deepEqual(
      samples.filter((x) => x.text.length === 1).map((x) => x.text),
      [...'0123456789']
    );
    const substitute = ui.measureGeometryProfile({ ...font, fontFamily: 'Missing' }, canvas());
    assert.equal(substitute.source, 'substituted-font');
    assert.equal(substitute.availability, 'substituted');
    assert.equal(substitute.renderedFontFamily, 'sans-serif');
    assert.equal(substitute.maximumDigitWidth, 8);
    const saved = globalThis.CanvasRenderingContext2D;
    delete globalThis.CanvasRenderingContext2D;
    try {
      assert.equal(ui.measureGeometryProfile(font).source, 'reference-fallback');
    } finally {
      globalThis.CanvasRenderingContext2D = saved;
    }
    assert.equal(
      ui.measureGeometryProfile(font, { font: '', measureText: () => ({ width: NaN }) }).source,
      'reference-fallback'
    );
  }
);
let fontResult;
function FontProbe({ source }) {
  fontResult = ui.useWorksheetGeometryProfile(source);
  return React.createElement(
    'output',
    { 'data-font': fontResult.font.fontFamily },
    String(fontResult.maximumDigitWidth)
  );
}
check(
  'font replacement masks previous metrics immediately and retired font-ready work cannot adopt another font',
  async () => {
    const root = createRoot(document.getElementById('root'));
    try {
      await act(async () => root.render(React.createElement(FontProbe, { source: sheet() })));
      assert.equal(fontResult.maximumDigitWidth, 7);
      await act(async () =>
        root.render(React.createElement(FontProbe, { source: sheet({ normalFont: { ...font, fontFamily: 'Arial' } }) }))
      );
      assert.equal(fontResult.maximumDigitWidth, 10);
      assert.equal(document.querySelector('output').dataset.font, 'Arial');
      await act(async () => fontsReady.resolve());
      assert.equal(fontResult.font.fontFamily, 'Arial');
      assert.equal(fontResult.maximumDigitWidth, 10);
    } finally {
      await act(async () => root.unmount());
    }
  }
);
async function mount(extra = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    reads = [];
  let props = {
    sheets: [sheet()],
    documentId: 'doc',
    ownerIdentity: 'owner',
    storageKey: 'head',
    readOnly: false,
    onEdit: (x) => writes.push(x),
    onBatch: (x) => writes.push(x),
    onCommandBatch: async (x) => writes.push(x),
    ...extra,
  };
  const render = async (patch) => {
    props = { ...props, ...patch };
    await act(async () => root.render(React.createElement(ui.SpreadsheetGrid, props)));
  };
  await render();
  return { writes, reads, render, close: async () => act(async () => root.unmount()) };
}
async function click(n) {
  assert(n);
  await act(async () => n.click());
}
const tab = (name) => [...document.querySelectorAll('[role="tab"]')].find((n) => n.textContent === name);
check(
  'actual Grid uses source Normal font and narrow columns; source metadata-only updates invalidate memo and readonly stays readonly',
  async () => {
    const source = sheet({
        widths: [{ min: 1, max: 1, width: 0.5, hidden: false }],
        normalFont: { ...font, fontFamily: 'Arial', fontSize: 12, bold: true, italic: true },
      }),
      grid = await mount({ sheets: [source], readOnly: true });
    try {
      const table = document.querySelector('.sheet-grid');
      assert.equal(table.dataset.maximumDigitWidth, '10');
      assert.equal(table.style.fontFamily, '"Arial", sans-serif');
      assert.equal(table.style.fontSize, '12pt');
      assert.equal(table.style.fontWeight, '700');
      assert.equal(table.style.fontStyle, 'italic');
      assert.equal(document.querySelector('[aria-label="Bold"]').getAttribute('aria-pressed'), 'true');
      assert.equal(document.querySelector('[aria-label="Italic"]').getAttribute('aria-pressed'), 'true');
      assert.equal(document.querySelector('.sheet-grid tbody tr').style.height, '20px');
      assert.equal(document.querySelector('.sheet-grid colgroup col:nth-child(2)').style.width, '5px');
      assert(document.querySelector('[aria-label="Sheet1 A1"]').readOnly);
      await grid.render({
        sheets: [
          { ...source, sheetFormat: { rowsHiddenByDefault: false, defaultRowHeight: 30 }, normalFont: { ...font } },
        ],
      });
      assert.equal(document.querySelector('.sheet-grid tbody tr').style.height, '40px');
      assert.equal(table.dataset.maximumDigitWidth, '7');
      assert.equal(document.querySelector('.sheet-grid colgroup col:nth-child(2)').style.width, '3px');
      assert.deepEqual(grid.writes, []);
    } finally {
      await grid.close();
    }
  }
);
check(
  'actual hidden-default Grid reads sparse visible bands, renders their source coordinates and skips hidden rows on keyboard movement',
  async () => {
    const source = sheet({
        rows: 1001,
        sheetFormat: { rowsHiddenByDefault: true },
        visibleRows: [0, 1000],
        cells: { A1: { value: 1 }, A1001: { value: 2 } },
      }),
      reads = [],
      grid = await mount({
        sheets: [source],
        readRange: async (name, range) => {
          reads.push({ name, range });
          return source.cells;
        },
      });
    try {
      assert(document.querySelector('[aria-label="Sheet1 A1"]'));
      assert(document.querySelector('[aria-label="Sheet1 A1001"]'));
      assert(reads.length <= 3, JSON.stringify(reads));
      assert(
        reads.every((x) => ['A1:AF32', 'A993:AF1024'].includes(x.range)),
        JSON.stringify(reads)
      );
      const cell = document.querySelector('[aria-label="Sheet1 A1"]');
      await act(async () =>
        cell.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true }))
      );
      assert.equal(document.querySelector('[aria-label="Go to cell"]').placeholder, 'A1001');
      assert.deepEqual(grid.writes, []);
    } finally {
      await grid.close();
    }
  }
);
check('roomy CSV layout stays explicit and does not inherit XLSX source geometry or font notices', async () => {
  const grid = await mount({ literal: true });
  try {
    assert.equal(document.querySelector('.sheet-grid tbody tr').style.height, '58px');
    assert.equal(document.querySelector('.sheet-grid colgroup col:nth-child(2)').style.width, '100px');
    assert.equal(document.querySelector('[aria-label="Worksheet font rendering"]'), null);
    assert.equal(document.querySelector('.sheet-grid').dataset.geometryProfile, 'literal');
  } finally {
    await grid.close();
  }
});
check('Home dimension fields display accepted source defaults instead of stale43.5pt/12.5 constants', async () => {
  const grid = await mount({
    sheets: [
      sheet({ sheetFormat: { rowsHiddenByDefault: false, defaultRowHeight: 22.5, defaultColumnWidth: 9.140625 } }),
    ],
  });
  try {
    const labels = [...document.querySelectorAll('.sheet-ribbon-menu label')],
      row = labels.find((n) => n.textContent === 'Row height'),
      column = labels.find((n) => n.textContent === 'Column width');
    assert.equal(row.querySelector('input').value, '22.5');
    assert.equal(column.querySelector('input').value, '9.140625');
    assert.deepEqual(grid.writes, []);
  } finally {
    await grid.close();
  }
});
const settings = {
  area: 'A1:B3',
  paper: 'Letter',
  orientation: 'portrait',
  marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
  scaling: { mode: 'scale', percent: 100 },
  repeatRows: null,
  repeatColumns: null,
  gridlines: false,
};
const page = {
  number: 1,
  rowIndices: [0],
  columnIndices: [0],
  cells: [{ ref: 'A1', text: '1', style: { fontSize: 11 }, bbox: [36, 36, 100, 30] }],
  objects: [],
};
const job = (id) => ({
  id,
  state: 'base',
  units: 'pt',
  pageSizePt: [612, 792],
  scale: 1,
  sourceFingerprint: 'a'.repeat(64),
  revision: 'source',
  sheet: 'Sheet1',
  settings,
  pageCount: 1,
  pages: [page],
});
check(
  'Grid preparation passes the captured metric profile and a Normal-font change retires the accepted physical page lease',
  async () => {
    const calls = [],
      releases = [],
      grid = await mount({
        storageKey: 'owner:base',
        printControls: {
          documentId: 'runtime',
          revision: 'base',
          definition: async () => ({ sheet: 'Sheet1', revision: null, settings, canEdit: true, protectedReasons: [] }),
          preview: async (...args) => {
            calls.push(args);
            return { ...job('font-job'), geometryProfile: args[3] };
          },
          page: async () => page,
          onBatch: async () => {
            throw Error('No writes');
          },
          publish: async () => {},
          release: async (id) => releases.push(id),
        },
      });
    try {
      await click(tab('View'));
      await click(document.querySelector('[data-ribbon-command="view.page-layout"]'));
      assert.equal(calls.length, 1);
      assert.equal(calls[0][3].maximumDigitWidth, 7);
      assert.equal(calls[0][3].font.fontFamily, 'Calibri');
      assert(document.querySelector('.sheet-page-layout iframe'));
      await grid.render({ sheets: [sheet({ normalFont: { ...font, fontFamily: 'Arial' } })] });
      assert.equal(document.querySelector('.sheet-page-layout iframe'), null);
      assert.deepEqual(releases, ['font-job']);
      assert.equal(calls[0][3].font.fontFamily, 'Calibri');
      assert.deepEqual(grid.writes, []);
    } finally {
      await grid.close();
    }
  }
);
let preview;
function PreviewProbe({ scope, callbacks }) {
  preview = ui.useAcceptedPrintPreview({ scope, callbacks });
  return React.createElement('output', null, preview.job?.id ?? 'empty');
}
check('geometry-key change masks held same-head prepare and releases only its returned old job', async () => {
  const root = createRoot(document.getElementById('root')),
    gate = defer(),
    released = [],
    scope = {
      documentId: 'runtime',
      grantDocumentId: 'grant',
      state: 'base',
      sheet: 'Sheet1',
      ownerKey: 'owner:base',
      activationEpoch: 0,
      active: true,
      geometryKey: JSON.stringify(profile(7)),
    },
    callbacks = { preview: () => gate.promise, page: async () => page, release: async (id) => released.push(id) };
  let pending;
  try {
    await act(async () => root.render(React.createElement(PreviewProbe, { scope, callbacks })));
    await act(async () => {
      pending = preview.prepare(settings, 'old');
      pending.catch(() => {});
    });
    await act(async () =>
      root.render(
        React.createElement(PreviewProbe, { scope: { ...scope, geometryKey: JSON.stringify(profile(10)) }, callbacks })
      )
    );
    assert.equal(preview.job, null);
    assert.equal(preview.preparing, false);
    await act(async () => gate.resolve(job('old-metric')));
    await assert.rejects(pending, ui.isPrintPreviewRetired);
    assert.deepEqual(released, ['old-metric']);
    assert.equal(preview.job, null);
  } finally {
    gate.resolve(job('cleanup'));
    await act(async () => root.unmount());
  }
});
check('accepted source XML produces identical Grid and read-only print axes under the same Normal-font profile', () => {
  const parts = unzipSync(createWorkbook()),
    source = strFromU8(parts['xl/worksheets/sheet1.xml']);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    source.replace(
      '<sheetData/>',
      '<sheetFormatPr defaultRowHeight="22.5" defaultColWidth="9.140625" zeroHeight="true"/><cols><col min="2" max="2" width="0.5"/><col min="3" max="3" width="10" hidden="1"/></cols><sheetData><row r="1" hidden="false"/><row r="3" hidden="0" ht="30"/><row r="4" hidden="1" ht="40"/></sheetData>'
    )
  );
  const bytes = zipSync(parts),
    before = bytes.slice(),
    view = readWorkbook(bytes),
    sourceSheet = view.sheets[0],
    bound = physical.referenceGeometryProfile(sourceSheet.normalFont);
  const printed = indexedPrintGeometry(bytes, view, sourceSheet.name, { ...settings, area: 'A1:D4' }, bound),
    screen = ui.createViewGeometry(sourceSheet, { rows: 4, columns: 4, geometryProfile: bound, zoom: 125 });
  assert.deepEqual(
    printed.rows.map((row) => [row.heightPt, row.hidden]),
    [
      [22.5, false],
      [0, true],
      [30, false],
      [0, true],
    ]
  );
  assert.deepEqual(
    printed.columns.map((col) => [col.widthPt, col.hidden]),
    [
      [48, false],
      [2.25, false],
      [0, true],
      [48, false],
    ]
  );
  for (const row of printed.rows)
    close(((row.heightPt * 96) / 72) * 1.25, screen.rowOffsets[row.index + 1] - screen.rowOffsets[row.index]);
  for (const col of printed.columns)
    close(((col.widthPt * 96) / 72) * 1.25, screen.columnOffsets[col.index + 1] - screen.columnOffsets[col.index]);
  assert.deepEqual(printed.geometryProfile, bound);
  assert.deepEqual(bytes, before);
});

check('an echoed accepted font profile must match the captured axes before any physical page adoption', async () => {
  const root = createRoot(document.getElementById('root')),
    released = [],
    scope = {
      documentId: 'runtime',
      grantDocumentId: 'grant',
      state: 'base',
      sheet: 'Sheet1',
      ownerKey: 'owner:base',
      activationEpoch: 0,
      active: true,
      geometryKey: JSON.stringify(profile(7)),
    };
  let response = { ...job('wrong-metrics'), geometryProfile: profile(10) };
  const callbacks = { preview: async () => response, page: async () => page, release: async (id) => released.push(id) };
  try {
    await act(async () => root.render(React.createElement(PreviewProbe, { scope, callbacks })));
    await act(async () => assert.rejects(preview.prepare(settings), /worksheet dimensions/));
    assert.equal(preview.job, null);
    assert.deepEqual(released, ['wrong-metrics']);
    response = job('missing-metrics');
    await act(async () => assert.rejects(preview.prepare(settings), /confirm the worksheet dimensions/));
    assert.equal(preview.job, null);
    assert.deepEqual(released, ['wrong-metrics', 'missing-metrics']);
    response = { ...job('right-metrics'), geometryProfile: profile(7) };
    await act(async () => preview.prepare(settings));
    assert.equal(preview.job.id, 'right-metrics');
  } finally {
    await act(async () => root.unmount());
  }
});

check('wrapped source text and its edit input honor explicit font overrides over the Normal font', async () => {
  const grid = await mount({
    sheets: [
      sheet({
        normalFont: { ...font, bold: true, italic: true },
        cells: {
          A1: {
            value: 'wrapped',
            wrapText: true,
            bold: false,
            italic: false,
            fontFamily: 'Times New Roman',
            fontSize: 10,
          },
        },
      }),
    ],
  });
  try {
    const input = document.querySelector('[aria-label="Sheet1 A1"]'),
      display = input.parentElement.querySelector('.sheet-wrapped-display');
    for (const node of [input, display]) {
      assert.equal(node.style.fontFamily, '"Times New Roman", sans-serif');
      assert.equal(node.style.fontSize, '10pt');
      assert.equal(node.style.fontWeight, '400');
      assert.equal(node.style.fontStyle, 'normal');
    }
    assert.equal(document.querySelector('[aria-label="Bold"]').getAttribute('aria-pressed'), 'false');
    assert.deepEqual(grid.writes, []);
  } finally {
    await grid.close();
  }
});
