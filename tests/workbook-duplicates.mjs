import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { auditPython } from '../scripts/audit/python.mjs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { WorkbookSession } from '../core/workbook-service.mjs';
import { duplicateConfiguration, duplicateCommand, validateDuplicatePreview } from '../core/duplicate-model.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import {
  dedupDesktopFixture,
  dedupSavedExpectation,
  dedupSheet,
  dedupRange,
  dedupColumns,
  dedupRemovedRows,
  dedupKeptRows,
} from './support/dedup-desktop-fixture.mjs';
const roots = [],
  clients = [];
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  str = (r, v) => `<c r="${r}" t="inlineStr"><is><t>${v}</t></is></c>`,
  num = (r, v) => `<c r="${r}"><v>${v}</v></c>`;
function bytes(body, manual = false, tail = '') {
  const p = unzipSync(createWorkbook());
  p['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetFormatPr defaultRowHeight="15"/><sheetData>${body}</sheetData>${tail}</worksheet>`
  );
  p['xl/workbook.xml'] = strToU8(
    strFromU8(p['xl/workbook.xml']).replace(
      '</workbook>',
      `<calcPr calcMode="${manual ? 'manual' : 'auto'}" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`
    )
  );
  p['custom/dedup-owned.bin'] = new Uint8Array([0, 33, 255]);
  return Buffer.from(zipSync(p));
}
const simple = () =>
  bytes(
    `<row r="1">${str('A1', 'key')}${str('B1', 'record')}</row><row r="2">${num('A2', '1')}${str('B2', 'first')}${str('C2', 'outside2')}</row><row r="3">${num('A3', '1.0')}${str('B3', 'duplicate')}${str('C3', 'outside3')}</row><row r="4">${num('A4', '2')}${str('B4', 'last')}${str('C4', 'outside4')}</row>`
  );
async function fixture(original) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-duplicates-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, original);
  return { root, source, recovery: path.join(root, 'recovery'), original };
}
async function nativeFixture(original) {
  const f = await fixture(original),
    client = createWorkbookNativeClient(nativeWorkbookBinary());
  clients.push(client);
  const call = (request) => client.call({ action: 'workbook', root: f.root, ...request }),
    index = await call({ operation: 'index' }),
    sheet = index.sheets[0].path;
  return { ...f, call, sheet, index };
}
const config = { range: 'A1:B4', hasHeaders: true, columns: ['A'] };
const directPreview = (f, options = {}) =>
  f.call({ operation: 'dedup-preview', version: 0, sheet: f.sheet, ...config, ...options });
const directApply = (f, fingerprint, options = {}) =>
  f.call({
    operation: 'dedup-apply',
    version: 0,
    sheet: f.sheet,
    source: 'source.xlsx',
    ...config,
    expectedFingerprint: fingerprint,
    ...options,
  });
async function owner(f, override) {
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  clients.push(client);
  const calls = [];
  const native = async (input) => {
    calls.push(structuredClone(input));
    return override ? override(input, client) : client.call(input);
  };
  const doc = new SpreadsheetDocument({ recoveryDirectory: f.recovery, native });
  await doc.open(f.source);
  return { doc, calls, native };
}
const preview = (doc, sheet = 'Sheet1', range = config.range, hasHeaders = true, columns = ['A'], offset = 0) =>
  doc.request({
    operation: 'duplicatePreview',
    documentId: doc.status().documentId,
    state: doc.status().state,
    sheet,
    range,
    hasHeaders,
    columns,
    offset,
  });
const apply = (doc, command) =>
  doc.request({
    operation: 'apply',
    documentId: doc.status().documentId,
    state: doc.status().state,
    commands: [command],
  });
const remove = (page) => ({
  operation: 'removeDuplicates',
  sheet: page.sheet,
  range: page.range,
  hasHeaders: page.hasHeaders,
  columns: page.columns,
  expectedFingerprint: page.fingerprint,
});
async function copy(doc, f, name) {
  const file = path.join(f.root, name + '.xlsx');
  await doc.request({ operation: 'copy', destination: file, expectedRevision: null });
  return file;
}
function check(name, fn) {
  test(name, async (t) => {
    const result = await fn(t);
  });
}
after(async () => {
  for (const client of clients) client.close();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
check('closed configuration and recorded receipt reject ambiguous fields and false authority', async () => {
  const c = duplicateConfiguration('Sheet1', config.range, true, ['A']);
  assert.equal(c.body, 1);
  for (const args of [
    ['A1', true, ['A']],
    ['A1:B4', true, []],
    ['A1:B4', true, ['A', 'A']],
    ['A1:B4', true, ['C']],
    ['A1:B4', true, ['a']],
    ['A1:XFD1048576', true, ['A']],
  ])
    assert.throws(() => duplicateConfiguration('Sheet1', ...args));
  const command = { operation: 'removeDuplicates', sheet: 'Sheet1', ...config, expectedFingerprint: 'a'.repeat(64) };
  assert.deepEqual(duplicateCommand(command), command);
  assert.throws(() => duplicateCommand({ ...command, rows: [3] }));
  return { closed: true };
});
check('native selected raw records compact once with own-row formulas and exact independent XML oracle', async () => {
  const f = await nativeFixture(dedupDesktopFixture()),
    page = await directPreview(f, { range: dedupRange, columns: dedupColumns });
  assert.equal(page.keptRows, 10);
  assert.equal(page.removedRows, 3);
  assert.deepEqual(page.samples, dedupRemovedRows);
  const result = await directApply(f, page.fingerprint, { range: dedupRange, columns: dedupColumns });
  const expected = path.join(f.root, 'expected.json');
  await fs.writeFile(
    expected,
    JSON.stringify({
      ...dedupSavedExpectation(),
      formulaRawValues: dedupKeptRows.map((row) => String((100 + row) * 10 + 10)),
    })
  );
  execFileSync(process.env.XLSX_DESKTOP_PYTHON ?? auditPython(), [
    'tests/support/check-dedup-desktop.py',
    f.source,
    path.join(f.root, result.file),
    expected,
  ]);
  assert.deepEqual(await fs.readFile(f.source), f.original);
  return { kept10Removed3: true, independentXML: true };
});
check('actual Document preview readonly footprint then one removal UndoRedo Save and cold', async () => {
  const f = await fixture(dedupDesktopFixture()),
    { doc, calls } = await owner(f);
  try {
    const before = doc.status(),
      read = await doc.request({ operation: 'read', state: before.state, sheet: dedupSheet, range: dedupRange });
    const root = calls.find((c) => c.operation === 'index').root,
      db = await fs.readFile(path.join(root, 'cells.sqlite')),
      page = (await preview(doc, dedupSheet, dedupRange, true, dedupColumns)).duplicatePreview;
    validateDuplicatePreview(
      page,
      {
        documentId: before.documentId,
        ownerId: before.view.storage.sessionId,
        state: before.state,
        revision: before.view.revision,
      },
      { sheet: dedupSheet, range: dedupRange, hasHeaders: true, columns: dedupColumns }
    );
    assert.equal(doc.status().state, before.state);
    assert.deepEqual(await fs.readFile(path.join(root, 'cells.sqlite')), db);
    assert.deepEqual(
      await doc.request({ operation: 'read', state: before.state, sheet: dedupSheet, range: dedupRange }),
      read
    );
    await apply(doc, remove(page));
    const removed = doc.status().state;
    assert.notEqual(removed, before.state);
    assert.equal(
      (await doc.request({ operation: 'read', state: removed, sheet: dedupSheet, range: 'D3:E3' })).cells.D3.value,
      'record-4'
    );
    await doc.request({ operation: 'undo' });
    assert.equal(doc.status().state, before.state);
    await doc.request({ operation: 'redo' });
    assert.equal(doc.status().state, removed);
    const file = await copy(doc, f, 'removed'),
      expected = path.join(f.root, 'expected.json');
    await fs.writeFile(expected, JSON.stringify(dedupSavedExpectation()));
    execFileSync(process.env.XLSX_DESKTOP_PYTHON ?? auditPython(), [
      'tests/support/check-dedup-desktop.py',
      f.source,
      file,
      expected,
    ]);
    assert.deepEqual(await fs.readFile(f.source), f.original);
    await doc.request({ operation: 'save' });
    await doc.close();
    const cold = await owner(f);
    try {
      assert.equal(cold.doc.status().dirty, false);
      assert.equal(
        (await cold.doc.request({ operation: 'read', state: cold.doc.status().state, sheet: dedupSheet, range: 'D3' }))
          .cells.D3.value,
        'record-4'
      );
      assert.equal((await preview(cold.doc, dedupSheet, 'A1:E11', true, dedupColumns)).duplicatePreview.removedRows, 0);
    } finally {
      await cold.doc.close();
    }
    return { previewNoWrites: true, oneUndoRedo: true, saveCold: true };
  } finally {
    await doc.close();
  }
});
check('stale fingerprint and closed fields refuse before head/source publication', async () => {
  const f = await fixture(simple()),
    { doc } = await owner(f);
  try {
    const old = (await preview(doc)).duplicatePreview;
    await apply(doc, { sheet: 'Sheet1', cell: 'A3', value: 2 });
    const before = doc.status();
    for (const bad of [
      remove(old),
      { ...remove(old), expectedFingerprint: '0'.repeat(64) },
      { ...remove(old), permutation: [2, 4] },
      { ...remove(old), columns: ['A', 'A'] },
    ]) {
      await assert.rejects(apply(doc, bad));
      assert.equal(doc.status().state, before.state);
      assert.deepEqual(await fs.readFile(f.source), f.original);
    }
    await assert.rejects(
      doc.request({
        operation: 'duplicatePreview',
        documentId: before.documentId,
        state: old.state,
        sheet: 'Sheet1',
        ...config,
      }),
      /stale/i
    );
    const n = await nativeFixture(simple()),
      page = await directPreview(n);
    for (const patch of [
      { version: 999 },
      { source: 'missing.xlsx' },
      { expectedFingerprint: '0'.repeat(64) },
      { columns: ['A', 'A'] },
      { permutation: [2, 4] },
    ])
      await assert.rejects(directApply(n, page.fingerprint, patch));
    assert.deepEqual(await fs.readFile(n.source), n.original);
    return { closedNativeAndOwner: true };
  } finally {
    await doc.close();
  }
});
check('sparse blanks compact only selected cells and protected dependency/style contexts refuse', async () => {
  const f = await nativeFixture(
      bytes(
        `<row r="1">${str('A1', 'key')}</row><row r="4">${num('A4', '2')}${str('B4', 'kept')}</row><row r="5">${num('A5', '2')}</row>`
      )
    ),
    page = await directPreview(f, { range: 'A1:B5' });
  assert.equal(page.keptRows, 2);
  await directApply(f, page.fingerprint, { range: 'A1:B5' });
  const view = readWorkbook(await fs.readFile(path.join(f.root, 'dedup.xlsx')));
  assert.equal(view.sheets[0].cells.A3.value, 2);
  assert.equal(view.sheets[0].cells.A4, undefined);
  for (const variant of ['outside', 'dynamic', 'implicit-style', 'active-filter', 'offbranch-chart']) {
    let body = `<row r="1">${str('A1', 'key')}</row><row r="2">${num('A2', '1')}${num('B2', '1')}</row><row r="3">${num('A3', '1')}${num('B3', '2')}</row><row r="4">${num('A4', '2')}${num('B4', '3')}</row>`;
    if (variant === 'outside') body += `<row r="8"><c r="A8"><f>SUM(A2:A4)</f><v>4</v></c></row>`;
    if (variant === 'dynamic') body = body.replace(num('B4', '3'), '<c r="B4"><f>INDIRECT("A4")</f><v>2</v></c>');
    if (variant === 'implicit-style') body = body.replace('<row r="4">', '<row r="4" s="1" customFormat="1">');
    let original = bytes(
      body,
      false,
      variant === 'active-filter'
        ? '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter>'
        : ''
    );
    if (variant === 'offbranch-chart') {
      const packageParts = unzipSync(original);
      packageParts['xl/charts/chart1.xml'] = strToU8(
        '<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
      );
      original = Buffer.from(zipSync(packageParts));
    }
    const n = await nativeFixture(original),
      p = await directPreview(n);
    await assert.rejects(directApply(n, p.fingerprint), undefined, `${variant} must refuse`);
    assert.deepEqual(await fs.readFile(n.source), n.original);
  }
  return { sparseCompaction: true, protectedProfiles: 5 };
});
check('Manual pending keys refuse and readonly native owner preview does not grant writes', async () => {
  const m = await fixture(
      bytes(
        `<row r="1">${str('A1', 'key')}</row><row r="2"><c r="A2"><f>1</f><v>913</v></c></row><row r="3"><c r="A3"><f>1</f><v>912</v></c></row>`,
        true
      )
    ),
    { doc } = await owner(m);
  try {
    await assert.rejects(preview(doc, 'Sheet1', 'A1:B3'), /pending|Calculate/i);
    assert.equal(doc.status().dirty, false);
  } finally {
    await doc.close();
  }
  const p = unzipSync(simple());
  p['[Content_Types].xml'] = strToU8(
    strFromU8(p['[Content_Types].xml']).replace(
      '</Types>',
      // An ActiveX control part keeps the workbook read-only (macros alone no longer do).
      '<Override PartName="/custom/dedup-owned.bin" ContentType="application/vnd.ms-office.activeX"/></Types>'
    )
  );
  const r = await fixture(Buffer.from(zipSync(p))),
    o = await owner(r);
  try {
    assert(o.doc.status().view.readOnly);
    const page = (await preview(o.doc)).duplicatePreview;
    assert.equal(page.removedRows, 1);
    await assert.rejects(apply(o.doc, remove(page)), /read.only/i);
    assert.equal(o.doc.status().dirty, false);
    assert.deepEqual(await fs.readFile(r.source), r.original);
  } finally {
    await o.doc.close();
  }
  return { pendingAndReadonly: true };
});
check('late actual native candidate failure rolls back accepted owner head/history/source', async () => {
  const f = await fixture(simple());
  let fail = false;
  const { doc } = await owner(f, async (input, client) => {
    if (fail && input.operation === 'index') throw Error('Intentional candidate indexing failure');
    return client.call(input);
  });
  try {
    const page = (await preview(doc)).duplicatePreview,
      before = doc.status();
    fail = true;
    await assert.rejects(apply(doc, remove(page)), /Intentional/);
    fail = false;
    assert.equal(doc.status().state, before.state);
    assert.equal(doc.status().canUndo, before.canUndo);
    assert.deepEqual(await fs.readFile(f.source), f.original);
    assert.equal((await preview(doc)).duplicatePreview.fingerprint, page.fingerprint);
    return { actualBuiltCandidateNotAccepted: true };
  } finally {
    await doc.close();
  }
});
check('Filter Off active/inactive restores lexical baseline then new range and Sort survive cold', async () => {
  const b = bytes(
      `<row r="1">${str('A1', 'key')}</row><row r="2" hidden="true">${num('A2', '1')}${str('B2', 'first')}</row><row r="3" hidden="false">${num('A3', '2')}${str('B3', 'last')}</row><row r="4">${num('A4', '1')}</row>`
    ),
    act = { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B4', column: 'A', values: [1] };
  for (const inactive of [false, true]) {
    const session = new WorkbookSession(b),
      base = await session.read(),
      first = await session.apply({ operationId: 'apply', expectedRevision: base.revision, commands: [act] });
    let current = first;
    if (inactive)
      current = await session.apply({
        operationId: 'clear',
        expectedRevision: current.view.revision,
        commands: [
          { operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: current.view.sheets[0].filter.revision },
        ],
      });
    const off = await session.apply({
      operationId: 'off',
      expectedRevision: current.view.revision,
      commands: [
        { operation: 'removeFilter', sheet: 'Sheet1', expectedFilterRevision: current.view.sheets[0].filter.revision },
      ],
    });
    assert.equal(off.view.sheets[0].filter, undefined);
    const xml = strFromU8(unzipSync(off.bytes)['xl/worksheets/sheet1.xml']);
    assert(!xml.includes('autoFilter'));
    assert(xml.includes('hidden="true"') && xml.includes('hidden="false"'));
    const f = await fixture(b),
      { doc } = await owner(f);
    try {
      await apply(doc, act);
      if (inactive)
        await apply(doc, {
          operation: 'clearFilter',
          sheet: 'Sheet1',
          expectedFilterRevision: doc.status().view.sheets[0].filter.revision,
        });
      const old = doc.status().state,
        revision = doc.status().view.sheets[0].filter.revision;
      await apply(doc, { operation: 'removeFilter', sheet: 'Sheet1', expectedFilterRevision: revision });
      const offState = doc.status().state;
      assert.equal(doc.status().view.sheets[0].filter, undefined);
      await doc.request({ operation: 'undo' });
      assert.equal(doc.status().state, old);
      await doc.request({ operation: 'redo' });
      assert.equal(doc.status().state, offState);
      await apply(doc, { ...act, range: 'A1:B3' });
      await apply(doc, {
        operation: 'removeFilter',
        sheet: 'Sheet1',
        expectedFilterRevision: doc.status().view.sheets[0].filter.revision,
      });
      await apply(doc, { operation: 'sortRows', sheet: 'Sheet1', range: 'A2:B4', column: 'A', direction: 'desc' });
      await doc.request({ operation: 'save' });
      await doc.close();
      const cold = await owner(f);
      try {
        assert.equal(cold.doc.status().view.sheets[0].filter, undefined);
        assert.equal(cold.doc.status().dirty, false);
      } finally {
        await cold.doc.close();
      }
    } finally {
      await doc.close();
    }
  }
  return { activeInactive: 2, portableAndNative: true, newRangeAndSort: true };
});

check('owned filter null-only and empty-only criteria reapply and survive exact typed cold state', async () => {
  for (const value of [null, '']) {
    const f = await fixture(
        bytes(
          `<row r="1">${str('A1', 'key')}</row><row r="2">${str('B2', 'blank-record')}</row><row r="3">${str('A3', '')}${str('B3', 'empty-record')}</row><row r="4">${num('A4', '1')}${str('B4', 'numeric-record')}</row>`
        )
      ),
      { doc } = await owner(f);
    try {
      const command = { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B4', column: 'A', values: [value] };
      await apply(doc, command);
      const want = value === null ? [2, 3] : [1, 3];
      assert.deepEqual(doc.status().view.sheets[0].hiddenRows, want);
      assert.deepEqual(doc.status().view.sheets[0].filter.values, [value]);
      const copied = await copy(doc, f, 'typed-blank'),
        copiedView = readWorkbook(await fs.readFile(copied));
      assert.deepEqual(copiedView.sheets[0].hiddenRows, want);
      assert.deepEqual(copiedView.sheets[0].filter.values, [value]);
      await doc.request({ operation: 'save' });
      await doc.close();
      const cold = await owner(f);
      try {
        assert.deepEqual(cold.doc.status().view.sheets[0].filter.values, [value]);
        assert.deepEqual(cold.doc.status().view.sheets[0].hiddenRows, want);
        await apply(cold.doc, { sheet: 'Sheet1', cell: 'A4', value: value });
        await apply(cold.doc, {
          operation: 'reapplyFilter',
          sheet: 'Sheet1',
          expectedFilterRevision: cold.doc.status().view.sheets[0].filter.revision,
        });
        assert.deepEqual(cold.doc.status().view.sheets[0].hiddenRows, value === null ? [2] : [1]);
      } finally {
        await cold.doc.close();
      }
    } finally {
      await doc.close();
    }
  }
  return { nullAndEmptyDistinct: true, actualColdReapply: true };
});
