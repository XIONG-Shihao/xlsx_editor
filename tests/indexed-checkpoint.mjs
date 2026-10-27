import assert from 'node:assert/strict';
import { test, after } from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createHash } from 'node:crypto';
import { DOMParser } from '@xmldom/xmldom';
import { createWorkbook, readWorkbook } from '../core/index.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { indexedCache, indexedCacheUsage, fileDigest } from '../core/indexed-cache.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { readWorkbookMetadata } from '../core/workbook-package.mjs';
import { indexedCalculationReusable, publishIndexedCheckpoint } from '../core/indexed-checkpoint.mjs';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const roots = [];
after(async () => {
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
async function fixture(formula = 'A1*2') {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-checkpoint-test-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>${formula}</f><v>999</v></c></row></sheetData></worksheet>`
  );
  parts['custom/preserved.bin'] = new Uint8Array([0, 255, 17]);
  await fs.writeFile(source, zipSync(parts));
  const client = createWorkbookNativeClient(nativeWorkbookBinary()),
    calls = [];
  const native = async (input) => {
    calls.push(input);
    return client.call(input);
  };
  const make = () =>
    new IndexedWorkbook(
      native,
      () => assert.fail('legacy writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
  return {
    root,
    source,
    parts,
    client,
    calls,
    native,
    make,
    options: { cacheIdentity: 'verified-test-policy', cacheOptions: { directory: path.join(root, 'cache') } },
  };
}

test('an exact-source fresh cache resets transferred live extent while preserving calculation certification', async () => {
  const f = await fixture(),
    parts = { ...f.parts };
  let book = f.make(),
    cached,
    cold;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><dimension ref="A1:Z80"/><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><f>A1*2</f><v>14</v></c></row><row r="80"><c r="Z80" s="0"/></row></sheetData></worksheet>`
  );
  await fs.writeFile(f.source, zipSync(parts));
  const original = await fileDigest(f.source);
  try {
    await book.open(f.source, f.options);
    const next = await book.apply('0', [{ operation: 'resizeRow', sheet: 'Sheet1', index: 0, height: 24 }]),
      state = next.view.storage.state;
    const live = await book.usedRange(state, 'Sheet1');
    assert.equal(live.status, 'complete');
    assert.equal(live.lastCell, 'Z80');
    const query = f.calls.findLast((call) => call.operation === 'used-range'),
      root = query.root,
      index = JSON.parse(await fs.readFile(path.join(root, 'index.json'), 'utf8'));
    assert.equal(index.sourceFactSchema, 'workbook.xlsx.used-range-facts.v2');
    const saved = path.join(f.root, 'exact-private-source.xlsx');
    await fs.copyFile(path.join(root, 'source.xlsx'), saved);
    const savedBytes = await fs.readFile(saved),
      xml = strFromU8(unzipSync(savedBytes)['xl/worksheets/sheet1.xml']);
    assert(!/\br="Z80"/.test(xml));
    assert(/<dimension\b[^>]*\bref="A1:B1"/.test(xml));
    // An actual export establishes cache-status's version-bound export proof.
    // The independently copied source is byte-exact to its native index. This
    // isolates the hash-equal cache boundary even if ZIP serialization differs.
    await book.export(state, { publication: true });
    const { storage, revision, ...view } = next.view,
      current = {
        root,
        index,
        view,
        version: query.version,
        styles: readWorkbookMetadata(await fs.readFile(path.join(root, 'metadata.xlsx'))).styles,
      };
    const identity = 'xlsx-b16-preserved-calculation-editing:verified-test-policy:null';
    const publication = await publishIndexedCheckpoint({
      source: saved,
      identity,
      current,
      native: f.native,
      options: f.options.cacheOptions,
    });
    assert(publication.published || publication.existing);
    assert.equal(
      (await book.usedRange(state, 'Sheet1')).lastCell,
      'Z80',
      'Cache sanitization must not reset the live source owner'
    );
    cached = f.make();
    f.calls.length = 0;
    await cached.open(saved, f.options);
    assert.equal(cached.cacheHit, true);
    assert(f.calls.some((call) => call.operation === 'used-range-source-snapshot'));
    assert(
      !f.calls.some((call) => call.operation === 'index'),
      'A sanitized exact-source checkpoint retains its calculation cache'
    );
    const cachedExtent = await cached.usedRange('0', 'Sheet1');
    cold = f.make();
    await cold.open(saved);
    const coldExtent = await cold.usedRange('0', 'Sheet1');
    assert.equal(cachedExtent.range, coldExtent.range);
    assert.equal(cachedExtent.lastCell, 'B1');
    assert.equal(cachedExtent.empty, coldExtent.empty);
    const formula = (await cached.read('0', 'Sheet1', 'B1')).cells.B1;
    assert.equal(formula.formula, 'A1*2');
    assert.equal(formula.calculation.status, 'calculated');
    assert.equal(formula.calculation.value, 14);
    assert.equal(await fileDigest(f.source), original);
    assert.deepEqual(await fs.readFile(saved), savedBytes);
  } finally {
    await Promise.all([book, cached, cold].filter(Boolean).map((owner) => owner.close()));
    f.client.close();
  }
});

test('changed saves reimport exact source spans, formulas and styles across edits/undo/reopen', async () => {
  const f = await fixture(),
    before = await fileDigest(f.source);
  let book = f.make();
  try {
    await book.open(f.source, f.options);
    let next = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A1', value: 11 },
      { sheet: 'Sheet1', cell: 'B1', bold: true },
    ]);
    const state = next.view.storage.state,
      output = await book.export(state);
    assert.equal(book.cachePublication.published, false);
    const saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(output.path, saved);
    const revision = await fileDigest(saved),
      bytes = await fs.readFile(saved);
    assert.equal(readWorkbook(bytes).sheets[0].cells.B1.cachedValue, 22);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 7);
    await book.close();
    book = f.make();
    f.calls.length = 0;
    const view = await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal(view.revision, revision);
    assert(
      f.calls.some((call) => call.operation === 'index'),
      'Changed saved packages must import their exact source fields'
    );
    const cells = (await book.read('0', 'Sheet1', 'A1:B1')).cells;
    assert.equal(cells.B1.calculation.value, 22);
    assert.equal(cells.B1.bold, true);
    assert.equal(cells.B1.formula, 'A1*2');
    const savedAgain = await book.export('0');
    assert.equal(savedAgain.revision, revision);
    assert.deepEqual(await fs.readFile(savedAgain.path), bytes, 'unmodified restored state exports exact saved bytes');
    next = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A1', value: 13 },
      { sheet: 'Sheet1', cell: 'B1', italic: true },
    ]);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 26);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 22);
    assert.equal((await book.export('0')).revision, revision);
    const changed = await book.export(next.view.storage.state),
      changedBytes = await fs.readFile(changed.path),
      reparsed = readWorkbook(changedBytes);
    assert.equal(reparsed.sheets[0].cells.B1.cachedValue, 26);
    assert.equal(reparsed.sheets[0].cells.B1.bold, true);
    assert.equal(reparsed.sheets[0].cells.B1.italic, true);
    assert.deepEqual(unzipSync(changedBytes)['custom/preserved.bin'], f.parts['custom/preserved.bin']);
    assert.equal(await fileDigest(f.source), before);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('checkpoint structural receipts exclude unchanged integer overlays across paginated base comparisons', async () => {
  const f = await fixture(),
    parts = { ...f.parts };
  let book = f.make();
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData/></worksheet>`);
  await fs.writeFile(f.source, zipSync(parts));
  const original = await fileDigest(f.source);
  try {
    await book.open(f.source, f.options);
    let next = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A1', value: 'Key' },
      { sheet: 'Sheet1', cell: 'B1', value: 'Value' },
      { sheet: 'Sheet1', cell: 'A2', value: 30 },
      { sheet: 'Sheet1', cell: 'B2', value: 3 },
      { sheet: 'Sheet1', cell: 'A3', value: 10 },
      { sheet: 'Sheet1', cell: 'B3', value: 1, bold: true },
    ]);
    next = await book.apply(next.view.storage.state, [
      { type: 'fillRange', sheet: 'Sheet1', source: 'B3', range: 'B4:B1200' },
    ]);
    next = await book.apply(next.view.storage.state, [{ sheet: 'Sheet1', cell: 'D1', formula: 'SUM(B4:B1200)' }]);
    const exported = await book.export(next.view.storage.state),
      saved = path.join(f.root, 'saved.xlsx');
    assert.equal(book.cachePublication.published, false);
    await fs.copyFile(exported.path, saved);
    const savedHash = await fileDigest(saved);
    await book.close();
    book = f.make();
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    const sorted = await book.apply('0', [
        { operation: 'sortRows', sheet: 'Sheet1', range: 'A2:B3', column: 'A', direction: 'asc' },
      ]),
      state = sorted.view.storage.state;
    const receipt = await book.affected('0', state);
    assert.equal(receipt.total, 4);
    assert.deepEqual(receipt.affected.map((cell) => cell.cell).sort(), ['A2', 'A3', 'B2', 'B3']);
    assert.equal(
      (await book.read('0', 'Sheet1', 'A2')).cells.A2.value,
      30,
      'The saved import remains an immutable prior state'
    );
    assert.equal((await book.read(state, 'Sheet1', 'A2')).cells.A2.value, 10);
    for (const revision of ['0', state]) {
      const cells = (await book.read(revision, 'Sheet1', 'B4:D1200')).cells;
      for (let row = 4; row <= 1200; row++) {
        assert.equal(cells[`B${row}`].value, 1);
        assert.equal(cells[`B${row}`].bold, true);
      }
      const formula = (await book.read(revision, 'Sheet1', 'D1')).cells.D1;
      assert.equal(formula.formula, 'SUM(B4:B1200)');
      assert.equal(formula.calculation.value, 1197);
    }
    const filtered = await book.apply(state, [
      { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B3', column: 'A', values: [30] },
    ]);
    assert.equal(
      (await book.affected(state, filtered.view.storage.state)).total,
      0,
      'A row visibility change does not author cell edits'
    );
    const output = await book.export(state),
      bytes = await fs.readFile(output.path),
      disk = readWorkbook(bytes).sheets[0].cells;
    assert.equal(disk.A2.value, 10);
    assert.equal(disk.A3.value, 30);
    assert.equal(disk.B1200.value, 1);
    assert.equal(disk.B1200.bold, true);
    assert.equal(disk.D1.formula, 'SUM(B4:B1200)');
    assert.equal(disk.D1.cachedValue, 1197);
    assert.deepEqual(unzipSync(bytes)['custom/preserved.bin'], parts['custom/preserved.bin']);
    assert.equal(await fileDigest(f.source), original);
    assert.equal(await fileDigest(saved), savedHash);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('dynamic, volatile, unsupported and failed formula results never become trusted saved checkpoints', async () => {
  for (const formula of ['INDIRECT("A1")', 'RAND()', '1/0', 'IF(1=1,2,RAND())', 'HEX2DEC("FF")']) {
    const f = await fixture(formula),
      book = f.make();
    try {
      await book.open(f.source, f.options);
      const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]);
      await book.export(next.view.storage.state);
      assert.equal(book.cachePublication.published, false, formula);
      assert.match(book.cachePublication.reason, /saved|plan|calculation|unsupported|function/i, formula);
    } finally {
      await book.close();
      f.client.close();
    }
  }
});

test('external saved edits, policy changes and corrupted original-source cache artifacts rebuild safely', async () => {
  const f = await fixture(),
    original = await fileDigest(f.source);
  let book = f.make();
  try {
    await book.open(f.source, f.options);
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]);
    const exported = await book.export(next.view.storage.state),
      saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(exported.path, saved);
    await book.close();
    book = f.make();
    await book.open(saved, { ...f.options, cacheIdentity: 'changed-calculation-policy' });
    assert.equal(book.cacheHit, false);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 22);
    await book.close();
    book = f.make();
    await book.open(saved, f.options);
    await book.close();
    const revision = await fileDigest(saved),
      entries = await fs.readdir(f.options.cacheOptions.directory);
    let corrupted = 0;
    for (const entry of entries) {
      if (!/^[a-f0-9]{64}$/.test(entry)) continue;
      const root = path.join(f.options.cacheOptions.directory, entry),
        manifest = JSON.parse(await fs.readFile(path.join(root, 'cache.json'), 'utf8'));
      if (manifest.revision === revision && !manifest.checkpoint) {
        const file = path.join(root, 'cells.sqlite'),
          bytes = await fs.readFile(file);
        bytes[0] ^= 1;
        await fs.writeFile(file, bytes);
        corrupted++;
      }
    }
    assert(corrupted > 0, 'An actual saved-source cache artifact must be corrupted');
    book = f.make();
    f.calls.length = 0;
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert(f.calls.some((call) => call.operation === 'index'));
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 22);
    await book.close();
    const parts = unzipSync(await fs.readFile(saved));
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
        '<v xmlns="' + ns + '">11</v>',
        '<v xmlns="' + ns + '">17</v>'
      )
    );
    await fs.writeFile(saved, zipSync(parts));
    book = f.make();
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 34);
    assert.equal(await fileDigest(f.source), original);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('ten changed save/import cycles stay cache-bounded and preserve three independent owners', async () => {
  const f = await fixture(),
    saved = path.join(f.root, 'saved.xlsx');
  await fs.copyFile(f.source, saved);
  let book;
  try {
    for (let i = 0; i < 10; i++) {
      book = f.make();
      await book.open(saved, f.options);
      if (i) assert.equal(book.cacheHit, false);
      const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: i + 20 }]);
      const result = await book.export(next.view.storage.state);
      assert.equal(book.cachePublication.published, false);
      assert.match(book.cachePublication.reason, /fresh import/);
      await fs.copyFile(result.path, saved);
      await book.close();
      book = null;
      const usage = await indexedCacheUsage(f.options.cacheOptions);
      assert(usage.entries <= 8);
      assert(usage.bytes <= usage.maxBytes);
      assert.equal(usage.staging, 0);
      for (const entry of await fs.readdir(f.options.cacheOptions.directory))
        if (/^[a-f0-9]{64}$/.test(entry)) {
          const root = path.join(f.options.cacheOptions.directory, entry),
            manifest = JSON.parse(await fs.readFile(path.join(root, 'cache.json'), 'utf8'));
          assert(
            manifest.files.every(
              (file) => !['export.xlsx', 'layout.xlsx', 'transform.xlsx', 'sort-order.jsonl'].includes(file.name)
            )
          );
        }
    }
    const owners = [f.make(), f.make(), f.make()];
    try {
      for (const owner of owners) await owner.open(saved, f.options);
      await owners[0].apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 100 }]);
      for (const owner of owners) assert.equal((await owner.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 58);
    } finally {
      for (const owner of owners) await owner.close();
    }
  } finally {
    await book?.close();
    f.client.close();
  }
});

test('named references and table boundaries remain revision-qualified through structural reopen', async () => {
  const { editWorkbook } = await import('../core/workbook-package.mjs');
  const f = await fixture();
  let book = f.make();
  try {
    let bytes = editWorkbook(
      createWorkbook(),
      [
        ['A1', 'Qty'],
        ['B1', 'Cost'],
        ['A2', 2],
        ['B2', 3],
        ['A3', 4],
        ['B3', 5],
      ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
    );
    bytes = editWorkbook(bytes, [{ operation: 'createTable', sheet: 'Sheet1', range: 'A1:B3', tableName: 'Sales' }]);
    bytes = editWorkbook(bytes, [{ sheet: 'Sheet1', cell: 'E1', formula: 'SUM(Sales[Cost])+Rate' }]);
    const parts = unzipSync(bytes);
    parts['xl/workbook.xml'] = strToU8(
      strFromU8(parts['xl/workbook.xml']).replace(
        '</workbook>',
        '<definedNames><definedName name="Rate">Sheet1!$B$2</definedName></definedNames></workbook>'
      )
    );
    await fs.writeFile(f.source, zipSync(parts));
    await book.open(f.source, f.options);
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 11);
    let next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B3', value: 9 }]);
    let output = await book.export(next.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    const saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(output.path, saved);
    await book.close();
    book = f.make();
    let view = await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal(view.tables[0].ref, 'A1:B3');
    assert.equal(view.namedRanges[0].formula, 'Sheet1!$B$2');
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 15);
    next = await book.apply('0', [{ operation: 'insertRow', sheet: 'Sheet1', index: 2 }]);
    next = await book.apply(next.view.storage.state, [
      { sheet: 'Sheet1', cell: 'A3', value: 7 },
      { sheet: 'Sheet1', cell: 'B3', value: 4 },
    ]);
    output = await book.export(next.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    await fs.copyFile(output.path, saved);
    await book.close();
    book = f.make();
    view = await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal(view.tables[0].ref, 'A1:B4');
    assert.equal(view.namedRanges[0].formula, 'Sheet1!$B$2');
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 19);
    await book.close();
    const changed = unzipSync(await fs.readFile(saved));
    changed['xl/workbook.xml'] = strToU8(strFromU8(changed['xl/workbook.xml']).replace('Sheet1!$B$2', 'Sheet1!$B$4'));
    await fs.writeFile(saved, zipSync(changed));
    book = f.make();
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 25);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('sparse shared-group edits survive fresh saved imports, reopened branches and another save', async () => {
  const f = await fixture(),
    size = 1024,
    parts = unzipSync(createWorkbook());
  let book = f.make();
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData>${Array.from({ length: size }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ` ref="A1:A${size}"` : ''}>${i === 0 ? '1+1' : ''}</f><v>999</v></c>${i === size - 1 ? `<c r="B${size}"><f>SUM(A1:A${size})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  await fs.writeFile(f.source, zipSync(parts));
  try {
    await book.open(f.source, f.options);
    const edit = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A500', value: 5 }]);
    let output = await book.export(edit.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    const saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(output.path, saved);
    assert.equal((await book.read('0', 'Sheet1', 'A500')).cells.A500.calculation.value, 2);
    await book.close();
    book = f.make();
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal((await book.read('0', 'Sheet1', 'B1024')).cells.B1024.calculation.value, 2051);
    assert.equal((await book.read('0', 'Sheet1', 'A1024')).cells.A1024.formula, '1+1');
    const again = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A700', value: 7 }]);
    assert.equal((await book.read(again.view.storage.state, 'Sheet1', 'B1024')).cells.B1024.calculation.value, 2056);
    assert.equal((await book.read('0', 'Sheet1', 'B1024')).cells.B1024.calculation.value, 2051);
    output = await book.export(again.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    await fs.copyFile(output.path, saved);
    await book.close();
    book = f.make();
    await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal((await book.read('0', 'Sheet1', 'B1024')).cells.B1024.calculation.value, 2056);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('checkpoint metadata matches a cold import after new distant cells, styles, formulas and clears', async () => {
  const f = await fixture();
  let book = f.make(),
    cold;
  try {
    await book.open(f.source, f.options);
    let next = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'X70', value: 'new' },
      { sheet: 'Sheet1', cell: 'C40', bold: true },
      { sheet: 'Sheet1', cell: 'Z80', formula: 'SUM(A1:C1)' },
    ]);
    next = await book.apply(next.view.storage.state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'B1' }]);
    const output = await book.export(next.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    assert.match(book.cachePublication.reason, /fresh import/);
    const saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(output.path, saved);
    const savedXml = strFromU8(unzipSync(await fs.readFile(saved))['xl/worksheets/sheet1.xml']),
      savedDoc = new DOMParser({
        errorHandler: {
          warning() {},
          error(message) {
            throw Error(message);
          },
          fatalError(message) {
            throw Error(message);
          },
        },
      }).parseFromString(savedXml, 'application/xml');
    const savedCells = Array.from(savedDoc.getElementsByTagNameNS(ns, 'c'));
    assert.deepEqual(savedCells.map((cell) => cell.getAttribute('r')).sort(), ['A1', 'C40', 'X70', 'Z80']);
    assert(
      !savedCells.some((cell) => cell.getAttribute('r') === 'B1'),
      'Cleared default B1 must be absent from saved XML'
    );
    const blank = savedCells.find((cell) => cell.getAttribute('r') === 'C40');
    assert.equal(blank.getAttribute('s'), '1');
    assert.equal(
      Array.from(blank.childNodes).filter((node) => node.nodeType === 1).length,
      0,
      'The meaningful styled blank remains a physical empty tag'
    );
    await book.close();
    book = f.make();
    const cachedView = await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    cold = f.make();
    const coldView = await cold.open(saved);
    for (const field of ['storedCells', 'storedFormulaCount', 'rows', 'columns'])
      assert.equal(cachedView.sheets[0][field], coldView.sheets[0][field], field);
    assert.equal(cachedView.sheets[0].storedCells, 4);
    assert.equal(cachedView.sheets[0].storedFormulaCount, 1);
    assert.equal(cachedView.sheets[0].rows, 80);
    assert.equal(cachedView.sheets[0].columns, 26);
    for (const range of ['A1:B1', 'C40', 'X70', 'Z80']) {
      const cached = (await book.read('0', 'Sheet1', range)).cells,
        actual = (await cold.read('0', 'Sheet1', range)).cells;
      for (const ref of Object.keys(actual))
        for (const key of ['value', 'formula', 'bold', 'italic', 'format'])
          if (!(key === 'value' && actual[ref].formula))
            assert.equal(cached[ref][key], actual[ref][key], `${ref}.${key}`);
    }
    assert.equal((await book.read('0', 'Sheet1', 'Z80')).cells.Z80.calculation.value, 7);
    assert.equal((await cold.read('0', 'Sheet1', 'Z80')).cells.Z80.calculation.value, 7);
    const retainedBlank = (await book.read('0', 'Sheet1', 'C40')).cells.C40;
    assert.equal(retainedBlank.value, null);
    assert.equal(retainedBlank.bold, true);
  } finally {
    await book.close();
    await cold?.close();
    f.client.close();
  }
});

test('saved-only cache mode reimports changed saved packages instead of borrowing old source fields', async () => {
  const f = await fixture(),
    options = { ...f.options, cacheMode: 'saved-only' };
  let book = f.make();
  try {
    await book.open(f.source, options);
    assert.equal(book.cacheHit, false);
    assert.equal((await indexedCacheUsage(options.cacheOptions)).entries, 0);
    assert(!f.calls.some((call) => call.method === 'cache-status'));
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]);
    const exported = await book.export(next.view.storage.state);
    assert.equal(book.cachePublication.published, false);
    assert.equal((await indexedCacheUsage(options.cacheOptions)).entries, 0);
    const saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(exported.path, saved);
    await book.close();
    book = f.make();
    f.calls.length = 0;
    await book.open(saved, options);
    assert.equal(book.cacheHit, false);
    assert(
      f.calls.some((call) => call.operation === 'index'),
      'Saved-only mode must also import exact rewritten source'
    );
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 22);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('a prior exported checkpoint under the same explicit policy identity is invalidated before saved-source import', async () => {
  const f = await fixture(),
    identity = 'xlsx-b16-preserved-calculation-editing:verified-test-policy:null';
  let book = f.make();
  try {
    await book.open(f.source, f.options);
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]);
    const exported = await book.export(edited.view.storage.state),
      saved = path.join(f.root, 'saved.xlsx');
    await fs.copyFile(exported.path, saved);
    assert.equal(book.cachePublication.published, false);
    assert.match(book.cachePublication.reason, /fresh import/);
    // Reproduce the prior producer using actual original native spans and the
    // accepted SQL version, then key it by the rewritten saved package. Its
    // declared policy is deliberately identical to the reopening owner.
    const last = f.calls.findLast((call) => call.operation === 'export'),
      root = last.root,
      index = JSON.parse(await fs.readFile(path.join(root, 'index.json'), 'utf8'));
    const { storage, revision, ...view } = edited.view,
      current = { root, index, view, version: last.version };
    const qualification = await indexedCalculationReusable(current, f.native, { exported: true });
    assert.equal(qualification.qualified, true);
    const qualifiedView = {
      ...view,
      sheets: view.sheets.map((sheet) => ({
        ...sheet,
        ...Object.fromEntries(
          ['storedCells', 'storedFormulaCount', 'rows', 'columns'].map((key) => [
            key,
            qualification.sheets.find((item) => item.path === sheet.path)[key],
          ])
        ),
      })),
    };
    const graph = createHash('sha256')
      .update(
        JSON.stringify({
          sheets: qualifiedView.sheets.map((s) => ({ name: s.name, path: s.path, date1904: !!s.date1904 })),
          names: qualifiedView.namedRanges,
          tables: qualifiedView.tables,
        })
      )
      .digest('hex');
    const metadata = createHash('sha256')
      .update(JSON.stringify([graph, qualifiedView.calculationPolicy ?? null]))
      .digest('hex');
    const cache = await indexedCache(saved, identity, f.options.cacheOptions),
      styles = readWorkbookMetadata(await fs.readFile(path.join(root, 'metadata.xlsx'))).styles;
    assert.notEqual(cache.revision, index.revision);
    const legacy = {
      schema: 'workbook.xlsx.exported-checkpoint.v1',
      publicRevision: cache.revision,
      baseRevision: index.revision,
      version: last.version,
      policy: identity,
      metadata,
      view: qualifiedView,
      styles,
      logicalSheets: qualification.sheets,
    };
    assert(
      await cache.publish(root, last.version, { checkpoint: legacy }),
      'The old cache shape must actually exist before reopening'
    );
    await book.close();
    book = f.make();
    f.calls.length = 0;
    const reopened = await book.open(saved, f.options);
    assert.equal(book.cacheHit, false);
    assert.equal(reopened.revision, cache.revision);
    assert(f.calls.some((call) => call.operation === 'index'));
    const canonical = readWorkbook(await fs.readFile(saved)).sheets[0].cells,
      actual = (await book.read('0', 'Sheet1', 'A1:B1')).cells;
    for (const [ref, cell] of Object.entries(actual)) {
      const { calculation, cachedResult, ...source } = cell;
      assert.deepEqual(source, canonical[ref], ref);
    }
    assert.equal(actual.B1.calculation.value, 22);
    assert.equal(actual.B1.calculation.revision, cache.revision);
    const fresh = [];
    for (const name of await fs.readdir(f.options.cacheOptions.directory)) {
      if (!/^[a-f0-9]{64}$/.test(name)) continue;
      const manifest = JSON.parse(
        await fs.readFile(path.join(f.options.cacheOptions.directory, name, 'cache.json'), 'utf8')
      );
      if (manifest.revision === cache.revision && manifest.identity === identity) fresh.push(manifest);
    }
    assert.equal(fresh.length, 1);
    assert.equal(fresh[0].checkpoint, false, 'The rejected entry is replaced only by exact-source cache data');
  } finally {
    await book.close();
    f.client.close();
  }
});

test('owner diagnostics expose scalar lifecycle counts and bounded projection accounting', async () => {
  const f = await fixture('HEX2DEC("FF")'),
    book = f.make();
  try {
    assert.equal(book.diagnosticStats().retainedStates, 0);
    assert.equal(book.diagnosticStats().storageRoots, 0);
    assert.equal(book.diagnosticStats().closed, 0);
    await book.open(f.source);
    const first = book.diagnosticStats();
    assert.equal(first.retainedStates, 1);
    assert.equal(first.storageRoots, 1);
    assert.equal(first.calculationCalls, 1);
    assert(first.lastProjectedBytes > 0);
    assert(first.lastInputCells > 0);
    first.retainedStates = 99;
    assert.equal(book.diagnosticStats().retainedStates, 1, 'The diagnostic snapshot cannot mutate ownership');
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'A1*2' }]);
    const current = book.diagnosticStats();
    assert.equal(current.retainedStates, 2);
    assert.equal(current.calculationCalls, 2);
    assert.equal(current.lastProjectedBytes, 0);
    assert.equal(current.lastEvaluatedFormulaCells, 1);
    assert.equal(current.peakProjectedBytes, first.peakProjectedBytes);
    assert(Object.values(current).every((value) => typeof value === 'number' && Number.isFinite(value)));
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 14);
    const styled = await book.apply(next.view.storage.state, [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:A100', format: { bold: true } },
    ]);
    await book.export(styled.view.storage.state);
    const exported = book.diagnosticStats();
    assert.equal(exported.lastExportStyleCount, 1, 'One style combination must not scale with formatted cells');
    assert(exported.lastExportMetadataBytes > 0);
    await book.compact(['0']);
    assert.equal(book.diagnosticStats().retainedStates, 1);
    await book.close();
    assert.equal(book.diagnosticStats().retainedStates, 0);
    assert.equal(book.diagnosticStats().storageRoots, 0);
    assert.equal(book.diagnosticStats().closed, 1);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('invalid cache modes and closed owners reject before allocating session storage', async () => {
  const f = await fixture(),
    book = f.make(),
    storageDirectory = path.join(f.root, 'never-created');
  try {
    await assert.rejects(book.open(f.source, { cacheMode: 'unknown', storageDirectory }), {
      code: 'session-malformed',
    });
    await assert.rejects(fs.stat(storageDirectory), { code: 'ENOENT' });
    assert.equal(f.calls.length, 0);
    await book.close();
    await assert.rejects(book.open(f.source, { ...f.options, storageDirectory }), /closed/);
    await assert.rejects(fs.stat(storageDirectory), { code: 'ENOENT' });
    assert.equal(f.calls.length, 0);
  } finally {
    await book.close();
    f.client.close();
  }
});

test('cached UI and structural operations match cold import after first/last formula and shared normalization', async () => {
  const { editWorkbook } = await import('../core/workbook-package.mjs');
  for (const kind of ['first', 'last', 'shared-normalized', 'shared-retained']) {
    const f = await fixture(),
      shared = kind.startsWith('shared');
    let book = f.make(),
      cold;
    if (kind === 'first')
      await fs.writeFile(
        f.source,
        editWorkbook(createWorkbook(), [
          { sheet: 'Sheet1', cell: 'A1', value: 7 },
          { sheet: 'Sheet1', cell: 'B1', value: 2 },
        ])
      );
    if (shared) {
      const parts = unzipSync(createWorkbook());
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        `<worksheet xmlns="${ns}"><sheetData>${Array.from({ length: 16 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ' ref="A1:A16"' : ''}>${i === 0 ? '1+1' : ''}</f><v>999</v></c>${i === 15 ? '<c r="B16"><f>SUM(A1:A16)</f></c>' : ''}</row>`).join('')}</sheetData></worksheet>`
      );
      await fs.writeFile(f.source, zipSync(parts));
    }
    try {
      await book.open(f.source, f.options);
      const command =
        kind === 'first'
          ? { cell: 'B1', formula: 'A1*2' }
          : kind === 'last'
            ? { cell: 'B1', value: 8 }
            : kind === 'shared-normalized'
              ? { cell: 'A8', value: 5 }
              : { cell: 'A1', bold: true };
      const next = await book.apply('0', [{ sheet: 'Sheet1', ...command }]),
        exported = await book.export(next.view.storage.state);
      assert.equal(book.cachePublication.published, false, kind);
      const saved = path.join(f.root, 'saved.xlsx');
      await fs.copyFile(exported.path, saved);
      await book.close();
      const xml = strFromU8(unzipSync(await fs.readFile(saved))['xl/worksheets/sheet1.xml']);
      if (shared) assert.equal((xml.match(/t="shared"/g) ?? []).length, kind === 'shared-retained' ? 16 : 0);
      book = f.make();
      const warmView = await book.open(saved, f.options);
      assert.equal(book.cacheHit, false);
      cold = f.make();
      const coldView = await cold.open(saved);
      assert.equal(warmView.structuralReason, coldView.structuralReason, kind);
      assert.deepEqual(warmView.sheets[0].capabilities, coldView.sheets[0].capabilities, kind);
      if (kind === 'last') assert.equal(warmView.structuralReason, undefined);
      else assert.match(warmView.structuralReason, /formula-reference/);
      // Shared groups no longer hold structure back: they are written out in full first.
      assert.equal(warmView.sheets[0].capabilities.insertRows, true);
      const outputs = [];
      for (const owner of [book, cold]) {
        const inserted = await owner.apply('0', [{ operation: 'insertRow', sheet: 'Sheet1', index: 1 }]);
        const deleted = await owner.apply(inserted.view.storage.state, [
          { operation: 'deleteRow', sheet: 'Sheet1', index: 1 },
        ]);
        outputs.push(readWorkbook(await fs.readFile((await owner.export(deleted.view.storage.state)).path)));
      }
      for (const ref of shared ? ['A1', 'A8', 'A16', 'B16'] : ['A1', 'B1'])
        for (const key of ['value', 'formula', 'cachedValue', 'bold'])
          assert.equal(
            outputs[0].sheets[0].cells[ref]?.[key],
            outputs[1].sheets[0].cells[ref]?.[key],
            `${kind}:${ref}:${key}`
          );
    } finally {
      await book.close();
      await cold?.close();
      f.client.close();
    }
  }
});
