import { assertRetainedPart } from '../scripts/test/package-oracle.mjs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { test, after } from 'node:test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { IndexedWorkbook } from '@office-editors/xlsx-editor/core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '@office-editors/xlsx-editor/core/indexed-calculation.mjs';
import {
  nativeWorkbookOperationAsync as native,
  closeNativeWorkbookWorker as closeNativeDocumentWorker,
} from '@office-editors/xlsx-editor/native';
import { createWorkbook, readWorkbook, editCalculatedWorkbookResult } from '@office-editors/xlsx-editor/engine';
const roots = [];
after(async () => {
  closeNativeDocumentWorker();
  await Promise.all(roots.map((root) => fs.rm(root, { recursive: true, force: true })));
});
async function fixture(xml, strings) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-index-test-'));
  roots.push(root);
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(xml);
  if (strings) parts['xl/sharedStrings.xml'] = strToU8(strings);
  parts['custom/opaque.bin'] = new Uint8Array([0, 1, 2, 255]);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  return { root, source, parts };
}
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sheet = (body) => `<worksheet xmlns="${ns}"><dimension ref="A1:C3"/><sheetData>${body}</sheetData></worksheet>`;
async function owner(source, transport = native, onLegacy = () => {}) {
  const book = new IndexedWorkbook(
    transport,
    async (source, edits, out) => {
      onLegacy(edits);
      const result = await editCalculatedWorkbookResult(await fs.readFile(source), edits);
      await fs.writeFile(out, result.bytes);
      return { revision: result.view.revision, edits: result.edits };
    },
    (input) => calculateIndexedWorkbook(input, transport)
  );
  await book.open(source);
  return book;
}
test('native range preserves typed values, raw precision, text escapes and styles', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>12345678901234567890</v></c><c r="C1" t="inlineStr"><is><t>_x005F_x0041_&#13; &amp; 中文</t></is></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2" t="e"><v>#DIV/0!</v></c><c r="C2"/></row>'
    ),
    `<sst xmlns="${ns}"><si><r><t>rich </t></r><r><t>text</t></r><rPh><t>annotation</t></rPh></si></sst>`
  );
  const expected = readWorkbook(await fs.readFile(source));
  const book = await owner(source);
  try {
    assert.equal(Object.keys(book.view('0').sheets[0].cells).length, 0);
    assert.deepEqual((await book.read('0', 'Sheet1', 'A1:C3')).cells, expected.sheets[0].cells);
  } finally {
    await book.close();
  }
});
test('native export inserts new/empty rows, expands dimensions, preserves opaque members and version branches', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is><extLst><ext uri="test"/></extLst></c><c r="C1"><v>9</v></c></row><row r="3" ht="24"/><row r="5"><c r="A5"><v>5</v></c></row>'
    )
  );
  const book = await owner(source);
  try {
    const edits = [
      { sheet: 'Sheet1', cell: 'A1', value: '_x000A_\r\n<& 中文' },
      { sheet: 'Sheet1', cell: 'B1', value: 2 },
      { sheet: 'Sheet1', cell: 'C1', value: 3 },
      { sheet: 'Sheet1', cell: 'B2', value: true },
      { sheet: 'Sheet1', cell: 'B3', value: 'empty-row' },
      { sheet: 'Sheet1', cell: 'P6', value: 6 },
    ];
    const next = await book.apply('0', edits),
      state = next.view.storage.state;
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'old');
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 'branch' }]);
    assert.equal((await book.read(state, 'Sheet1', 'A1')).cells.A1.value, edits[0].value);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'A1')).cells.A1.value, 'branch');
    const output = await book.export(state),
      bytes = await fs.readFile(output.path),
      saved = readWorkbook(bytes),
      actual = unzipSync(bytes);
    for (const edit of edits) assert.equal(saved.sheets[0].cells[edit.cell].value, edit.value);
    for (const name of Object.keys(parts))
      if (name !== 'xl/worksheets/sheet1.xml') assertRetainedPart(name, parts[name], actual[name]);
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /ext uri="test"/);
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /ref="A1:P6"/);
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /ht="24"/);
  } finally {
    await book.close();
  }
});
test('empty sheetData export, first formula projection and structural fallback', async () => {
  const { source } = await fixture(`<worksheet xmlns="${ns}"><sheetData/></worksheet>`);
  const book = await owner(source);
  try {
    let result = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A1', value: 7 },
      { sheet: 'Sheet1', cell: 'B1', value: 8 },
    ]);
    const literal = await book.export(result.view.storage.state);
    assert.equal(readWorkbook(await fs.readFile(literal.path)).sheets[0].cells.A1.value, 7);
    result = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'C1', formula: 'SUM(A1:B1)' }]);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C1')).cells;
    assert.equal(cells.C1.calculation.status, 'calculated');
    assert.equal(cells.C1.calculation.value, 15);
    const formula = await book.export(result.view.storage.state);
    assert.equal(readWorkbook(await fs.readFile(formula.path)).sheets[0].cells.C1.cachedValue, 15);
    const formatted = await book.apply(result.view.storage.state, [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true } },
    ]);
    assert.equal((await book.read(formatted.view.storage.state, 'Sheet1', 'A1')).cells.A1.bold, true);
    const resized = await book.apply(formatted.view.storage.state, [
      { sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 25 },
    ]);
    assert.equal(
      (await book.read(resized.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.revision,
      resized.view.revision
    );
  } finally {
    await book.close();
  }
});
test('scan pages include offscreen base cells and overlay insertions exactly once', async () => {
  const { source } = await fixture(
    sheet(Array.from({ length: 4200 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join(''))
  );
  const book = await owner(source);
  try {
    const result = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'B2050', value: 'new' },
      { sheet: 'Sheet1', cell: 'A3000', value: 'changed' },
    ]);
    let after;
    const found = new Map();
    do {
      const page = await book.scan(result.view.storage.state, 'Sheet1', after);
      for (const [ref, cell] of Object.entries(page.cells)) {
        assert(!found.has(ref));
        found.set(ref, cell.value);
      }
      after = page.nextCursor;
    } while (after);
    assert.equal(found.size, 4201);
    assert.equal(found.get('A3000'), 'changed');
    assert.equal(found.get('B2050'), 'new');
    assert.equal(found.get('A4200'), 4199);
  } finally {
    await book.close();
  }
});
test('sparse formula discovery crosses page boundaries and preserves branch overlays and old indexes', async () => {
  const rows = Array.from(
    { length: 513 },
    (_, i) => `<row r="${i + 1}"><c r="B${i + 1}"><f>${i + 1}</f></c><c r="C${i + 1}"><v>99</v></c></row>`
  ).join('');
  const { source } = await fixture(sheet(rows));
  let root;
  const transport = async (request) => {
    root = request.root;
    return native(request);
  };
  const book = await owner(source, transport);
  try {
    const sheetPath = book.view('0').sheets[0].path;
    const changes = [
      ...['B1', 'B257', 'B512'].map((cell) => ({ sheet: sheetPath, cell, value: { value: 42, dataType: 'number' } })),
      ...[
        ['A1', '4'],
        ['A514', '9'],
        ['A515', '11'],
        ['B513', '7'],
      ].map(([cell, formula]) => ({ sheet: sheetPath, cell, value: { formula } })),
    ];
    const patched = await native({ action: 'workbook', operation: 'patch', root, version: 0, changes });
    const expected = new Map(Array.from({ length: 513 }, (_, i) => [`B${i + 1}`, String(i + 1)]));
    for (const change of changes)
      if (change.value.formula) expected.set(change.cell, change.value.formula);
      else expected.delete(change.cell);
    async function scan(version, useCursor) {
      const seen = new Map();
      let offset = 0,
        after;
      do {
        const page = await native({
          action: 'workbook',
          operation: 'range',
          root,
          version,
          sheet: sheetPath,
          range: 'A1:XFD1048576',
          formulas: true,
          limit: 127,
          ...(useCursor ? { after } : { offset }),
        });
        for (const [ref, cell] of page.cells) {
          assert(!seen.has(ref), ref);
          assert(cell.formula, ref);
          seen.set(ref, cell.formula);
        }
        offset = page.nextOffset;
        after = page.nextCursor;
      } while (useCursor ? after : offset !== null);
      return seen;
    }
    for (const useCursor of [false, true]) assert.deepEqual(await scan(patched.version, useCursor), expected);
    const result = await calculateIndexedWorkbook(
      { root, version: patched.version, revision: 'page-test', view: book.view('0') },
      native
    );
    assert.equal(result.nativeSqlite, true);
    assert.equal(result.evaluatedFormulaCells, 513);
    const values = await native({
      action: 'workbook',
      operation: 'range',
      root,
      version: result.version,
      sheet: sheetPath,
      range: 'A513:B515',
    });
    assert.equal(new Map(values.cells).get('A515').calculation.value, 11);
    assert.equal(new Map(values.cells).get('B513').calculation.value, 7);
    const original = await scan(0, true);
    assert.equal(original.size, 513);
    assert.equal(original.get('B1'), '1');
    assert(!original.has('A515'));
    // Existing persisted indexes without the optimization must remain readable.
    const { DatabaseSync } = await import('node:sqlite');
    const db = new DatabaseSync(path.join(root, 'cells.sqlite'));
    try {
      assert(db.prepare("SELECT name FROM sqlite_master WHERE name='cells_formulas'").get());
      db.exec('DROP INDEX cells_formulas');
    } finally {
      db.close();
    }
    assert.deepEqual(await scan(patched.version, true), expected);
    const legacy = await calculateIndexedWorkbook(
      { root, version: patched.version, revision: 'old-index-test', view: book.view('0') },
      native
    );
    assert.equal(legacy.nativeSqlite, true);
    assert.equal(legacy.evaluatedFormulaCells, 513);
  } finally {
    await book.close();
  }
});
test('failed atomic batch leaves prior accepted native state intact', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>2</v></c></row>'));
  const book = await owner(source);
  try {
    await assert.rejects(
      book.apply('0', [
        { sheet: 'Sheet1', cell: 'A1', value: 8 },
        { sheet: 'Sheet1', cell: 'A2', value: 3, expectedValue: 999 },
      ])
    );
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 2);
  } finally {
    await book.close();
  }
});
test('format-only export retains precise numeric text and rich text runs', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>12345678901234567890</v></c><c r="B1" t="inlineStr"><is><r><t>rich </t></r><r><t>text</t></r></is></c></row>'
    )
  );
  const book = await owner(source);
  try {
    const changed = await book.apply('0', [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true, numberFormat: 'text' } },
    ]);
    const saved = await book.export(changed.view.storage.state),
      bytes = await fs.readFile(saved.path),
      cells = readWorkbook(bytes).sheets[0].cells;
    assert.equal(cells.A1.rawValue, '12345678901234567890');
    assert.equal(cells.A1.bold, true);
    assert.equal(cells.B1.bold, true);
    assert.match(
      strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']),
      /<is><r><t>rich <\/t><\/r><r><t>text<\/t><\/r><\/is>/
    );
  } finally {
    await book.close();
  }
});
test('native owner accepts relationship-targeted worksheet paths outside xl/worksheets', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>7</v></c></row>'));
  const parts = unzipSync(await fs.readFile(source));
  parts['xl/data/custom.xml'] = parts['xl/worksheets/sheet1.xml'];
  delete parts['xl/worksheets/sheet1.xml'];
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace('xl/worksheets/sheet1.xml', 'xl/data/custom.xml')
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace('worksheets/sheet1.xml', 'data/custom.xml')
  );
  await fs.writeFile(source, zipSync(parts));
  const book = await owner(source);
  try {
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 7);
  } finally {
    await book.close();
  }
});
test('negative zero and overflowing stored numbers survive native and JSON transports', async () => {
  const { encodeOwnerResult, decodeOwnerResult, ownerCommands, decodedCommands } = await import(
    '@office-editors/xlsx-editor/transport'
  );
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>-0</v></c><c r="B1"><v>1e999</v></c></row>')),
    book = await owner(source);
  try {
    const read = await book.read('0', 'Sheet1', 'A1:B1'),
      wire = JSON.parse(JSON.stringify(encodeOwnerResult(read))),
      decoded = decodeOwnerResult(wire);
    assert(Object.is(decoded.cells.A1.value, -0));
    assert.equal(decoded.cells.B1.value, Infinity);
    const commands = decodedCommands(
      JSON.parse(JSON.stringify(ownerCommands([{ sheet: 'Sheet1', cell: 'C1', value: -0 }])))
    );
    assert(Object.is(commands[0].value, -0));
    const edited = await book.apply('0', commands);
    assert(Object.is((await book.read(edited.view.storage.state, 'Sheet1', 'C1')).cells.C1.value, -0));
    const saved = await book.export(edited.view.storage.state);
    assert(Object.is(readWorkbook(await fs.readFile(saved.path)).sheets[0].cells.C1.value, -0));
  } finally {
    await book.close();
  }
});
test('crash cleanup removes only dead indexed owners and keeps live leases', async () => {
  const parent = path.join(
    os.tmpdir(),
    `workbook-indexes-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
  );
  await fs.mkdir(parent, { recursive: true });
  const dead = await fs.mkdtemp(path.join(parent, 'owner-')),
    live = await fs.mkdtemp(path.join(parent, 'owner-'));
  await fs.writeFile(
    path.join(dead, 'owner.json'),
    JSON.stringify({ schema: 'workbook.workbook.index-owner.v1', pid: 2147483647 })
  );
  await fs.writeFile(
    path.join(live, 'owner.json'),
    JSON.stringify({ schema: 'workbook.workbook.index-owner.v1', pid: process.pid })
  );
  const { createIndexedStorage } = await import(
      import.meta.resolve('@office-editors/xlsx-editor/core/indexed-storage.mjs') + `?lease=${Date.now()}`
    ),
    created = await createIndexedStorage('cleanup-test');
  try {
    await assert.rejects(fs.stat(dead), { code: 'ENOENT' });
    assert((await fs.stat(live)).isDirectory());
    assert((await fs.stat(created)).isDirectory());
  } finally {
    await fs.rm(live, { recursive: true, force: true });
    await fs.rm(created, { recursive: true, force: true });
  }
});
test('SpreadsheetML escapes are decoded per text run, without joining partial escapes', async () => {
  const runs = '<r><t>_x00</t></r><r><t>41_</t></r>';
  const { source } = await fixture(
      sheet(`<row r="1"><c r="A1" t="inlineStr"><is>${runs}</is></c><c r="B1" t="s"><v>0</v></c></row>`),
      `<sst xmlns="${ns}"><e:si xmlns:e="urn:extension"/><si>${runs}</si></sst>`
    ),
    book = await owner(source);
  try {
    const cells = (await book.read('0', 'Sheet1', 'A1:B1')).cells;
    assert.equal(cells.A1.value, '_x0041_');
    assert.equal(cells.B1.value, '_x0041_');
  } finally {
    await book.close();
  }
});

test('large native format retains compact commands, streams unique styles and replays recovery', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>12345678901234567890</v></c><c r="B1" t="inlineStr"><is><r><t>rich </t></r><r><t>text</t></r></is><extLst><ext uri="kept"/></extLst></c></row>'
    )
  );
  const calls = [],
    legacy = [];
  const transport = async (input) => {
    calls.push({
      operation: input.operation,
      bytes: JSON.stringify(input).length,
      hasChanges: Object.hasOwn(input, 'changes'),
    });
    return native(input);
  };
  const book = await owner(source, transport, (edits) => legacy.push(edits));
  let recovery;
  try {
    calls.length = 0;
    const commands = [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:B6000', format: { bold: true, numberFormat: 'text' } },
    ];
    const result = await book.apply('0', commands);
    recovery = JSON.parse(JSON.stringify(result.edits));
    assert.deepEqual(recovery, commands);
    assert.equal(JSON.stringify(recovery).length < 200, true);
    assert.deepEqual(
      calls.map((c) => c.operation),
      ['format']
    );
    assert.equal(legacy.length, 0);
    assert.equal((await book.read('0', 'Sheet1', 'B6000')).cells.B6000, undefined);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'B6000')).cells.B6000.bold, true);
    calls.length = 0;
    const saved = await book.export(result.view.storage.state),
      bytes = await fs.readFile(saved.path),
      actual = unzipSync(bytes),
      cells = readWorkbook(bytes).sheets[0].cells;
    assert.equal(Object.keys(cells).length, 12000);
    for (const cell of Object.values(cells)) {
      assert.equal(cell.bold, true);
      assert.equal(cell.format, 49);
    }
    assert.equal(cells.A1.rawValue, '12345678901234567890');
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /<is><r><t>rich <\/t><\/r><r><t>text<\/t><\/r><\/is>/);
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /ext uri="kept"/);
    for (const name of Object.keys(parts))
      if (!['xl/worksheets/sheet1.xml', 'xl/styles.xml'].includes(name))
        assertRetainedPart(name, parts[name], actual[name]);
    assert.deepEqual(
      calls.map((c) => c.operation),
      ['export-plan', 'export-styles', 'artifact-write', 'export']
    );
    assert(calls.every((c) => c.bytes < (c.operation === 'artifact-write' ? 16384 : 500) && !c.hasChanges));
    assert.equal(legacy.length, 0);
  } finally {
    await book.close();
  }
  const restored = await owner(source);
  try {
    const applied = await restored.apply('0', recovery);
    assert.equal((await restored.read(applied.view.storage.state, 'Sheet1', 'B6000')).cells.B6000.bold, true);
  } finally {
    await restored.close();
  }
});

test('native format validates atomically, preserves special formulas and keeps branches isolated', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>4</v></c></row>'));
  let root;
  const book = await owner(source, async (input) => {
    root = input.root;
    return native(input);
  });
  const format = (range, format) => ({ type: 'formatRange', sheet: 'Sheet1', range, format });
  try {
    await assert.rejects(
      book.apply('0', [format('A1:A40', { bold: true }), format('A50', { fontSize: 100 })]),
      /Invalid range format/
    );
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'range',
        root,
        version: 1,
        sheet: 'xl/worksheets/sheet1.xml',
        range: 'A1',
      }),
      { code: 'internal-unknown-reference' }
    );
    for (const invalid of [{ styleId: 9999 }, { bold: 'yes' }, { value: 'not formatting' }, { fontFamily: '\u0000' }])
      await assert.rejects(book.apply('0', [format('A1', invalid)]));
    const bold = await book.apply('0', [format('A1:B40', { bold: true, fontColor: 'FF0000' })]);
    const italic = await book.apply(bold.view.storage.state, [
      format('B20', { italic: true }),
      format('B20', { styleId: 0, underline: true }),
    ]);
    const branch = await book.apply('0', [format('B20', { italic: true })]);
    for (const [state, expected] of [
      [bold.view.storage.state, { bold: true }],
      [italic.view.storage.state, { bold: false, underline: true }],
      [branch.view.storage.state, { bold: false, italic: true }],
    ]) {
      const saved = await book.export(state),
        cell = readWorkbook(await fs.readFile(saved.path)).sheets[0].cells.B20;
      for (const [key, value] of Object.entries(expected)) assert.equal(!!cell[key], value, key);
    }
    const structural = await book.apply(bold.view.storage.state, [
      { sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 30 },
    ]);
    assert.deepEqual(structural.edits, [{ sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 30 }]);
    assert.equal((await book.read(structural.view.storage.state, 'Sheet1', 'B40')).cells.B40.bold, true);
  } finally {
    await book.close();
  }
  const special = await fixture(sheet('<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">1</f><v>1</v></c></row>'));
  await native({ action: 'workbook', operation: 'index', root: special.root });
  const styled = await native({
    action: 'workbook',
    operation: 'format',
    root: special.root,
    version: 0,
    commands: [{ sheet: 'xl/worksheets/sheet1.xml', range: 'A1:A2', format: { bold: true } }],
  });
  const page = await native({
    action: 'workbook',
    operation: 'range',
    root: special.root,
    version: styled.version,
    sheet: 'xl/worksheets/sheet1.xml',
    range: 'A1',
  });
  assert.equal(page.cells[0][1].bold, true);
  assert.equal(page.cells[0][1].formulaType, 'shared');
  assert.equal(
    (
      await native({
        action: 'workbook',
        operation: 'range',
        root: special.root,
        version: 0,
        sheet: 'xl/worksheets/sheet1.xml',
        range: 'A1',
      })
    ).cells[0][1].bold,
    undefined
  );
});

test('native export honors latest value/formula intent and style-only numeric fidelity', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>12345678901234567890</v></c><c r="B1"><v>3</v></c></row>')
  );
  const book = await owner(source);
  try {
    let result = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A1', bold: true },
      { sheet: 'Sheet1', cell: 'B1', formula: '2+3' },
    ]);
    result = await book.apply(result.view.storage.state, [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { italic: true } },
    ]);
    const formulaState = result.view.storage.state;
    result = await book.apply(formulaState, [
      { sheet: 'Sheet1', cell: 'B1', value: 9 },
      { sheet: 'Sheet1', cell: 'B1', bold: true },
    ]);
    let saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B1.value, 9);
    assert.equal(saved.sheets[0].cells.B1.formula, undefined);
    assert.equal(saved.sheets[0].cells.B1.italic, true);
    assert.equal(saved.sheets[0].cells.A1.rawValue, '12345678901234567890');
    assert.equal(saved.sheets[0].cells.A1.bold, true);
    saved = readWorkbook(await fs.readFile((await book.export(formulaState)).path));
    assert.equal(saved.sheets[0].cells.B1.cachedValue, 5);
  } finally {
    await book.close();
  }
});

test('mixed recovery journal keeps range commands compact and publishes only a complete batch', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'));
  const book = await owner(source, native, () => assert.fail('Mixed recovery must not use the legacy writer'));
  const commands = [
    { sheet: 'Sheet1', cell: 'A1', value: 2, expectedValue: 1 },
    { type: 'formatRange', sheet: 'Sheet1', range: 'A1:A12000', format: { bold: true } },
    { sheet: 'Sheet1', cell: 'A1', value: 3, expectedValue: 2 },
  ];
  try {
    await assert.rejects(
      book.apply('0', [...commands, { sheet: 'Sheet1', cell: 'A1', value: 4, expectedValue: 999 }]),
      /Conflict/
    );
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 1);
    const result = await book.apply('0', JSON.parse(JSON.stringify(commands)));
    assert.deepEqual(result.edits, commands);
    const first = (await book.read(result.view.storage.state, 'Sheet1', 'A1')).cells.A1;
    assert.equal(first.value, 3);
    assert.equal(first.bold, true);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'A12000')).cells.A12000.bold, true);
    const exported = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(exported.sheets[0].cells.A1.value, 3);
    assert.equal(exported.sheets[0].cells.A12000.bold, true);
  } finally {
    await book.close();
  }
});

test('failed streamed export cannot replace the last complete exported file', async () => {
  const { root } = await fixture(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'));
  const index = await native({ action: 'workbook', operation: 'index', root });
  const request = { action: 'workbook', operation: 'export', root, revision: index.revision };
  await native({ ...request, changes: [{ sheet: 'xl/worksheets/sheet1.xml', cell: 'A1', value: 2 }] });
  const before = await fs.readFile(path.join(root, 'export.xlsx'));
  await assert.rejects(
    native({ ...request, changes: [{ sheet: 'xl/worksheets/sheet1.xml', cell: 'A1', value: '\u0000' }] }),
    { code: /^(note-malformed|package-malformed)$/ }
  );
  assert.deepEqual(await fs.readFile(path.join(root, 'export.xlsx')), before);
});

test('native clear/fill/paste values keep compact mixed journals and snapshot overlapping sources', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>')
  );
  const legacy = [],
    book = await owner(source, native, (edits) => legacy.push(edits));
  try {
    let result = await book.apply('0', [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1', format: { bold: true, numberFormat: 'percent' } },
      { type: 'fillRange', sheet: 'Sheet1', source: 'A1:B1', range: 'B1:C1' },
    ]);
    let cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C1')).cells;
    assert.equal(cells.B1.value, 1);
    assert.equal(cells.B1.bold, true);
    assert.equal(cells.B1.format, 10);
    assert.equal(cells.C1.value, 2);
    result = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', cell: 'A2', formula: '4+5' },
      { type: 'clearRange', sheet: 'Sheet1', range: 'A1:C1' },
    ]);
    cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C2')).cells;
    assert.equal(cells.B1.value, null);
    assert.equal(cells.B1.bold, true);
    assert.equal(cells.A2.calculation.value, 9);
    result = await book.apply(result.view.storage.state, [
      { type: 'pasteValues', sheet: 'Sheet1', source: 'A2', range: 'B1:B12000' },
    ]);
    assert.equal(result.edits.length, 1);
    assert.equal(result.edits[0].type, 'pasteValues');
    cells = (await book.read(result.view.storage.state, 'Sheet1', 'B11999:B12000')).cells;
    assert.equal(cells.B12000.value, 9);
    assert.equal(cells.B12000.formula, undefined);
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B12000.value, 9);
    assert.equal(saved.sheets[0].cells.B1.bold, true);
    assert.equal(saved.sheets[0].cells.A2.cachedValue, 9);
    assert.equal(legacy.length, 0);
    result = await book.apply(result.view.storage.state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'A2' }]);
    assert.equal(result.view.sheets[0].storedFormulaCount, 0);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 1);
  } finally {
    await book.close();
  }
});

test('native copy refuses precision loss and preserves formula-copy compatibility', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>12345678901234567890</v></c></row>'));
  const book = await owner(source);
  try {
    await assert.rejects(
      book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'A1', range: 'A2:A3' }]),
      /precision/
    );
    const literal = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'B1', value: 5 },
      { sheet: 'Sheet1', cell: 'C1', formula: 'B1*2' },
    ]);
    const copied = await book.apply(literal.view.storage.state, [
      { type: 'fillRange', sheet: 'Sheet1', source: 'C1', range: 'D1' },
    ]);
    const cell = (await book.read(copied.view.storage.state, 'Sheet1', 'D1')).cells.D1;
    assert.equal(cell.formula, 'C1*2');
    assert.equal(cell.calculation.value, 20);
  } finally {
    await book.close();
  }
});

test('native search bounds each work page and counts matches beyond returned items', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 9000 },
        (_, i) =>
          `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${i === 8999 ? '远处 NEEDLE' : 'match'}</t></is></c></row>`
      ).join('')
    )
  );
  const book = await owner(source);
  try {
    let after,
      total = 0,
      pages = 0;
    do {
      const page = await book.query('0', 'Sheet1', 'MATCH', 'raw', after);
      assert(page.scanned <= 8192);
      assert(Object.keys(page.cells).length <= 1000);
      total += page.matches;
      after = page.nextCursor;
      pages++;
    } while (after);
    assert.equal(total, 8999);
    assert.equal(pages, 2);
    const first = await book.query('0', 'Sheet1', 'needle', 'raw');
    assert.equal(first.matches, 0);
    const last = await book.query('0', 'Sheet1', 'needle', 'raw', first.nextCursor);
    assert.equal(last.cells.A9000.value, '远处 NEEDLE');
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B8999', formula: '2+2' }]);
    const found = await book.query(edited.view.storage.state, 'Sheet1', '2+2', 'formula', 'A8998');
    assert.equal(found.cells.B8999.formula, '2+2');
  } finally {
    await book.close();
  }
});

test('native layout changes preserve values, row attributes, column properties and undo roots', async () => {
  const { source, parts } = await fixture(
    `<x:worksheet xmlns:x="${ns}" xmlns:e="urn:opaque"><x:dimension ref="A1:C3"/><x:sheetViews><x:sheetView workbookViewId="0" zoomScale="80"/></x:sheetViews><x:cols><x:col min="1" max="3" width="10" style="0"/></x:cols><x:sheetData><x:row r="1" hidden="1"><x:c r="A1"><x:v>12345678901234567890</x:v></x:c></x:row><x:row r="3" customFormat="1"/></x:sheetData><x:extLst><e:row r="1" ht="7"/></x:extLst></x:worksheet>`
  );
  const book = await owner(source, native, () => assert.fail('Layout must not use the full JavaScript writer'));
  try {
    let result = await book.apply('0', [{ sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 30 }]);
    result = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', operation: 'resizeRow', index: 1, height: 40 },
    ]);
    result = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', operation: 'resizeRow', index: 2, height: 50 },
    ]);
    result = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', operation: 'resizeColumn', index: 1, width: 25 },
    ]);
    result = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', operation: 'freeze', rows: 1, columns: 1 },
    ]);
    const saved = await fs.readFile((await book.export(result.view.storage.state)).path),
      actual = unzipSync(saved),
      view = readWorkbook(saved);
    assert.equal(view.sheets[0].cells.A1.rawValue, '12345678901234567890');
    assert.deepEqual(view.sheets[0].rowHeights, [
      { row: 0, height: 30 },
      { row: 1, height: 40 },
      { row: 2, height: 50 },
    ]);
    assert(view.sheets[0].hiddenRows.includes(0));
    assert.equal(view.sheets[0].freeze.rows, 1);
    assert.deepEqual(
      view.sheets[0].widths.map((w) => [w.min, w.max, w.width]),
      [
        [1, 1, 10],
        [2, 2, 25],
        [3, 3, 10],
      ]
    );
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /zoomScale="80"/);
    assert.match(strFromU8(actual['xl/worksheets/sheet1.xml']), /e:row r="1" ht="7"/);
    for (const name of Object.keys(parts))
      if (name !== 'xl/worksheets/sheet1.xml') assertRetainedPart(name, parts[name], actual[name]);
    assert.equal(book.view('0').sheets[0].rowHeights.length, 0);
  } finally {
    await book.close();
  }
});

test('calculation coalesces prefix ranges and skips only unrelated static edits', async () => {
  const rows = Array.from(
    { length: 500 },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c>${i < 100 ? `<c r="B${i + 1}"><f>SUM(A1:A${i + 100})</f></c>` : ''}</row>`
  ).join('');
  const { source } = await fixture(sheet(rows)),
    calculations = [];
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('Unexpected fallback'),
    async (input) => {
      const result = await calculateIndexedWorkbook(input, native);
      calculations.push(result);
      return result;
    }
  );
  try {
    await book.open(source);
    assert(calculations[0].rangeReads < 10, JSON.stringify(calculations[0]));
    assert.equal((await book.read('0', 'Sheet1', 'B100')).cells.B100.calculation.value, (199 * 200) / 2);
    let result = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D500', value: 8 }]);
    assert.equal(calculations.at(-1).skipped, true);
    result = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 2 }]);
    assert.notEqual(calculations.at(-1).skipped, true);
    assert.equal(
      (await book.read(result.view.storage.state, 'Sheet1', 'B100')).cells.B100.calculation.value,
      (199 * 200) / 2 + 1
    );
    result = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'D499', formula: '1+2' }]);
    assert.notEqual(calculations.at(-1).skipped, true);
  } finally {
    await book.close();
  }
});

test('compact Paste Values recovery evaluates preceding formula edits before copying', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>2</v></c></row>'));
  const journal = [
    { sheet: 'Sheet1', cell: 'B1', formula: 'A1*2' },
    { sheet: 'Sheet1', cell: 'A1', value: 7 },
    { type: 'pasteValues', sheet: 'Sheet1', source: 'B1', range: 'C1:C3' },
    { sheet: 'Sheet1', cell: 'A1', value: 9 },
  ];
  for (const structural of [false, true]) {
    const book = await owner(source);
    try {
      const result = await book.apply(
        '0',
        structural ? [...journal, { sheet: 'Sheet1', operation: 'resizeRow', index: 0, height: 30 }] : journal
      );
      const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C3')).cells;
      assert.equal(cells.C3.value, 14);
      assert.equal(cells.B1.calculation.value, 18);
      const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
      assert.equal(saved.sheets[0].cells.C3.value, 14);
      assert.equal(saved.sheets[0].cells.B1.cachedValue, 18);
      const chain = await book.apply('0', [
        { sheet: 'Sheet1', cell: 'B1', formula: 'A1*2' },
        { type: 'pasteValues', sheet: 'Sheet1', source: 'B1', range: 'A1' },
        { type: 'pasteValues', sheet: 'Sheet1', source: 'B1', range: 'C1' },
      ]);
      assert.equal((await book.read(chain.view.storage.state, 'Sheet1', 'C1')).cells.C1.value, 8);
    } finally {
      await book.close();
    }
  }
});
test('native keyed queries page duplicate matches and guarded updates resolve unique keys near storage', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>first</t></is></c><c r="B2"><v>2</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>last</t></is></c><c r="B3"><v>3</v></c></row>'
    )
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected fallback');
  });
  try {
    const keys = await book.keys('0', 'Sheet1', 'A2:B3', 'A', ['last', 'missing'], { offset: 0, limit: 1 });
    assert.deepEqual(keys.rows, [2]);
    assert.deepEqual(keys.matches, [1, 0]);
    const changed = await book.apply('0', [
      {
        type: 'updateRows',
        sheet: 'Sheet1',
        range: 'A2:B3',
        keyColumn: 'A',
        updates: [{ key: 'last', cells: [{ column: 'B', value: 7, expectedValue: 3 }] }],
      },
    ]);
    assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'B3')).cells.B3.value, 7);
    const duplicate = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A3', value: 'first' }]);
    assert.deepEqual(
      (await book.keys(duplicate.view.storage.state, 'Sheet1', 'A2:B5', 'A', ['first', null], {})).matches,
      [2, 2]
    );
    await assert.rejects(
      book.apply(duplicate.view.storage.state, [
        {
          type: 'updateRows',
          sheet: 'Sheet1',
          range: 'A2:B3',
          keyColumn: 'A',
          updates: [{ key: 'first', cells: [{ column: 'B', value: 8, expectedValue: 2 }] }],
        },
      ]),
      { code: /^(csv-duplicate|duplicates-duplicate|workbook-duplicate)$/ }
    );
    await assert.rejects(
      book.apply('0', [
        {
          type: 'updateRows',
          sheet: 'Sheet1',
          range: 'A2:B3',
          keyColumn: 'A',
          updates: [{ key: 'last', cells: [{ column: 'B', value: 8, expectedValue: 99 }] }],
        },
      ]),
      /Conflict/
    );
    assert.equal((await book.read('0', 'Sheet1', 'B3')).cells.B3.value, 3);
  } finally {
    await book.close();
  }
});
test('selective calculation follows transitive dependencies and retains independent formula definitions', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f></c><c r="C1"><f>B1+1</f></c><c r="D1"><v>10</v></c><c r="E1"><f>D1*3</f></c></row>'
    )
  );
  const results = [];
  const book = new IndexedWorkbook(
    native,
    () => {
      throw Error('Unexpected fallback');
    },
    async (input) => {
      const result = await calculateIndexedWorkbook(input, native);
      results.push(result);
      return result;
    }
  );
  try {
    await book.open(source);
    const result = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 4 }]);
    assert.equal(results.at(-1).evaluatedFormulaCells, 2);
    assert.equal(results.at(-1).mode, 'affected-native-sqlite');
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:E1')).cells;
    assert.equal(cells.B1.calculation.value, 8);
    assert.equal(cells.C1.calculation.value, 9);
    assert.equal(cells.E1.calculation.value, 30);
    assert.equal(cells.E1.formula, 'D1*3');
    const output = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(output.sheets[0].cells.E1.formula, 'D1*3');
    assert.equal(output.sheets[0].cells.C1.cachedValue, 9);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D1', value: 20 }]);
    assert.equal(results.at(-1).evaluatedFormulaCells, 1);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'C1:E1')).cells.C1.calculation.value, 5);
  } finally {
    await book.close();
  }
});
test('native sort and filter preserve row metadata, rich text and unaffected package members', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1" ht="29" customHeight="1"><c r="A1"><v>30</v></c><c r="B1" t="inlineStr"><is><r><t>thirty</t></r></is></c></row><row r="2"><c r="A2"><v>10</v></c><c r="B2" t="inlineStr"><is><t>ten</t></is></c></row><row r="3" hidden="1"><c r="A3"><v>20</v></c><c r="B3" t="inlineStr"><is><t>twenty</t></is></c></row>'
    )
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected fallback');
  });
  try {
    const sorted = await book.apply('0', [
      { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:B3', column: 'A', direction: 'asc' },
    ]);
    let cells = (await book.read(sorted.view.storage.state, 'Sheet1', 'A1:B3')).cells;
    assert.deepEqual([cells.A1.value, cells.A2.value, cells.A3.value], [10, 20, 30]);
    assert.equal(cells.B3.value, 'thirty');
    assert.deepEqual(sorted.view.sheets[0].rowHeights, [{ row: 2, height: 29 }]);
    assert.deepEqual(sorted.view.sheets[0].hiddenRows, [1]);
    const bytes = await fs.readFile((await book.export(sorted.view.storage.state)).path),
      saved = unzipSync(bytes);
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<r><t>thirty<\/t><\/r>/);
    for (const name of Object.keys(parts))
      if (name !== 'xl/worksheets/sheet1.xml') assertRetainedPart(name, parts[name], saved[name]);
    const filtered = await book.apply(sorted.view.storage.state, [
      { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B3', column: 'A', values: [30] },
    ]);
    assert.deepEqual(filtered.view.sheets[0].hiddenRows, [1]);
    const cleared = await book.apply(filtered.view.storage.state, [
      { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B3', column: 'A', values: [] },
    ]);
    assert.deepEqual(cleared.view.sheets[0].hiddenRows, [1]);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 30);
    // A sort over fewer columns than the sheet uses moves only those columns, as in Excel.
    const narrow = await book.apply('0', [
        { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:A3', column: 'A', direction: 'asc' },
      ]),
      before = (await book.read('0', 'Sheet1', 'A1:B3')).cells,
      after = (await book.read(narrow.view.storage.state, 'Sheet1', 'A1:B3')).cells,
      column = (cells, letter) => [1, 2, 3].map((row) => cells[letter + row]?.value);
    assert.deepEqual(
      column(after, 'A'),
      column(before, 'A').sort((x, y) => x - y)
    );
    assert.deepEqual(column(after, 'B'), column(before, 'B'));
  } finally {
    await book.close();
  }
});
test('native unreferenced structural transforms move rows and columns with exact XML payloads', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>12345678901234567890</v></c><c r="B1" t="inlineStr"><is><r><t>rich</t></r></is></c></row><row r="2" ht="30" customHeight="1"><c r="B2"><v>5</v></c></row>'
    )
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected fallback');
  });
  try {
    const inserted = await book.apply('0', [{ operation: 'insertRow', sheet: 'Sheet1', index: 0, count: 2 }]);
    assert.equal((await book.read(inserted.view.storage.state, 'Sheet1', 'B4')).cells.B4.value, 5);
    const column = await book.apply(inserted.view.storage.state, [
      { operation: 'insertColumn', sheet: 'Sheet1', index: 0 },
    ]);
    assert.equal((await book.read(column.view.storage.state, 'Sheet1', 'C4')).cells.C4.value, 5);
    const deleted = await book.apply(column.view.storage.state, [
      { operation: 'deleteRow', sheet: 'Sheet1', index: 2 },
    ]);
    assert.equal((await book.read(deleted.view.storage.state, 'Sheet1', 'C3')).cells.C3.value, 5);
    assert.equal((await book.read(deleted.view.storage.state, 'Sheet1', 'B3')).cells.B3, undefined);
    const back = await book.apply(column.view.storage.state, [
      { operation: 'deleteColumn', sheet: 'Sheet1', index: 0 },
    ]);
    const saved = unzipSync(await fs.readFile((await book.export(back.view.storage.state)).path));
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /12345678901234567890/);
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<r><t>rich<\/t><\/r>/);
    assert.deepEqual(saved['custom/opaque.bin'], parts['custom/opaque.bin']);
  } finally {
    await book.close();
  }
});
test('compaction retains selected undo branches and removes abandoned native versions', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'));
  const book = await owner(source);
  try {
    let state = '0';
    const kept = [];
    for (let i = 0; i < 15; i++) {
      const result = await book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', value: i + 2 }]);
      state = result.view.storage.state;
      if (i === 4 || i === 14) kept.push(state);
    }
    const branch = await book.apply(kept[0], [{ sheet: 'Sheet1', cell: 'B1', value: 'branch' }]);
    kept.push(branch.view.storage.state);
    const result = await book.compact(kept);
    assert(result.reports.some((r) => r.afterVersions < r.beforeVersions));
    assert.equal((await book.read(kept[0], 'Sheet1', 'A1')).cells.A1.value, 6);
    assert.equal((await book.read(kept[1], 'Sheet1', 'A1')).cells.A1.value, 16);
    assert.equal((await book.read(kept[2], 'Sheet1', 'B1')).cells.B1.value, 'branch');
    await assert.rejects(book.read('2', 'Sheet1', 'A1'), /Unknown workbook/);
    const next = await book.apply(kept[1], [{ sheet: 'Sheet1', cell: 'A1', value: 99 }]);
    assert.equal(
      readWorkbook(await fs.readFile((await book.export(next.view.storage.state)).path)).sheets[0].cells.A1.value,
      99
    );
  } finally {
    await book.close();
  }
});
test('persistent index clones are revision-bound and cannot leak mutations between owners', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>7</v></c></row>'));
  const identity = `test-${Date.now()}-${Math.random()}`;
  const make = () =>
    new IndexedWorkbook(
      native,
      () => {
        throw Error('Unexpected fallback');
      },
      (input) => calculateIndexedWorkbook(input, native)
    );
  const first = make(),
    second = make();
  try {
    await first.open(source, { cacheIdentity: identity });
    assert.equal(first.cacheHit, false);
    await first.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 99 }]);
    await second.open(source, { cacheIdentity: identity });
    assert.equal(second.cacheHit, true);
    assert.equal((await second.read('0', 'Sheet1', 'A1')).cells.A1.value, 7);
  } finally {
    await first.close();
    await second.close();
  }
});
test('close cooperatively interrupts a native bulk operation without publishing a new state', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'));
  let started;
  const ready = new Promise((resolve) => {
    started = resolve;
  });
  const transport = (input) => {
    if (input.operation === 'bulk') started();
    return native(input);
  };
  const book = await owner(source, transport);
  const work = book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range: 'A1:Z1000000' }]);
  await ready;
  await new Promise((resolve) => setTimeout(resolve, 50));
  const start = performance.now();
  const rejected = assert.rejects(work, /cancel|interrupt|closed/i);
  await book.close();
  await rejected;
  // Cancellation is proved by the rejection above. The time limit only catches
  // a hang, so it is generous: the suite runs many files at once.
  assert(performance.now() - start < 30000);
  assert.equal(readWorkbook(await fs.readFile(source)).sheets[0].cells.A1.value, 1);
});
test('native structural guard rejects opaque reference-bearing worksheets before any writer', async () => {
  const { source } = await fixture(
    `<worksheet xmlns="${ns}" xmlns:x="urn:opaque"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><x:opaque ref="A1"/></sheetData></worksheet>`
  );
  let fallback = 0;
  const book = await owner(source, native, () => {
    fallback++;
    throw Error('Fallback reached');
  });
  try {
    await assert.rejects(book.apply('0', [{ operation: 'insertColumn', sheet: 'Sheet1', index: 0 }]), /not qualified/);
    assert.equal(fallback, 0);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 1);
  } finally {
    await book.close();
  }
});
test('native inserted rows inherit previous cell styles and retain exact payloads in moved rows', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>7</v></c><c r="B1" t="inlineStr"><is><t>label</t></is></c></row><row r="2"><c r="A2"><v>8</v></c></row>'
    )
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected fallback');
  });
  try {
    const formatted = await book.apply('0', [
      { type: 'formatRange', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true } },
    ]);
    const inserted = await book.apply(formatted.view.storage.state, [
      { operation: 'insertRow', sheet: 'Sheet1', index: 1, count: 2 },
    ]);
    const cells = (await book.read(inserted.view.storage.state, 'Sheet1', 'A1:B4')).cells;
    assert.equal(cells.A2.value, null);
    assert.equal(cells.B3.bold, true);
    assert.equal(cells.A4.value, 8);
  } finally {
    await book.close();
  }
});
test('dependency plan reuse invalidates on formula changes and restores an older branch correctly', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>A1*2</f></c><c r="C1"><f>B1+1</f></c></row>')
  );
  const book = await owner(source);
  try {
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'A1*3' }]);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 10);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 11);
  } finally {
    await book.close();
  }
});
test('persistent formula indexes reuse static results but recalculate volatile contexts', async () => {
  for (const [formula, expectedCalculations] of [
    ['1+2', 1],
    ['RAND()', 2],
  ]) {
    const { source } = await fixture(sheet(`<row r="1"><c r="A1"><f>${formula}</f></c></row>`));
    let calls = 0;
    const identity = `volatile-cache-${Date.now()}-${Math.random()}`;
    const make = () =>
      new IndexedWorkbook(
        native,
        () => {
          throw Error('Unexpected fallback');
        },
        (input) => {
          calls++;
          return calculateIndexedWorkbook(input, native);
        }
      );
    for (let i = 0; i < 2; i++) {
      const book = make();
      try {
        await book.open(source, { cacheIdentity: identity });
        assert.equal(book.cacheHit, i === 1);
        assert.equal(
          (await book.read('0', 'Sheet1', 'A1')).cells.A1.calculation.status,
          formula === 'RAND()' ? 'unresolved' : 'calculated'
        );
      } finally {
        await book.close();
      }
    }
    assert.equal(calls, expectedCalculations);
  }
});
test('indexed row insertion preserves formulas, names and table references without the whole-workbook fallback', async () => {
  const { spawnSync } = await import('node:child_process');
  for (const id of ['budget', 'pricing-schedule']) {
    const corpus = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-indexed-insertion-corpus-'));
    roots.push(corpus);
    const made = spawnSync(
      'python3',
      [fileURLToPath(new URL('../fixtures/build-corpus.py', import.meta.url)), '--output', corpus],
      { encoding: 'utf8' }
    );
    assert.equal(made.status, 0, made.stderr);
    let original = await fs.readFile(path.join(corpus, id + '.xlsx'));
    if (id === 'pricing-schedule') {
      const parts = unzipSync(original);
      parts['xl/workbook.xml'] = strToU8(
        strFromU8(parts['xl/workbook.xml']).replace(
          '</definedNames>',
          '<definedName name="SecondAmount">Pricing!$D$3</definedName></definedNames>'
        )
      );
      parts['xl/worksheets/sheet2.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet2.xml']).replace(
          '</sheetData>',
          '<row r="2"><c r="B2"><f>Pricing!D4</f></c><c r="C2"><f>SecondAmount</f></c></row></sheetData>'
        )
      );
      // The last three columns are calculated columns, as Excel records a
      // table column whose cells share one formula. A row inserted inside the
      // table continues those columns and no others.
      parts['xl/tables/table1.xml'] = strToU8(
        strFromU8(parts['xl/tables/table1.xml']).replace(
          /<tableColumn id="([456])" name="([^"]*)"\s*\/>/g,
          '<tableColumn id="$1" name="$2"><calculatedColumnFormula>0</calculatedColumnFormula></tableColumn>'
        )
      );
      assert.equal(strFromU8(parts['xl/tables/table1.xml']).match(/calculatedColumnFormula>/g)?.length, 6);
      original = zipSync(parts);
    }
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'indexed-insertion-'));
    roots.push(root);
    const source = path.join(root, 'source.xlsx');
    await fs.writeFile(source, original);
    const book = await owner(source, native, () => {
      throw Error('Unexpected whole-workbook fallback');
    });
    try {
      const sheetName = id === 'budget' ? 'Budget' : 'Pricing';
      assert.equal(book.view('0').sheets.find((s) => s.name === sheetName).capabilities.insertRows, true);
      const result = await book.apply('0', [{ sheet: sheetName, operation: 'insertRow', index: 2 }]);
      const edits = [
        { sheet: sheetName, cell: 'A3', value: 'New' },
        { sheet: sheetName, cell: 'B3', value: 4 },
        { sheet: sheetName, cell: 'C3', value: 25 },
        { sheet: sheetName, cell: 'D3', formula: 'B3*C3' },
      ];
      const changed = await book.apply(result.view.storage.state, edits);
      const saved = await book.export(changed.view.storage.state);
      const expected =
        id === 'budget'
          ? {
              'Budget!D2': { formula: '$B2*C$2', calculated: 100 },
              'Budget!D3': { formula: 'B3*C3', calculated: 100 },
              'Budget!D4': { formula: 'B4*C4', calculated: 150 },
              'Budget!D5': { formula: 'SUM(D2:D4)', calculated: 350 },
              'Budget!F2': { formula: 'Rates!B1*D5', calculated: 35 },
              'Budget!G2': { formula: 'AVERAGE(D2:D4)', calculated: 350 / 3 },
              'Budget!H2': { formula: 'MIN(D2:D4)', calculated: 100 },
              'Budget!I2': { formula: 'MAX(D2:D4)', calculated: 150 },
              'Budget!J2': { formula: 'COUNT(D2:D4)', calculated: 3 },
            }
          : {
              'Rates!B2': { formula: 'Pricing!D5', calculated: 430 },
              'Rates!C2': { formula: 'SecondAmount', calculated: 80 },
              'Pricing!D2': { formula: 'B2*C2', calculated: 250 },
              'Pricing!D3': { formula: 'B3*C3', calculated: 100 },
              'Pricing!D4': { formula: 'B4*C4', calculated: 80 },
              'Pricing!D5': { formula: 'SUM(D2:D4)', calculated: 430 },
              'Pricing!D6': { formula: 'SUM(PricingItems[Subtotal])', calculated: 430 },
              'Pricing!E3': { formula: 'TaxRate', calculated: 0.09 },
              'Pricing!E4': { formula: 'TaxRate', calculated: 0.09 },
              'Pricing!F3': { formula: 'D3*(1+E3)', calculated: 109 },
              'Pricing!F4': { formula: 'D4*(1+E4)', calculated: 87.2 },
              'Pricing!F5': { formula: 'SUM(F2:F4)', calculated: 468.7 },
            };
      const check = spawnSync(
        'python3',
        [fileURLToPath(new URL('../scripts/test/xlsx-package-check.py', import.meta.url))],
        {
          input: JSON.stringify({
            source,
            saved: saved.path,
            structural: { [sheetName]: { index: 2 } },
            targets: {
              [sheetName]: ['A3', 'B3', 'C3', 'D3', ...(id === 'pricing-schedule' ? ['E3', 'F3'] : [])],
              ...(id === 'pricing-schedule' ? { Rates: ['B2'] } : {}),
            },
            expected,
            ...(id === 'pricing-schedule'
              ? { expectedNames: { SecondAmount: 'Pricing!$D$4' }, expectedTables: { 'xl/tables/table1.xml': 'A1:F4' } }
              : {}),
          }),
          encoding: 'utf8',
        }
      );
      assert.equal(check.status, 0, check.stderr || check.stdout);
      const report = JSON.parse(check.stdout);
      assert.equal(report.passed, true, JSON.stringify(report.failures?.slice(0, 10)));
      const read = (await book.read(changed.view.storage.state, sheetName, 'D3:D6')).cells;
      assert.equal(read.D5.calculation.value, id === 'budget' ? 350 : 430);
      assert.equal(
        (await book.read('0', sheetName, 'D4')).cells.D4.formula,
        id === 'budget' ? 'SUM(D2:D3)' : 'SUM(D2:D3)'
      );
    } finally {
      await book.close();
    }
  }
});
test('native dependency selection avoids reading unrelated formulas and retains invalidation fences through compaction', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f></c><c r="C1"><f>B1+1</f></c><c r="D1"><f>7+8</f></c></row>'
    )
  );
  const calls = [];
  const transport = (input) => {
    calls.push(input);
    return native(input);
  };
  const book = await owner(source, transport);
  try {
    calls.length = 0;
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]);
    assert.equal(
      calls.some((c) => c.operation === 'range' && c.formulas),
      false
    );
    assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 11);
    const volatile = await book.apply(changed.view.storage.state, [{ sheet: 'Sheet1', cell: 'B1', formula: 'RAND()' }]);
    await book.compact(['0', volatile.view.storage.state]);
    calls.length = 0;
    const updated = await book.apply(volatile.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
    assert.equal(
      calls.some((c) => c.operation === 'range' && c.formulas),
      true
    );
    assert.equal(
      (await book.read(updated.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.status,
      'unresolved'
    );
  } finally {
    await book.close();
  }
});
test('storage-side structural planning preserves overlays, leaves new rows empty and undoes without cell projection', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 600 },
        (_, i) =>
          `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c><c r="B${i + 1}"><f>$A${i + 1}*2</f></c><c r="C${i + 1}"><f>B${i + 1}+1</f></c></row>`
      ).join('')
    )
  );
  const calls = [];
  const book = await owner(source, async (input) => {
    const result = await native(input);
    calls.push({ input, result });
    return result;
  });
  try {
    const edited = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'B300', formula: '$A300*3' },
      { sheet: 'Sheet1', cell: 'B310', value: 77 },
      { sheet: 'Sheet1', cell: 'D300', formula: 'B300+5' },
    ]);
    calls.length = 0;
    const moved = await book.apply(edited.view.storage.state, [
      { sheet: 'Sheet1', operation: 'insertRow', index: 300, count: 2 },
    ]);
    assert.equal(
      calls.some((c) => c.input.operation === 'range'),
      false,
      'Qualified planning must not fetch formula or predecessor cells into JS'
    );
    const plan = calls.find((c) => c.input.operation === 'plan-references');
    assert(plan.result.supported);
    assert.equal(plan.result.inputCells, 0);
    // Excel leaves inserted rows empty outside a table: no formula is continued into them.
    assert.equal(plan.result.formulaCount, 1200);
    const cells = (await book.read(moved.view.storage.state, 'Sheet1', 'A300:D303')).cells;
    assert.equal(cells.B300.calculation.value, 900);
    assert.equal(cells.B301?.formula, undefined);
    assert.equal(cells.D302?.formula, undefined);
    assert.equal(cells.B303.formula, '$A303*2');
    assert.equal(cells.B303.calculation.value, 602);
    assert.equal((await book.read(moved.view.storage.state, 'Sheet1', 'C312')).cells.C312.calculation.value, 78);
    assert.equal((await book.read(edited.view.storage.state, 'Sheet1', 'B301')).cells.B301.calculation.value, 602);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 599 }]);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'B599')).cells.B599.calculation.value, 1198);
    assert.equal((await book.read('0', 'Sheet1', 'B300')).cells.B300.calculation.value, 600);
    const saved = readWorkbook(await fs.readFile((await book.export(moved.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.D302?.formula, undefined);
    assert.equal(saved.sheets[0].cells.B312.value, 77);
  } finally {
    await book.close();
  }
});
test('whole-row structural references stay native and semantic misses preserve the spool', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>A1*2</f></c></row><row r="2"><c r="B2"><f>SUM(A:A)^2+SUM(3:3)</f></c></row>'
    )
  );
  let root;
  const calls = [];
  const book = await owner(source, async (input) => {
    root = input.root;
    const result = await native(input);
    calls.push({ input, result });
    return result;
  });
  try {
    const spool = path.join(root, 'insertion-formulas.jsonl');
    await fs.writeFile(spool, 'previous complete spool');
    const result = await native({
      action: 'workbook',
      operation: 'plan-references',
      root,
      version: 0,
      view: book.view('0'),
      command: { sheet: 'Sheet1', operation: 'insertRow', index: 1 },
    });
    assert.equal(result.supported, true);
    const complete = await fs.readFile(spool, 'utf8');
    const bad = await native({
      action: 'workbook',
      operation: 'patch',
      root,
      version: 0,
      changes: [{ sheet: 'xl/worksheets/sheet1.xml', cell: 'B2', value: { formula: 'SUM((A1,A2*2))' } }],
    });
    const miss = await native({
      action: 'workbook',
      operation: 'plan-references',
      root,
      version: bad.version,
      view: book.view('0'),
      command: { sheet: 'Sheet1', operation: 'insertRow', index: 1 },
    });
    assert.equal(miss.supported, false);
    assert.equal(await fs.readFile(spool, 'utf8'), complete);
    await assert.rejects(fs.stat(path.join(root, 'insertion-formulas.partial')), { code: 'ENOENT' });
    const boundary = await book.apply('0', [{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]);
    assert.equal((await book.read(boundary.view.storage.state, 'Sheet1', 'B3')).cells.B3.calculation.value, 9);
    const inserted = await book.apply('0', [{ sheet: 'Sheet1', operation: 'insertRow', index: 0 }]);
    assert(calls.some((c) => c.input.operation === 'plan-references' && c.result.supported));
    assert(!calls.some((c) => c.input.operation === 'range' && c.input.formulas));
    assert.equal((await book.read(inserted.view.storage.state, 'Sheet1', 'B3')).cells.B3.calculation.value, 9);
    assert.equal((await book.read('0', 'Sheet1', 'B2')).cells.B2.calculation.value, 9);
  } finally {
    await book.close();
  }
});
test('structural spool cancellation and quota failures do not publish a plan or mutate accepted versions', async () => {
  const { DatabaseSync } = await import('node:sqlite');
  const { source } = await fixture(
    sheet(
      Array.from({ length: 30000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f>${i + 1}+1</f></c></row>`).join('')
    )
  );
  let root;
  const book = await owner(source, async (input) => {
    root = input.root;
    return native(input);
  });
  const inventory = () => {
    const db = new DatabaseSync(path.join(root, 'cells.sqlite'));
    try {
      return ['versions', 'changes', 'dependency_plans', 'dependency_bindings'].map(
        (t) => db.prepare(`SELECT count(*) AS n FROM ${t}`).get().n
      );
    } finally {
      db.close();
    }
  };
  try {
    const before = inventory(),
      spool = path.join(root, 'insertion-formulas.jsonl');
    await fs.writeFile(spool, 'accepted sentinel');
    const request = {
      action: 'workbook',
      operation: 'plan-references',
      root,
      version: 0,
      view: book.view('0'),
      command: { sheet: 'Sheet1', operation: 'insertRow', index: 15000 },
    };
    let outcome;
    const work = native(request).then(
        (value) => {
          outcome = { value };
        },
        (error) => {
          outcome = { error };
        }
      ),
      deadline = Date.now() + 30000;
    while (
      !(await fs.stat(path.join(root, 'insertion-formulas.partial')).then(
        (s) => s.size > 0,
        () => false
      ))
    ) {
      assert(!outcome, 'Planner finished before cancellation checkpoint');
      assert(Date.now() < deadline);
      await new Promise((r) => setTimeout(r, 1));
    }
    await fs.writeFile(path.join(root, 'cancelled'), 'test');
    await work;
    assert.match(String(outcome.error), /cancel|interrupt/i);
    await fs.rm(path.join(root, 'cancelled'));
    assert.equal(await fs.readFile(spool, 'utf8'), 'accepted sentinel');
    await assert.rejects(fs.stat(path.join(root, 'insertion-formulas.partial')), { code: 'ENOENT' });
    assert.deepEqual(inventory(), before);
    const lease = path.join(path.dirname(root), 'owner.json'),
      original = await fs.readFile(lease, 'utf8');
    try {
      await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(original), maxSessionBytes: 128 * 1024 }));
      await assert.rejects(native(request), /budget/);
      await fs.writeFile(lease, original);
      const used = async (directory) => {
        let bytes = 0;
        for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
          const file = path.join(directory, entry.name);
          bytes += entry.isDirectory() ? await used(file) : (await fs.stat(file)).size;
        }
        return bytes;
      };
      const limit =
        (await used(path.dirname(root))) + 2 * (await fs.stat(path.join(root, 'cells.sqlite'))).size + 6 * 65536;
      await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(original), maxSessionBytes: limit }));
      await assert.rejects(native(request), /budget/);
      assert.equal(await fs.readFile(spool, 'utf8'), 'accepted sentinel');
      await assert.rejects(fs.stat(path.join(root, 'insertion-formulas.partial')), { code: 'ENOENT' });
    } finally {
      await fs.writeFile(lease, original);
    }
    assert.deepEqual(inventory(), before);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.calculation.value, 2);
  } finally {
    await fs.rm(path.join(root, 'cancelled'), { force: true });
    await book.close();
  }
});
test('disk admission refuses oversized growth and keeps the accepted version readable', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>7</v></c></row>'));
  const book = new IndexedWorkbook(
    native,
    () => {
      throw Error('Unexpected fallback');
    },
    (input) => calculateIndexedWorkbook(input, native)
  );
  try {
    await book.open(source, { maxSessionBytes: 1024 * 1024 });
    await assert.rejects(
      book.apply('0', [{ type: 'formatRange', sheet: 'Sheet1', range: 'A1:Z100000', format: { bold: true } }]),
      /disk|full|budget/i
    );
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 7);
  } finally {
    await book.close();
  }
});
test('native import and export cancel during streaming without changing the source', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 25000 },
        (_, i) =>
          `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="B${i + 1}" t="inlineStr"><is><t>${'preserved text '.repeat(12)}</t></is></c></row>`
      ).join('')
    )
  );
  const before = await fs.readFile(source);
  for (const phase of ['index', 'export']) {
    let started;
    const ready = new Promise((resolve) => {
      started = resolve;
    });
    const transport = (input) => {
      if (input.operation === phase) started(input.root);
      return native(input);
    };
    const book = new IndexedWorkbook(
      transport,
      () => {
        throw Error('Unexpected fallback');
      },
      (input) => calculateIndexedWorkbook(input, transport)
    );
    let work;
    if (phase === 'index') work = book.open(source);
    else {
      await book.open(source);
      const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
      work = book.export(changed.view.storage.state);
    }
    // Wait for the streaming artifact, not a fixed delay: optimized export can
    // legitimately finish within 20 ms, before the old test asked to cancel it.
    let outcome;
    const settled = work.then(
      (value) => {
        outcome = { value };
      },
      (error) => {
        outcome = { error };
      }
    );
    try {
      const root = await ready,
        artifact = path.join(root, phase === 'index' ? 'sheet-0.xml' : 'export.partial');
      const deadline = Date.now() + 30000;
      while (
        !(await fs.access(artifact).then(
          () => true,
          () => false
        ))
      ) {
        assert(!outcome, `${phase} finished before streaming was observed`);
        assert(Date.now() < deadline, `${phase} did not start streaming`);
        await new Promise((resolve) => setTimeout(resolve, 1));
      }
      const begin = performance.now();
      await book.close();
      await settled;
      assert.match(String(outcome.error), /cancel|interrupt|closed/i, phase);
      assert(performance.now() - begin < 30000, phase);
    } finally {
      await book.close();
      await settled;
    }
  }
  assert.deepEqual(await fs.readFile(source), before);
});
test('concurrent immutable cache publications respect entry and staging-inclusive byte limits', async () => {
  const { indexedCacheUsage, indexedCache } = await import('@office-editors/xlsx-editor/core/indexed-cache.mjs');
  const { root, source } = await fixture(sheet('<row r="1"><c r="A1"><v>1</v></c></row>'));
  for (let start = 0; start < 10; start += 5)
    await Promise.all(
      Array.from({ length: 5 }, async (_, i) => {
        const book = new IndexedWorkbook(
          native,
          () => {
            throw Error('Unexpected fallback');
          },
          (input) => calculateIndexedWorkbook(input, native)
        );
        try {
          await book.open(source, { cacheIdentity: `quota-cache-${Date.now()}-${start + i}` });
        } finally {
          await book.close();
        }
      })
    );
  const usage = await indexedCacheUsage();
  assert(usage.entries <= usage.maxEntries);
  assert(usage.bytes <= usage.maxBytes);
  assert.equal(usage.staging, 0);
  const oversized = path.join(root, 'oversized');
  await fs.mkdir(oversized);
  const file = await fs.open(path.join(oversized, 'sparse-index'), 'w');
  await file.truncate(usage.maxBytes + 1);
  await file.close();
  const cache = await indexedCache(source, `oversized-${Date.now()}`);
  await cache.publish(oversized, 0);
  const output = path.join(root, 'restore');
  await fs.mkdir(output);
  assert.equal(await cache.restore(output), null);
  assert((await indexedCacheUsage()).bytes <= usage.maxBytes);
});
test('closing during worker calculation keeps cancellation markers until the native lease releases', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><f>1+2</f></c></row>'));
  let ready, root;
  const started = new Promise((resolve) => {
    ready = resolve;
  });
  const book = new IndexedWorkbook(
    native,
    () => {
      throw Error('Unexpected fallback');
    },
    async (input, signal) => {
      root = input.root;
      await fs.writeFile(path.join(path.dirname(root), 'native-active.json'), JSON.stringify({ pid: process.pid }));
      ready();
      return new Promise((_, reject) =>
        signal.addEventListener('abort', () => reject(Error('Workbook operation cancelled')), { once: true })
      );
    }
  );
  const opening = book.open(source),
    rejected = assert.rejects(opening, /cancelled/);
  await started;
  const closing = book.close();
  await new Promise((resolve) => setTimeout(resolve, 30));
  assert((await fs.stat(path.join(root, 'cancelled'))).isFile());
  await fs.rm(path.join(path.dirname(root), 'native-active.json'));
  await closing;
  await rejected;
  await assert.rejects(fs.stat(root), { code: 'ENOENT' });
});
test('native compound total insertion saves correct references and caches while preserving the accepted base', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row><row r="3"><c r="B3"><f>SUM(A1:A2)+MAX(A1:A2)*2</f></c></row>'
    )
  );
  const workbook = strFromU8(parts['xl/workbook.xml']);
  assert(!workbook.includes('<definedNames'));
  parts['xl/workbook.xml'] = strToU8(
    workbook.replace(
      '</workbook>',
      '<definedNames><definedName name="Compound">SUM(Sheet1!$A$1:$A$2)+MAX(Sheet1!$A$1:$A$2)</definedName><definedName name="Scalar">Sheet1!$A$3</definedName></definedNames></workbook>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const calls = [];
  const book = await owner(
    source,
    (input) => {
      calls.push(input);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy mutation');
    }
  );
  try {
    const inserted = await book.apply('0', [{ sheet: 'Sheet1', operation: 'insertRow', index: 2 }]);
    const updated = await book.apply(inserted.view.storage.state, [{ sheet: 'Sheet1', cell: 'A3', value: 4 }]);
    const saved = await book.export(updated.view.storage.state),
      bytes = await fs.readFile(saved.path),
      view = readWorkbook(bytes);
    assert.equal(view.sheets[0].cells.B4.formula, 'SUM(A1:A3)+MAX(A1:A3)*2');
    assert.equal(view.sheets[0].cells.B4.cachedValue, 17);
    // A name is not stretched by a row inserted just below its range, as in Excel.
    assert.equal(
      view.namedRanges.find((n) => n.name === 'Compound').formula,
      'SUM(Sheet1!$A$1:$A$2)+MAX(Sheet1!$A$1:$A$2)'
    );
    assert.equal(view.namedRanges.find((n) => n.name === 'Scalar').formula, 'Sheet1!$A$4');
    assert.equal((await book.read('0', 'Sheet1', 'B3')).cells.B3.calculation.value, 11);
    // The engine plans the whole change itself (names, totals, references).
    assert(calls.some((c) => c.operation === 'plan-references'));
    assert.deepEqual(unzipSync(bytes)['custom/opaque.bin'], parts['custom/opaque.bin']);
    const reopened = await owner(saved.path);
    try {
      assert.equal((await reopened.read('0', 'Sheet1', 'B4')).cells.B4.calculation.value, 17);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
  }
});
test('SQLite calculation streams sparse million-row ranges and rolls back unsupported batches', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>0.1</v></c><c r="B1"><f>ROUND(SUM(A1:A1048576),2)</f></c><c r="C1"><f>IF(B1&gt;0,ABS(-0.2),1/0)</f></c></row><row r="1048576"><c r="A1048576"><v>0.2</v></c></row>'
    )
  );
  const calls = [],
    book = await owner(source, async (input) => {
      const reply = await native(input);
      calls.push({ input, reply });
      return reply;
    });
  try {
    const first = calls.find((c) => c.input.operation === 'calculate');
    assert.equal(first.reply.nativeSqlite, true);
    assert.equal(first.reply.inputCells, 0);
    assert.equal(first.reply.evaluatedFormulaCells, 2);
    const cells = (await book.read('0', 'Sheet1', 'B1:C1')).cells;
    assert.equal(cells.B1.calculation.value, 0.3);
    assert.equal(cells.C1.calculation.value, 0.2);
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 0.4 }]);
    assert.equal((await book.read(edited.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 0.6);
    const root = calls.find((c) => c.input.operation === 'calculate').input.root;
    const unsupported = await book.apply(edited.view.storage.state, [
      { sheet: 'Sheet1', cell: 'D1', formula: 'IF(TRUE,1,INDIRECT("R1C1",FALSE))' },
    ]);
    assert.equal((await book.read(unsupported.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 0.6);
    assert(calls.some((c) => c.input.operation === 'calculate' && c.reply.supported === false));
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 0.3);
    assert(root);
  } finally {
    await book.close();
  }
});
test('reference-aware row deletion shrinks ranges, moves formulas and refuses unsafe table boundaries', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM($A$1:A4)</f></c><c r="C1"><f>A4*2</f></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f>A3+1</f></c></row><row r="4"><c r="A4"><v>4</v></c></row>'
    )
  );
  const book = await owner(source);
  try {
    const result = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1 }]);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C4')).cells;
    assert.equal(cells.B1.formula, 'SUM($A$1:A3)');
    assert.equal(cells.B1.calculation.value, 8);
    assert.equal(cells.C1.formula, 'A3*2');
    assert.equal(cells.B2.formula, 'A2+1');
    assert.equal(cells.B2.calculation.value, 4);
    assert.equal(cells.A3.value, 4);
    assert(!cells.A4);
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B1.cachedValue, 8);
    assert.equal((await book.read('0', 'Sheet1', 'A2')).cells.A2.value, 2);
  } finally {
    await book.close();
  }
});
test('SQLite calculation cancellation and semantic misses leave no staged version or dependency plan', async () => {
  const { spawnSync } = await import('node:child_process');
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 12000 },
        (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>0.1</v></c><c r="B${i + 1}"><f>A${i + 1}+0.2</f></c></row>`
      ).join('')
    )
  );
  let request;
  const book = await owner(source, async (input) => {
    if (input.operation === 'calculate') request = structuredClone(input);
    return native(input);
  });
  const inventory = () => {
    const script =
      "import sqlite3,json,sys;d=sqlite3.connect(sys.argv[1]);tables={r[0] for r in d.execute('select name from sqlite_master')};print(json.dumps([d.execute('select count(*) from '+t).fetchone()[0] if t in tables else 0 for t in ['versions','changes','dependency_plans','dependency_bindings','shared_validations']]))";
    const r = spawnSync('python3', ['-c', script, path.join(request.root, 'cells.sqlite')], { encoding: 'utf8' });
    assert.equal(r.status, 0, r.stderr);
    return JSON.parse(r.stdout);
  };
  try {
    const before = inventory();
    const calculating = native(request);
    const rejected = assert.rejects(calculating, /cancel|interrupt/i);
    await new Promise((resolve) => setTimeout(resolve, 30));
    await fs.writeFile(path.join(request.root, 'cancelled'), 'cancel test');
    await rejected;
    await fs.rm(path.join(request.root, 'cancelled'));
    assert.deepEqual(inventory(), before);
    // The final formula fails after earlier results have been staged in SQLite.
    const patched = await native({
      action: 'workbook',
      operation: 'patch',
      root: request.root,
      version: request.version,
      revision: 'test',
      changes: [
        {
          sheet: request.view.sheets[0].path,
          cell: 'B12000',
          value: { value: null, formula: '1/0' },
          edit: { formula: '1/0' },
        },
      ],
    });
    const stagedBefore = inventory();
    const result = await native({ ...request, version: patched.version });
    assert.equal(result.supported, false);
    assert.deepEqual(inventory(), stagedBefore);
    const lease = path.join(path.dirname(request.root), 'owner.json'),
      originalLease = await fs.readFile(lease, 'utf8');
    try {
      await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(originalLease), maxSessionBytes: 128 * 1024 }));
      await assert.rejects(native(request), /budget/);
    } finally {
      await fs.writeFile(lease, originalLease);
    }
    assert.deepEqual(inventory(), stagedBefore);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 0.1 + 0.2);
  } finally {
    await book.close();
  }
});

test('native formula fill and paste keep compact replay, snapshot overlap, styles and replacement dependencies', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>$A1+A$1</f></c><c r="C1"><f>A1*10</f></c><c r="D1"><v>19</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>'
    )
  );
  let legacy = 0;
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      const result = await native(input);
      calls.push({ input, result });
      return result;
    },
    () => legacy++
  );
  try {
    const filled = await book.apply('0', [
      { type: 'formatRange', sheet: 'Sheet1', range: 'B1', format: { bold: true } },
      { type: 'fillRange', sheet: 'Sheet1', source: 'B1', range: 'B2:B1200' },
    ]);
    assert.equal(filled.edits.length, 2);
    assert.equal(legacy, 0);
    assert.equal((await book.read(filled.view.storage.state, 'Sheet1', 'B1200')).cells.B1200.formula, '$A1200+A$1');
    assert.equal((await book.read(filled.view.storage.state, 'Sheet1', 'B2')).cells.B2.calculation.value, 5);
    const pasted = await book.apply(filled.view.storage.state, [
      { type: 'pasteFormulas', sheet: 'Sheet1', source: 'B1:D1', range: 'C1:E1' },
    ]);
    const cells = (await book.read(pasted.view.storage.state, 'Sheet1', 'C1:E1')).cells;
    assert.equal(cells.C1.formula, '$A1+B$1');
    assert.equal(cells.D1.formula, 'B1*10');
    assert.equal(cells.E1, undefined);
    assert.equal(cells.C1.bold, false);
    assert(calls.some((c) => c.input.operation === 'bulk' && c.result.formulaTouched));
    const changed = await book.apply(pasted.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]);
    assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 100);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.value, 19);
    const saved = readWorkbook(await fs.readFile((await book.export(changed.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B1200.cachedValue, 5);
    assert.equal(saved.sheets[0].cells.B1200.bold, true);
    const replay = await owner(source);
    try {
      const repeated = await replay.apply('0', [
        ...filled.edits,
        ...pasted.edits,
        { sheet: 'Sheet1', cell: 'A1', value: 5 },
      ]);
      assert.equal((await replay.read(repeated.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 100);
    } finally {
      await replay.close();
    }
    await assert.rejects(book.apply('0', [{ type: 'pasteFormulas', sheet: 'Sheet1', source: 'A1', range: 'A2' }]), {
      code: /^(fill-bad-input|formula-bad-input)$/,
    });
    await assert.rejects(
      book.apply('0', [{ type: 'pasteFormulas', sheet: 'Sheet1', source: 'B1', range: 'B2:B3' }]),
      /matching/
    );
    assert.equal(legacy, 0);
  } finally {
    await book.close();
  }
});

test('native series uses accepted seeds, preserves compact replay and atomically refuses occupied destinations', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>0.1</v></c><c r="B1"><v>10</v></c><c r="C1"><v>7</v></c></row><row r="2"><c r="A2"><v>0.2</v></c></row>'
    )
  );
  const book = await owner(source, native, () => assert.fail('Series must remain native'));
  try {
    const command = { type: 'fillSeries', sheet: 'Sheet1', range: 'A1:A12000' },
      result = await book.apply('0', [command]);
    assert.deepEqual(result.edits, [command]);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A11999:A12000')).cells;
    assert.equal(cells.A12000.value, 1200);
    const branch = await book.apply('0', [{ type: 'fillSeries', sheet: 'Sheet1', range: 'B1:E1' }]);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'E1')).cells.E1.value, 1);
    await assert.rejects(
      book.apply(result.view.storage.state, [{ type: 'fillSeries', sheet: 'Sheet1', range: 'A1:A12000' }]),
      { code: /^(fill-conflict|workbook-conflict)$/ }
    );
    await assert.rejects(book.apply('0', [{ type: 'fillSeries', sheet: 'Sheet1', range: 'A1:C3' }]), {
      code: /^(fill-bad-input|workbook-bad-input)$/,
    });
    assert.equal((await book.read('0', 'Sheet1', 'A3')).cells.A3, undefined);
    const ties = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'J1', value: 123456789012342.5 },
      { sheet: 'Sheet1', cell: 'J2', value: 123456789012343.5 },
      { type: 'fillSeries', sheet: 'Sheet1', range: 'J1:J3' },
    ]);
    assert.equal(
      (await book.read(ties.view.storage.state, 'Sheet1', 'J3')).cells.J3.value,
      Number((123456789012344.5).toPrecision(15))
    );
    const sourceBytes = await fs.readFile(source),
      legacy = await editCalculatedWorkbookResult(sourceBytes, [
        { type: 'fillSeries', sheet: 'Sheet1', range: 'A1:A40' },
      ]);
    const expected = legacy.view.sheets[0].cells;
    const values = (await book.read(result.view.storage.state, 'Sheet1', 'A1:A40')).cells;
    for (let row = 1; row <= 40; row++) assert.equal(values[`A${row}`].value, expected[`A${row}`].value);
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.A12000.value, 1200);
  } finally {
    await book.close();
  }
});

test('mixed fallback keeps supported native closures in SQLite and recalculates dependent compatibility formulas', async () => {
  const rows = Array.from(
    { length: 10000 },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}"><v>1</v></c>${i === 0 ? '<c r="B1"><f>LEN(PROPER("aaa"))</f></c><c r="C1"><f>SUM(A1:A10000)</f></c><c r="D1"><f>C1+LEN(PROPER("aaa"))</f></c><c r="E1"><f>IFERROR(#REF!,8)</f></c>' : ''}</row>`
  ).join('');
  const { source } = await fixture(sheet(rows));
  const diagnostics = [];
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('Unexpected writer fallback'),
    async (input) => {
      const result = await calculateIndexedWorkbook(input, native);
      diagnostics.push(result);
      return result;
    }
  );
  try {
    await book.open(source);
    let cells = (await book.read('0', 'Sheet1', 'B1:E1')).cells;
    assert.equal(cells.B1.calculation.value, 3);
    assert.equal(cells.C1.calculation.value, 10000);
    assert.equal(cells.D1.calculation.value, 10003);
    assert.equal(cells.E1.calculation.value, 8);
    assert(diagnostics.at(-1).inputCells <= 3);
    assert.equal(diagnostics.at(-1).nativePartialCells, 2);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A9000', value: 4 }]);
    cells = (await book.read(changed.view.storage.state, 'Sheet1', 'B1:E1')).cells;
    assert.equal(cells.C1.calculation.value, 10003);
    assert.equal(cells.D1.calculation.value, 10006);
    assert.equal(cells.E1.calculation.value, 8);
    assert(diagnostics.at(-1).inputCells <= 3);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 10003);
    const saved = readWorkbook(await fs.readFile((await book.export(changed.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.C1.formula, 'SUM(A1:A10000)');
    assert.equal(saved.sheets[0].cells.C1.cachedValue, 10003);
    assert.equal(saved.sheets[0].cells.D1.cachedValue, 10006);
    assert.equal(saved.sheets[0].cells.E1.cachedValue, 8);
  } finally {
    await book.close();
  }
});

test('whole-column names and aggregates stay in SQLite and invalidate on previously blank distant inputs', async () => {
  const { source, parts } = await fixture(
    sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>SUM(A:A)</f></c><c r="C1"><f>SUM(AllNumbers)</f></c></row>')
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</sheets>',
      '</sheets><definedNames><definedName name="AllNumbers">Sheet1!$A:$A</definedName></definedNames>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const diagnostics = [];
  const book = new IndexedWorkbook(
    native,
    () => assert.fail(),
    async (input) => {
      const result = await calculateIndexedWorkbook(input, native);
      diagnostics.push(result);
      return result;
    }
  );
  try {
    await book.open(source);
    assert.equal(diagnostics.at(-1).nativeSqlite, true);
    assert.equal(diagnostics.at(-1).inputCells, 0);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A900000', value: 3 }]);
    const cells = (await book.read(changed.view.storage.state, 'Sheet1', 'B1:C1')).cells;
    assert.equal(cells.B1.calculation.value, 5);
    assert.equal(cells.C1.calculation.value, 5);
    assert.equal(diagnostics.at(-1).nativeSqlite, true);
    assert.equal(diagnostics.at(-1).inputCells, 0);
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 2);
  } finally {
    await book.close();
  }
});

test('native selection summary scans complete accepted ranges with bounded replies and branch isolation', async () => {
  const { source } = await fixture(
    sheet(Array.from({ length: 10000 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>1</v></c></row>`).join(''))
  );
  const traffic = [];
  const book = await owner(source, async (input) => {
    const output = await native(input);
    traffic.push({ operation: input.operation, bytes: JSON.stringify(output).length });
    return output;
  });
  try {
    traffic.length = 0;
    assert.deepEqual(await book.selectionSummary('0', 'Sheet1', 'A1:A10000'), {
      count: 10000,
      numericCount: 10000,
      sum: 10000,
      average: 1,
      min: 1,
      max: 1,
      errors: 0,
      unresolved: 0,
      unsafeNumbers: 0,
    });
    assert.deepEqual(
      traffic.map((t) => t.operation),
      ['selection']
    );
    assert(traffic[0].bytes < 512);
    const changed = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A9999', value: 5 },
      { sheet: 'Sheet1', cell: 'B1', formula: 'TRUNC(-3)' },
      { sheet: 'Sheet1', cell: 'B2', formula: 'DATE(1900,2,29)' },
      { sheet: 'Sheet1', cell: 'B3', value: true },
      { sheet: 'Sheet1', cell: 'B4', value: 'text' },
      { sheet: 'Sheet1', cell: 'B5', formula: '1/0' },
      { sheet: 'Sheet1', cell: 'B6', formula: 'UNKNOWN_FN(1)' },
    ]);
    const summary = await book.selectionSummary(changed.view.storage.state, 'Sheet1', 'A1:B10000');
    assert.equal(summary.count, 10006);
    assert.equal(summary.numericCount, 10002);
    assert.equal(summary.sum, 10061);
    assert.equal(summary.errors, 1);
    assert.equal(summary.unresolved, 1);
    assert.equal(summary.min, -3);
    assert.equal(summary.max, 60);
    assert.equal((await book.selectionSummary('0', 'Sheet1', 'A1:A10000')).sum, 10000);
    await assert.rejects(book.selectionSummary('0', 'Sheet1', 'A10:A1'), /Invalid range/);
    assert.deepEqual(await book.selectionSummary('0', 'Sheet1', 'B1:C3'), {
      count: 0,
      numericCount: 0,
      sum: 0,
      average: null,
      min: null,
      max: null,
      errors: 0,
      unresolved: 0,
      unsafeNumbers: 0,
    });
  } finally {
    await book.close();
  }
  const precise = await fixture(sheet('<row r="1"><c r="A1"><v>12345678901234567890</v></c></row>'));
  const special = await owner(precise.source);
  try {
    const r = await special.selectionSummary('0', 'Sheet1', 'A1');
    assert.equal(r.unsafeNumbers, 1);
    assert.equal(r.sum, null);
  } finally {
    await special.close();
  }
});

test('native whole-column formulas and names survive column insertion/deletion and export', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="D1"><f>SUM(A:B)</f></c><c r="E1"><f>SUM(Whole)</f></c></row>'
    )
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</sheets>',
      '</sheets><definedNames><definedName name="Whole">Sheet1!$A:$B</definedName></definedNames>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const traffic = [];
  const book = await owner(
    source,
    async (i) => {
      traffic.push(i);
      return native(i);
    },
    () => assert.fail('whole-column transformation must be native')
  );
  try {
    const next = await book.apply('0', [{ sheet: 'Sheet1', operation: 'insertColumn', index: 1 }]);
    const cells = (await book.read(next.view.storage.state, 'Sheet1', 'A1:F1')).cells;
    assert.equal(cells.E1.formula, 'SUM(A:C)');
    assert.equal(cells.F1.calculation.value, 5);
    assert.equal(next.view.namedRanges[0].formula, 'Sheet1!$A:$C');
    assert(!traffic.some((i) => i.operation === 'range' && i.formulas), 'no formula projection');
    const deleted = await book.apply(next.view.storage.state, [
      { sheet: 'Sheet1', operation: 'deleteColumn', index: 2 },
    ]);
    const out = readWorkbook(await fs.readFile((await book.export(deleted.view.storage.state)).path));
    assert.equal(out.sheets[0].cells.D1.formula, 'SUM(A:B)');
    assert.equal(out.sheets[0].cells.E1.cachedValue, 2);
    assert.equal(out.namedRanges[0].formula, 'Sheet1!$A:$B');
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 5);
  } finally {
    await book.close();
  }
});

test('native formula sort matches compatibility ordering, references, metadata and caches without cell projection', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1" ht="29" customHeight="1"><c r="A1"><v>3</v></c><c r="B1"><f>A1*2+$D$8</f></c><c r="C1" t="inlineStr"><is><r><t>three</t></r></is></c></row><row r="2" hidden="1"><c r="A2"><v>1</v></c><c r="B2"><f>A2*2+$D$8</f></c><c r="C2" t="inlineStr"><is><t>first</t></is></c></row><row r="3"><c r="A3"><v>1</v></c><c r="B3"><f>A3*2+$D$8</f></c><c r="C3" t="inlineStr"><is><t>second</t></is></c></row><row r="8"><c r="D8"><v>10</v></c><c r="E8"><f>D8*2</f></c></row>'
    )
  );
  const calls = [];
  let legacy = 0;
  const book = await owner(
    source,
    async (input) => {
      calls.push(input.operation);
      return native(input);
    },
    () => legacy++
  );
  try {
    for (const direction of ['asc', 'desc']) {
      calls.length = 0;
      const command = { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:C3', column: 'B', direction };
      const result = await book.apply('0', [command]);
      assert.equal(legacy, 0);
      assert(calls.includes('transform'));
      assert(!calls.includes('range'));
      const saved = await fs.readFile((await book.export(result.view.storage.state)).path),
        actual = readWorkbook(saved);
      const expected = (await editCalculatedWorkbookResult(await fs.readFile(source), [command])).view;
      const values = (v) =>
        Object.fromEntries(
          Object.entries(v.sheets[0].cells).map(([r, c]) => [
            r,
            { value: c.value, formula: c.formula, cachedValue: c.cachedValue },
          ])
        );
      assert.deepEqual(values(actual), values(expected));
      assert.deepEqual(actual.sheets[0].rowHeights, expected.sheets[0].rowHeights);
      assert.deepEqual(actual.sheets[0].hiddenRows, expected.sheets[0].hiddenRows);
      assert.equal(actual.sheets[0].cells[direction === 'asc' ? 'C1' : 'C2'].value, 'first');
      const zip = unzipSync(saved);
      assert.match(strFromU8(zip['xl/worksheets/sheet1.xml']), /<r><t>three<\/t><\/r>/);
      assert.deepEqual(zip['custom/opaque.bin'], parts['custom/opaque.bin']);
    }
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 3);
  } finally {
    await book.close();
  }
});

test('formula sort moves formulas as Excel does and rejects dynamic references atomically', async () => {
  // Relative references keep their distance from the formula's new row; absolute ones stay.
  for (const [formula, moved] of [
    ['SUM(A1:A2)', 'SUM(A2:A3)'],
    ['$A$1*2', '$A$1*2'],
  ]) {
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>${formula}</f></c></row><row r="2"><c r="A2"><v>1</v></c></row>`
      )
    );
    const book = await owner(source);
    try {
      const sorted = await book.apply('0', [
        { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:B2', column: 'A', direction: 'asc' },
      ]);
      const cells = (await book.read(sorted.view.storage.state, 'Sheet1', 'A1:B2')).cells;
      assert.equal(cells.A1.value, 1);
      assert.equal(cells.B2.formula, moved);
    } finally {
      await book.close();
    }
  }
  for (const formula of ['INDIRECT("A1")']) {
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>${formula}</f></c></row><row r="2"><c r="A2"><v>1</v></c></row>`
      )
    );
    const book = await owner(source);
    try {
      await assert.rejects(
        book.apply('0', [{ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:B2', column: 'A', direction: 'asc' }]),
        /relative references|safely update|Dynamic formula sorting is not qualified/
      );
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 2);
    } finally {
      await book.close();
    }
  }
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="9"><c r="B9"><f>A1*2</f></c></row>'
    )
  );
  const book = await owner(source);
  try {
    // A formula outside the sorted rows keeps its text and follows the new
    // content of the row it points at, as in Excel.
    const sorted = await book.apply('0', [
        { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:A2', column: 'A', direction: 'asc' },
      ]),
      after = (await book.read(sorted.view.storage.state, 'Sheet1', 'A1:B9')).cells;
    assert.equal(after.A1.value, 1);
    assert.equal(after.B9.formula, 'A1*2');
    assert.equal(after.B9.calculation.value, 2);
    assert.equal((await book.read('0', 'Sheet1', 'B9')).cells.B9.formula, 'A1*2');
  } finally {
    await book.close();
  }
});

test('native multirow deletion updates formulas, names and object ranges without projecting cells', async () => {
  const { source, parts } = await fixture(
    `<worksheet xmlns="${ns}"><sheetData>${Array.from({ length: 6 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c>${i === 0 ? '<c r="B1"><f>SUM(A1:A6)</f></c><c r="C1"><f>A3</f></c>' : ''}</row>`).join('')}</sheetData><mergeCells count="1"><mergeCell ref="D5:E6"/></mergeCells><dataValidations count="1"><dataValidation type="whole" sqref="A2:A6"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations><hyperlinks><hyperlink ref="F6" location="Sheet1!A6"/></hyperlinks></worksheet>`
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="Numbers">Sheet1!$A$1:$A$6</definedName></definedNames></workbook>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      const result = await native(input);
      calls.push({ input, result });
      return result;
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    calls.length = 0;
    const result = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1, count: 3 }]);
    assert(calls.some((c) => c.input.operation === 'plan-references' && c.result.supported));
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    const c = saved.sheets[0].cells;
    assert.deepEqual([c.A1.value, c.A2.value, c.A3.value], [1, 5, 6]);
    assert.equal(c.B1.formula, 'SUM(A1:A3)');
    assert.equal(c.B1.cachedValue, 12);
    assert.equal(c.C1.formula, '#REF!');
    assert.equal(saved.namedRanges[0].formula, 'Sheet1!$A$1:$A$3');
    assert.deepEqual(saved.sheets[0].merges, ['D2:E3']);
    const xml = strFromU8(
      unzipSync(await fs.readFile((await book.export(result.view.storage.state)).path))['xl/worksheets/sheet1.xml']
    );
    assert.match(xml, /sqref="A2:A3"/);
    assert.match(xml, /ref="F3" location="Sheet1!A3"/);
    assert.equal((await book.read('0', 'Sheet1', 'A3')).cells.A3.value, 3);
    await assert.rejects(
      book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 0, count: 0 }]),
      /count|range/
    );
    await assert.rejects(
      book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1048575, count: 2 }]),
      /count|boundary/
    );
  } finally {
    await book.close();
  }
});

test('native multicolumn deletion clips width spans and rewrites anchored names and formulas', async () => {
  const { source, parts } = await fixture(
    `<worksheet xmlns="${ns}"><cols><col min="1" max="5" width="22" customWidth="1"/><col min="6" max="6" width="33"/></cols><sheetData><row r="1">${[1, 2, 3, 4, 5, 6].map((n, i) => `<c r="${'ABCDEF'[i]}1"><v>${n}</v></c>`).join('')}<c r="G1"><f>SUM(A:F)</f></c></row></sheetData></worksheet>`
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected legacy writer');
  });
  try {
    const result = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteColumn', index: 1, count: 3 }]);
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.deepEqual(
      Object.values(saved.sheets[0].cells)
        .filter((c) => !c.formula)
        .map((c) => c.value),
      [1, 5, 6]
    );
    assert.equal(saved.sheets[0].cells.D1.formula, 'SUM(A:C)');
    assert.equal(saved.sheets[0].cells.D1.cachedValue, 12);
    assert.deepEqual(
      saved.sheets[0].widths.map((w) => [w.min, w.max, w.width]),
      [
        [1, 2, 22],
        [3, 3, 33],
      ]
    );
    const inserted = await book.apply('0', [{ sheet: 'Sheet1', operation: 'insertColumn', index: 1, count: 3 }]);
    assert.equal((await book.read(inserted.view.storage.state, 'Sheet1', 'J1')).cells.J1.formula, 'SUM(A:I)');
    const restored = await book.apply(inserted.view.storage.state, [
      { sheet: 'Sheet1', operation: 'deleteColumn', index: 1, count: 3 },
    ]);
    assert.equal((await book.read(restored.view.storage.state, 'Sheet1', 'G1')).cells.G1.calculation.value, 21);
    assert.deepEqual(
      unzipSync(await fs.readFile((await book.export(result.view.storage.state)).path))['custom/opaque.bin'],
      parts['custom/opaque.bin']
    );
  } finally {
    await book.close();
  }
});

test('native table filtering matches compatibility and multirow deletion guards table boundaries', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Key</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>' +
        [1, 2, 3, 4]
          .map(
            (n) =>
              `<row r="${n + 1}"><c r="A${n + 1}"><v>${n % 2}</v></c><c r="B${n + 1}"><v>${(n % 2) * 2}</v></c></row>`
          )
          .join('')
    )
  );
  await fs.writeFile(
    source,
    editWorkbook(
      editWorkbook(await fs.readFile(source), [
        { operation: 'createTable', sheet: 'Sheet1', range: 'A1:B5', tableName: 'Data' },
      ]),
      [2, 3, 4, 5].map((row) => ({ sheet: 'Sheet1', cell: `B${row}`, formula: `A${row}*2` }))
    )
  );
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      calls.push(input.operation);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    calls.length = 0;
    const command = { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B5', column: 'B', values: [2] };
    const result = await book.apply('0', [command]);
    assert(calls.includes('transform'));
    assert(!calls.includes('range'));
    const saved = await fs.readFile((await book.export(result.view.storage.state)).path),
      expected = await editCalculatedWorkbookResult(await fs.readFile(source), [command]);
    assert.deepEqual(readWorkbook(saved).sheets[0].hiddenRows, expected.view.sheets[0].hiddenRows);
    const parts = unzipSync(saved),
      prior = unzipSync(await fs.readFile(source));
    const table = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
    assert.match(strFromU8(parts[table]), /colId="1"/);
    assert.match(strFromU8(parts[table]), /val="2"/);
    assert(!strFromU8(parts['xl/worksheets/sheet1.xml']).includes('<autoFilter'));
    for (const name of Object.keys(prior))
      if (![table, 'xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(name))
        assert.deepEqual(parts[name], prior[name], name);
    const cleared = await book.apply(result.view.storage.state, [{ ...command, values: [] }]);
    assert.deepEqual(cleared.view.sheets[0].hiddenRows, []);
    const removed = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1, count: 2 }]);
    assert.equal(removed.view.tables[0].ref, 'A1:B3');
    await assert.rejects(
      book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 0, count: 2 }]),
      /table header/
    );
    const emptied = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1, count: 4 }]);
    assert.equal(emptied.view.tables[0].ref, 'A1:B2');
    const tableRemoved = await book.apply('0', [{ sheet: 'Sheet1', operation: 'deleteColumn', index: 0, count: 2 }]);
    assert.equal(tableRemoved.view.tables.length, 0);
    assert.equal(book.view('0').tables.length, 1);
    await assert.rejects(book.apply('0', [{ ...command, range: 'A2:B5' }]), /complete table/);
  } finally {
    await book.close();
  }
});

test('native merge and unmerge stream cells, preserve metadata and reject hidden content atomically', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><r><t>title</t></r></is></c></row><row r="20000"><c r="B20000"><v>9</v></c></row>'
    )
  );
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      calls.push(input.operation);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    calls.length = 0;
    const command = { operation: 'mergeCells', sheet: 'Sheet1', range: 'A1:C2' },
      merged = await book.apply('0', [command]);
    assert(calls.includes('layout'));
    assert(!calls.includes('range'));
    assert.deepEqual(merged.view.sheets[0].merges, ['A1:C2']);
    const saved = unzipSync(await fs.readFile((await book.export(merged.view.storage.state)).path));
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<r><t>title<\/t><\/r>/);
    for (const name of Object.keys(parts))
      if (name !== 'xl/worksheets/sheet1.xml') assertRetainedPart(name, parts[name], saved[name]);
    const unmerged = await book.apply(merged.view.storage.state, [{ ...command, operation: 'unmergeCells' }]);
    assert.deepEqual(unmerged.view.sheets[0].merges, []);
    assert.equal((await book.read(unmerged.view.storage.state, 'Sheet1', 'B20000')).cells.B20000.value, 9);
    await assert.rejects(book.apply('0', [{ ...command, range: 'A1:C20000' }]), /nonempty cell: B20000/);
    await assert.rejects(book.apply(merged.view.storage.state, [{ ...command, range: 'B2:D3' }]), /overlaps/);
    // A formula elsewhere on the sheet does not stop a merge; a formula in a
    // cell the merge would hide does.
    const formula = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D10', formula: '1+2' }]);
    assert.deepEqual((await book.apply(formula.view.storage.state, [command])).view.sheets[0].merges, ['A1:C2']);
    const hidden = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', formula: '1+2' }]);
    await assert.rejects(book.apply(hidden.view.storage.state, [command]), /nonempty cell: B2/);
    assert.deepEqual(book.view('0').sheets[0].merges, []);
  } finally {
    await book.close();
  }
});

test('formula sort cancellation and disk refusal leave accepted state and published artifact intact', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 12000 },
        (_, i) =>
          `<row r="${i + 1}"><c r="A${i + 1}"><v>${12000 - i}</v></c><c r="B${i + 1}"><f>A${i + 1}*2</f></c></row>`
      ).join('')
    )
  );
  let request;
  const book = await owner(
    source,
    async (input) => {
      if (input.operation === 'transform') request = structuredClone(input);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    const accepted = await book.apply('0', [
      { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:B12000', column: 'B', direction: 'asc' },
    ]);
    const published = path.join(request.root, 'transform.xlsx'),
      prior = await fs.readFile(published);
    const pending = native(request),
      refused = assert.rejects(pending, /cancel|interrupt/i);
    await new Promise((r) => setTimeout(r, 10));
    await fs.writeFile(path.join(request.root, 'cancelled'), 'test');
    await refused;
    await fs.rm(path.join(request.root, 'cancelled'));
    assert.deepEqual(await fs.readFile(published), prior);
    const lease = path.join(path.dirname(request.root), 'owner.json'),
      text = await fs.readFile(lease, 'utf8');
    await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(text), maxSessionBytes: 128 * 1024 }));
    try {
      await assert.rejects(native(request), /budget|disk|reserve/);
    } finally {
      await fs.writeFile(lease, text);
    }
    assert.deepEqual(await fs.readFile(published), prior);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 12000);
    assert.equal((await book.read(accepted.view.storage.state, 'Sheet1', 'A1')).cells.A1.value, 1);
  } finally {
    await book.close();
  }
});

test('mixed structural batches publish only the final revision and replay merge-and-center natively', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>Title</t></is></c></row>'));
  let fallback = 0;
  const book = await owner(source, native, () => fallback++);
  try {
    const merge = { operation: 'mergeCells', sheet: 'Sheet1', range: 'A1:C2' },
      center = { type: 'formatRange', sheet: 'Sheet1', range: 'A1', format: { horizontalAlign: 'center' } };
    const result = await book.apply('0', [merge, center]);
    assert.equal(fallback, 0);
    assert.equal(result.edits.length, 2);
    assert.deepEqual(result.view.sheets[0].merges, ['A1:C2']);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'A1')).cells.A1.horizontalAlign, 'center');
    assert.throws(() => book.view('1'), /Unknown/);
    const saved = await fs.readFile((await book.export(result.view.storage.state)).path),
      replay = await owner(source, native, () => {
        throw Error('Unexpected legacy writer');
      });
    try {
      const replayed = await replay.apply('0', result.edits);
      assert.deepEqual(await fs.readFile((await replay.export(replayed.view.storage.state)).path), saved);
    } finally {
      await replay.close();
    }
    await assert.rejects(
      book.apply('0', [merge, { sheet: 'Sheet1', cell: 'A1', expectedValue: 'wrong', value: 'new' }]),
      /Conflict|expected|guard/i
    );
    assert.deepEqual(book.view('0').sheets[0].merges, []);
    assert.throws(() => book.view('3'), /Unknown/);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 'Title');
  } finally {
    await book.close();
  }
});

test('two-key native formula sorts retain stable secondary order in one atomic batch', async () => {
  const { source } = await fixture(
    sheet(
      [
        [2, 2],
        [1, 2],
        [2, 1],
        [1, 1],
      ]
        .map(
          ([a, b], i) =>
            `<row r="${i + 1}"><c r="A${i + 1}"><v>${a}</v></c><c r="B${i + 1}"><v>${b}</v></c><c r="C${i + 1}"><f>A${i + 1}*10+B${i + 1}</f></c></row>`
        )
        .join('')
    )
  );
  const book = await owner(source, native, () => {
    throw Error('Unexpected fallback');
  });
  try {
    const result = await book.apply(
      '0',
      ['B', 'A'].map((column) => ({ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:C4', column, direction: 'asc' }))
    );
    assert.equal(result.edits.length, 2);
    assert.throws(() => book.view('1'), /Unknown/);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'A1:C4')).cells;
    assert.deepEqual(
      [1, 2, 3, 4].map((r) => cells[`C${r}`].calculation.value),
      [11, 12, 21, 22]
    );
    assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 22);
  } finally {
    await book.close();
  }
});

test('display search matches grid formatting with bounded pages, exact cursors and branch isolation', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>0.25</v></c><c r="B1"><f>1/0</f></c></row><row r="2"><c r="A2"><v>60</v></c></row>' +
        Array.from(
          { length: 9000 },
          (_, i) => `<row r="${i + 3}"><c r="A${i + 3}" t="inlineStr"><is><t>entry-${i}</t></is></c></row>`
        ).join('') +
        '<row r="10001"><c r="A10001"><v>1234.5</v></c></row>'
    )
  );
  await fs.writeFile(
    source,
    editWorkbook(await fs.readFile(source), [
      { sheet: 'Sheet1', cell: 'A1', numberFormat: 'percent' },
      { sheet: 'Sheet1', cell: 'A2', numberFormat: 'date' },
      { sheet: 'Sheet1', cell: 'A10001', numberFormat: 'currency' },
    ])
  );
  const book = await owner(source);
  try {
    for (const [text, ref] of [
      ['25.00%', 'A1'],
      ['1900-02-29', 'A2'],
      ['#DIV/0!', 'B1'],
      ['$1,234.50', 'A10001'],
    ]) {
      let cursor,
        found = [];
      do {
        const page = await book.query('0', 'Sheet1', text, 'display', cursor);
        assert(page.scanned <= 8192);
        found.push(...Object.keys(page.cells));
        cursor = page.nextCursor;
      } while (cursor);
      assert.deepEqual(found, [ref]);
    }
    let cursor;
    const found = [];
    do {
      const page = await book.query('0', 'Sheet1', 'entry-', 'display', cursor, { range: 'A3:A22', matchLimit: 3 });
      assert(page.matches <= 3);
      found.push(...Object.keys(page.cells));
      cursor = page.nextCursor;
    } while (cursor);
    assert.equal(found.length, 20);
    assert.equal(new Set(found).size, 20);
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 0.5 }]);
    assert.equal((await book.query(edited.view.storage.state, 'Sheet1', '50.00%', 'display')).matches, 1);
    assert.equal((await book.query('0', 'Sheet1', '50.00%', 'display')).matches, 0);
  } finally {
    await book.close();
  }
});

test('mixed-type sorts agree with shared Excel type ordering without projecting keys', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  const values = [2, 10, '15', null, true, '2', 1e21, '😀', 'é', '\uE000', false, -3, '-2', 2, '15'];
  const { source } = await fixture(sheet(''));
  await fs.writeFile(
    source,
    editWorkbook(
      await fs.readFile(source),
      values.flatMap((value, i) => [
        { sheet: 'Sheet1', cell: `A${i + 1}`, value },
        { sheet: 'Sheet1', cell: `B${i + 1}`, value: i },
      ])
    )
  );
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      calls.push(input);
      return native(input);
    },
    () => {
      throw Error('Unexpected whole-workbook fallback');
    }
  );
  try {
    for (const direction of ['asc', 'desc']) {
      calls.length = 0;
      const command = { operation: 'sortRows', sheet: 'Sheet1', range: `A1:B${values.length}`, column: 'A', direction },
        result = await book.apply('0', [command]);
      assert(calls.some((c) => c.operation === 'transform'));
      assert(!calls.some((c) => c.mixedOrder));
      assert(!calls.some((c) => c.operation === 'range'));
      const actual = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path)),
        expected = (await editCalculatedWorkbookResult(await fs.readFile(source), [command])).view;
      assert.deepEqual(
        Array.from({ length: values.length }, (_, i) => actual.sheets[0].cells[`B${i + 1}`].value),
        Array.from({ length: values.length }, (_, i) => expected.sheets[0].cells[`B${i + 1}`].value)
      );
    }
  } finally {
    await book.close();
  }
});

test('native sheet and table metadata edits preserve literal payloads and package relationships', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Key</t></is></c><c r="B1" t="inlineStr"><is><t>Price</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><r><t>rich</t></r></is></c><c r="B2"><v>12345678901234567890</v></c></row><row r="10000"><c r="A10000"><v>7</v></c></row>'
    )
  );
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      calls.push(input);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    let result = await book.apply('0', [
      { operation: 'duplicateSheet', sheet: 'Sheet1', newName: 'Copy' },
      { operation: 'addSheet', sheet: 'Blank' },
      { operation: 'renameSheet', sheet: 'Copy', newName: 'Renamed' },
    ]);
    assert.equal(result.view.sheets.length, 3);
    assert(!calls.some((c) => c.operation === 'range'));
    let saved = unzipSync(await fs.readFile((await book.export(result.view.storage.state)).path));
    assertRetainedPart(
      'xl/worksheets/sheet1.xml',
      parts['xl/worksheets/sheet1.xml'],
      saved['xl/worksheets/sheet1.xml']
    );
    assertRetainedPart(
      result.view.sheets.find((s) => s.name === 'Renamed').path,
      parts['xl/worksheets/sheet1.xml'],
      saved[result.view.sheets.find((s) => s.name === 'Renamed').path]
    );
    calls.length = 0;
    result = await book.apply(result.view.storage.state, [
      { operation: 'createTable', sheet: 'Sheet1', range: 'A1:B2', tableName: 'Prices' },
      {
        operation: 'addConditionalRule',
        sheet: 'Sheet1',
        range: 'A2:B2',
        rule: 'greaterThan',
        threshold: 5,
        fillColor: 'FF0000',
      },
    ]);
    assert(calls.some((c) => c.operation === 'package-edit'));
    assert(calls.filter((c) => c.operation === 'range').every((c) => c.range === 'A1:B1'));
    saved = unzipSync(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<v>12345678901234567890<\/v>/);
    assert.match(strFromU8(saved['xl/worksheets/sheet1.xml']), /<r><t>rich<\/t><\/r>/);
    assert.deepEqual(saved['custom/opaque.bin'], parts['custom/opaque.bin']);
    assert(result.view.tables.some((t) => t.name === 'Prices'));
    assert(Object.keys(saved).some((n) => n === 'xl/worksheets/_rels/sheet1.xml.rels'));
    const formulas = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', formula: '1+2' }]);
    await assert.rejects(
      book.apply(formulas.view.storage.state, [
        { operation: 'createTable', sheet: 'Sheet1', range: 'A1:B2', tableName: 'Refused' },
      ]),
      /over formulas/
    );
    assert.equal(book.view('0').tables.length, 0);
  } finally {
    await book.close();
  }
});

test('table-internal columns update headers, ids, selectors, names and calculated-column metadata', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  const { source } = await fixture(sheet(''));
  let bytes = editWorkbook(
    await fs.readFile(source),
    [
      ['A1', 'Qty'],
      ['B1', 'Price'],
      ['C1', 'Total'],
      ['A2', 2],
      ['B2', 3],
      ['C2', 6],
      ['A3', 4],
      ['B3', 5],
      ['C3', 20],
    ].map(([cell, value]) => ({ sheet: 'Sheet1', cell, value }))
  );
  bytes = editWorkbook(bytes, [{ operation: 'createTable', sheet: 'Sheet1', range: 'A1:C3', tableName: 'Sales' }]);
  bytes = editWorkbook(bytes, [
    { sheet: 'Sheet1', cell: 'E1', formula: 'SUM(Sales[Price])' },
    { sheet: 'Sheet1', cell: 'F1', formula: 'SUM(Sales[Qty])' },
  ]);
  const parts = unzipSync(bytes),
    table = Object.keys(parts).find((p) => p.startsWith('xl/tables/'));
  parts[table] = strToU8(
    strFromU8(parts[table]).replace(
      '<tableColumn id="3" name="Total"/>',
      '<tableColumn id="3" name="Total"><calculatedColumnFormula>[@Qty]*[@Price]</calculatedColumnFormula></tableColumn>'
    )
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="PriceRange">Sales[Price]</definedName></definedNames></workbook>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const book = await owner(source, native, () => {
    throw Error('Unexpected legacy writer');
  });
  try {
    const inserted = await book.apply('0', [{ operation: 'insertColumn', sheet: 'Sheet1', index: 1, count: 2 }]);
    assert.deepEqual(
      inserted.view.tables[0].columns.map((c) => c.name),
      ['Qty', 'Column1', 'Column2', 'Price', 'Total']
    );
    let saved = readWorkbook(await fs.readFile((await book.export(inserted.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B1.value, 'Column1');
    assert.equal(saved.sheets[0].cells.C1.value, 'Column2');
    assert.equal(saved.sheets[0].cells.G1.cachedValue, 8);
    assert.equal(saved.tables[0].ref, 'A1:E3');
    const deleted = await book.apply('0', [{ operation: 'deleteColumn', sheet: 'Sheet1', index: 1 }]);
    saved = readWorkbook(await fs.readFile((await book.export(deleted.view.storage.state)).path));
    assert.equal(saved.tables[0].ref, 'A1:B3');
    assert.deepEqual(
      saved.tables[0].columns.map((c) => c.name),
      ['Qty', 'Total']
    );
    assert.equal(saved.sheets[0].cells.D1.formula, 'SUM(#REF!)');
    assert.equal(saved.sheets[0].cells.E1.cachedValue, 6);
    assert.equal(saved.namedRanges[0].formula, '#REF!');
    const zip = unzipSync(await fs.readFile((await book.export(deleted.view.storage.state)).path));
    assert.match(strFromU8(zip[table]), /id="3" name="Total"/);
    assert.match(strFromU8(zip[table]), /<calculatedColumnFormula>\[@Qty\]\*#REF!<\/calculatedColumnFormula>/);
    assert.equal((await book.read('0', 'Sheet1', 'B2')).cells.B2.value, 3);
  } finally {
    await book.close();
  }
});

test('native sheet rename updates cross-sheet formulas, scoped names and validation references', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>Sheet1!A1*2</f></c></row>'));
  let bytes = editWorkbook(await fs.readFile(source), [{ operation: 'addSheet', sheet: 'Other' }]);
  bytes = editWorkbook(bytes, [{ sheet: 'Other', cell: 'A1', formula: 'Sheet1!B1+1' }]);
  const parts = unzipSync(bytes);
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="Values">Sheet1!$A$1</definedName><definedName name="Local" localSheetId="0">Sheet1!$A$1</definedName></definedNames></workbook>'
    )
  );
  parts['xl/worksheets/sheet2.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet2.xml']).replace(
      '</worksheet>',
      '<dataValidations count="1"><dataValidation type="whole" sqref="B2"><formula1>Sheet1!$A$1</formula1></dataValidation></dataValidations></worksheet>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const calls = [];
  const book = await owner(
    source,
    async (input) => {
      calls.push(input.operation);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    calls.length = 0;
    const result = await book.apply('0', [{ operation: 'renameSheet', sheet: 'Sheet1', newName: "O'Brien Budget" }]);
    assert(!calls.includes('range'));
    const saved = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(saved.sheets[0].name, "O'Brien Budget");
    assert.equal(saved.sheets[0].cells.B1.formula, "'O''Brien Budget'!A1*2");
    assert.equal(saved.sheets[1].cells.A1.cachedValue, 5);
    assert(saved.namedRanges.every((n) => n.formula === "'O''Brien Budget'!$A$1"));
    const xml = strFromU8(
      unzipSync(await fs.readFile((await book.export(result.view.storage.state)).path))['xl/worksheets/sheet2.xml']
    );
    assert.match(xml, /<formula1>'O''Brien Budget'!\$A\$1<\/formula1>/);
    assert.equal(book.view('0').sheets[0].name, 'Sheet1');
  } finally {
    await book.close();
  }
});

test('package edits preserve ZIP directories and published state on cancellation or disk refusal', async () => {
  const { source, parts } = await fixture(
    sheet(
      Array.from({ length: 16000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f>${i + 1}+1</f></c></row>`).join('')
    )
  );
  parts['custom/'] = new Uint8Array();
  await fs.writeFile(source, zipSync(parts));
  let request;
  const book = await owner(
    source,
    async (input) => {
      if (input.operation === 'package-edit') request = structuredClone(input);
      return native(input);
    },
    () => {
      throw Error('Unexpected legacy writer');
    }
  );
  try {
    const accepted = await book.apply('0', [{ operation: 'renameSheet', sheet: 'Sheet1', newName: 'Renamed' }]);
    const published = path.join(request.root, 'package.xlsx'),
      prior = await fs.readFile(published);
    assert('custom/' in unzipSync(prior));
    const pending = native(request),
      refused = assert.rejects(pending, /cancel|interrupt/i);
    await new Promise((r) => setTimeout(r, 10));
    await fs.writeFile(path.join(request.root, 'cancelled'), 'test');
    await refused;
    await fs.rm(path.join(request.root, 'cancelled'));
    assert.deepEqual(await fs.readFile(published), prior);
    const lease = path.join(path.dirname(request.root), 'owner.json'),
      text = await fs.readFile(lease, 'utf8');
    await fs.writeFile(lease, JSON.stringify({ ...JSON.parse(text), maxSessionBytes: 128 * 1024 }));
    try {
      await assert.rejects(native(request), /budget|disk|reserve/);
    } finally {
      await fs.writeFile(lease, text);
    }
    assert.deepEqual(await fs.readFile(published), prior);
    assert.equal(book.view('0').sheets[0].name, 'Sheet1');
    assert.equal(
      (await book.read(accepted.view.storage.state, 'Renamed', 'A16000')).cells.A16000.calculation.value,
      16001
    );
  } finally {
    await fs.rm(path.join(request.root, 'cancelled'), { force: true });
    await book.close();
  }
});

test('mixed dependency plans select only affected compatibility formulas and preserve branches', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 600 },
        (_, i) =>
          `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>abc</t></is></c><c r="B${i + 1}"><f>LEN(A${i + 1})</f></c><c r="C${i + 1}"><f>B${i + 1}+1</f></c></row>`
      ).join('')
    )
  );
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
    const result = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A300', value: 'abcdef' }]);
    assert.equal(diagnostic.nativePlanHit, true);
    assert.equal(diagnostic.evaluatedFormulaCells, 2);
    assert(diagnostic.inputCells <= 3);
    let cells = (await book.read(result.view.storage.state, 'Sheet1', 'B300:C300')).cells;
    assert.equal(cells.B300.calculation.value, 6);
    assert.equal(cells.C300.calculation.value, 7);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'C301')).cells.C301.calculation.value, 4);
    assert.equal((await book.read('0', 'Sheet1', 'C300')).cells.C300.calculation.value, 4);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A301', value: 'x' }]);
    assert.equal(diagnostic.nativePlanHit, true);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'C301')).cells.C301.calculation.value, 2);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'C301')).cells.C301.calculation.value, 4);
    const unaffected = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'D1', value: 7 }]);
    assert(diagnostic.skipped);
    assert.equal((await book.read(unaffected.view.storage.state, 'Sheet1', 'C300')).cells.C300.calculation.value, 7);
  } finally {
    await book.close();
  }
});

test('nonzero-digit TRUNC mixed calculation survives edits, export and reopen', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>-2.7</v></c><c r="B1"><f>TRUNC(A1,1)</f></c><c r="C1"><f>B1*2</f></c></row>')
  );
  const book = await owner(source, native, () => assert.fail('writer fallback'));
  try {
    assert.equal((await book.read('0', 'Sheet1', 'B1:C1')).cells.B1.calculation.value, -2.7);
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: -3.9 }]),
      state = edited.view.storage.state;
    const cells = (await book.read(state, 'Sheet1', 'B1:C1')).cells;
    assert.equal(cells.B1.calculation.value, -3.9);
    assert.equal(cells.C1.calculation.value, -7.8);
    assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, -5.4);
    const saved = await book.export(state),
      reparsed = readWorkbook(await fs.readFile(saved.path));
    assert.equal(reparsed.sheets[0].cells.B1.cachedValue, -3.9);
    assert.equal(reparsed.sheets[0].cells.C1.cachedValue, -7.8);
    const reopened = await owner(saved.path);
    try {
      assert.equal((await reopened.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, -7.8);
    } finally {
      await reopened.close();
    }
  } finally {
    await book.close();
  }
});

test('dense compatibility recalculation pages selected formulas without per-cell native requests', async () => {
  const count = 2200;
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: count },
        (_, i) =>
          `<row r="${i + 1}">${i === 0 ? '<c r="A1"><v>1</v></c>' : ''}<c r="B${i + 1}"><f>DEC2HEX($A$1+${i + 1},8)</f></c><c r="C${i + 1}"><f>B${i + 1}</f></c></row>`
      ).join('')
    )
  );
  let ranges = [];
  const transport = async (input) => {
    if (input.operation === 'range') ranges.push(input);
    return native(input);
  };
  const book = await owner(source, transport, () => assert.fail('writer fallback'));
  try {
    ranges = [];
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 2 }]);
    assert(
      ranges.filter((request) => request.formulas).length < 10,
      'Selected formulas should use bounded pages, not one request per cell'
    );
    assert(
      ranges.length < 100,
      `Dense compatibility recalculation made ${ranges.length} range requests including input tiles`
    );
    const cells = (await book.read(next.view.storage.state, 'Sheet1', `B1:C${count}`)).cells;
    for (let row = 1; row <= count; row++)
      for (const col of ['B', 'C'])
        assert.equal(cells[col + row].calculation.value, (row + 2).toString(16).toUpperCase().padStart(8, '0'));
    assert.equal(
      (await book.read('0', 'Sheet1', `C${count}`)).cells[`C${count}`].calculation.value,
      (count + 1).toString(16).toUpperCase().padStart(8, '0')
    );
    const saved = readWorkbook(await fs.readFile((await book.export(next.view.storage.state)).path));
    assert.equal(
      saved.sheets[0].cells[`C${count}`].cachedValue,
      (count + 2).toString(16).toUpperCase().padStart(8, '0')
    );
  } finally {
    await book.close();
  }
});

test('sparse compatibility selections exclude unaffected formulas across row and column gaps', async () => {
  const { source } = await fixture(
    sheet(
      Array.from({ length: 12 }, (_, i) => {
        const row = i + 1;
        return `<row r="${row}">${row === 1 ? '<c r="A1"><v>1</v></c>' : row === 2 ? '<c r="A2"><v>8</v></c>' : ''}<c r="B${row}"><f>DEC2HEX($A$${row % 3 === 0 ? 2 : 1},8)</f></c><c r="C${row}"><f>DEC2HEX($A$2,8)</f></c><c r="D${row}"><f>DEC2HEX($A$${row % 3 === 0 ? 2 : 1},8)</f></c></row>`;
      }).join('')
    )
  );
  const selectedReads = [];
  let recording = false;
  const transport = async (input) => {
    const result = await native(input);
    if (recording && input.operation === 'range' && input.formulas)
      selectedReads.push(...result.cells.map(([ref]) => ref));
    return result;
  };
  const book = await owner(source, transport, () => assert.fail('writer fallback'));
  try {
    recording = true;
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 15 }]);
    recording = false;
    const expected = Array.from({ length: 12 }, (_, i) => i + 1)
      .filter((row) => row % 3 !== 0)
      .flatMap((row) => ['B' + row, 'D' + row]);
    assert.deepEqual(
      selectedReads.toSorted(),
      expected.toSorted(),
      'Formula fetches must cover only the selected cells, without holes or duplicates'
    );
    const cells = (await book.read(next.view.storage.state, 'Sheet1', 'B1:D12')).cells;
    for (let row = 1; row <= 12; row++) {
      assert.equal(cells['C' + row].calculation.value, '00000008');
      for (const col of ['B', 'D'])
        assert.equal(cells[col + row].calculation.value, row % 3 === 0 ? '00000008' : '0000000F');
    }
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, '00000001');
  } finally {
    await book.close();
  }
});

test('compatibility input budget rejects atomically without publishing partial calculation', async () => {
  const { source } = await fixture(
    sheet(
      Array.from(
        { length: 100 },
        (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>abc</t></is></c></row>`
      ).join('')
    )
  );
  let limited = false;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    (input) => calculateIndexedWorkbook({ ...input, ...(limited ? { maxProjectionBytes: 256 } : {}) }, native)
  );
  try {
    await book.open(source);
    limited = true;
    await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'CONCAT(A1:A100)' }]), {
      code: /^(csv-limit|formula-limit|session-limit)$/,
    });
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1, undefined);
    limited = false;
    const ok = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'LEN(A1)' }]);
    assert.equal((await book.read(ok.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 3);
  } finally {
    await book.close();
  }
});

test('array constants, empty calls and dynamic-reference copying avoid the whole-workbook writer', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>SUM({1,2})+A1</f></c><c r="C1"><f>SUM()</f></c><c r="D1"><f>INDIRECT("A1")+A1</f></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>SUM({1,2})+A2</f></c><c r="C2"><f>SUM()</f></c></row>'
    )
  );
  const book = await owner(source, native, () => assert.fail('whole-workbook writer'));
  try {
    const filled = await book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'B1:D1', range: 'B3:D4' }]);
    const cells = (await book.read(filled.view.storage.state, 'Sheet1', 'B3:D4')).cells;
    assert.equal(cells.B3.formula, 'SUM({1,2})+A3');
    assert.equal(cells.C4.formula, 'SUM()');
    assert.equal(cells.D4.formula, 'INDIRECT("A1")+A4');
    assert.equal(cells.D4.calculation.value, 2);
    const withoutDynamic = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D1', value: null }]);
    const sorted = await book.apply(withoutDynamic.view.storage.state, [
      { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:D2', column: 'A', direction: 'asc' },
    ]);
    assert.equal((await book.read(sorted.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 4);
  } finally {
    await book.close();
  }
});

test('shared formula metadata stays opaque while qualified calculation refreshes only caches', async () => {
  const { source, parts } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">B1+1</f><v>999</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>999</v></c></row>'
    )
  );
  const book = await owner(source, native, () => assert.fail('whole-workbook writer'));
  try {
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 2 }]);
    const output = unzipSync(await fs.readFile((await book.export(edited.view.storage.state)).path));
    const xml = strFromU8(output['xl/worksheets/sheet1.xml']);
    for (const ref of ['A1', 'A2']) {
      const pattern = new RegExp(`<c r="${ref}">[\\s\\S]*?<\\/c>`);
      assert.equal(
        xml.match(pattern)[0].replace(/<v\b[^>]*>[\s\S]*?<\/v>/g, ''),
        strFromU8(parts['xl/worksheets/sheet1.xml'])
          .match(pattern)[0]
          .replace(/<v\b[^>]*>[\s\S]*?<\/v>/g, '')
      );
    }
    assert.equal(
      readWorkbook(await fs.readFile((await book.export(edited.view.storage.state)).path)).sheets[0].cells.B1.value,
      2
    );
    assert.equal((await book.read(edited.view.storage.state, 'Sheet1', 'A1')).cells.A1.calculation.value, 3);
    assert.doesNotMatch(strFromU8(output['xl/workbook.xml']), /fullCalcOnLoad="1"/);
    const partial = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 3 }]);
    assert.equal((await book.read(partial.view.storage.state, 'Sheet1', 'A2')).cells.A2.formula, 'B2+1');
  } finally {
    await book.close();
  }
});

test('long reference lists copy natively while unqualified unions still reject atomically', async () => {
  const longFormula = 'SUM(' + Array(300).fill('A1').join(',') + ')';
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>' +
        longFormula +
        '</f></c></row><row r="100000"><c r="A100000"><v>12345678901234567890</v></c></row>'
    )
  );
  const calls = [];
  const book = await owner(
    source,
    async (i) => {
      calls.push(i);
      return native(i);
    },
    () => assert.fail('whole-workbook writer')
  );
  try {
    const command = { type: 'fillRange', sheet: 'Sheet1', source: 'B1', range: 'B2:B5' },
      r = await book.apply('0', [command]);
    assert.deepEqual(r.edits, [command]);
    assert.equal(
      (await book.read(r.view.storage.state, 'Sheet1', 'B5')).cells.B5.formula,
      longFormula.replaceAll('A1', 'A5')
    );
    assert(calls.some((i) => i.operation === 'bulk'));
    assert(!calls.some((i) => i.operation === 'bulk' && i.formulaOverrides));
    const bytes = await fs.readFile((await book.export(r.view.storage.state)).path);
    assert.match(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']), /12345678901234567890/);
    const replay = await book.apply('0', r.edits);
    assert.equal(
      (await book.read(replay.view.storage.state, 'Sheet1', 'B5')).cells.B5.formula,
      longFormula.replaceAll('A1', 'A5')
    );
    const union = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'SUM((A1,A1*2))' }]);
    await assert.rejects(book.apply(union.view.storage.state, [command]), /not qualified for native preservation/);
    assert.equal((await book.read(union.view.storage.state, 'Sheet1', 'B2')).cells.B2, undefined);
  } finally {
    await book.close();
  }
});

test('sort guards and malformed worksheet edits reject before compatibility loading', async () => {
  const { editWorkbook } = await import('@office-editors/xlsx-editor/core/workbook-package.mjs');
  let { source } = await fixture(
    sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>Key</t></is></c></row><row r="2"><c r="A2"><v>2</v></c></row>')
  );
  await fs.writeFile(
    source,
    editWorkbook(await fs.readFile(source), [
      { operation: 'createTable', sheet: 'Sheet1', range: 'A1:A2', tableName: 'Data' },
    ])
  );
  let book = await owner(source, native, () => assert.fail('whole-workbook writer'));
  try {
    await assert.rejects(
      book.apply('0', [{ operation: 'sortRows', sheet: 'Sheet1', range: 'A1:A2', column: 'A', direction: 'asc' }]),
      { code: 'table-or-protection' }
    );
    assert.equal((await book.read('0', 'Sheet1', 'A2')).cells.A2.value, 2);
  } finally {
    await book.close();
  }
  ({ source } = await fixture(`<worksheet xmlns="${ns}"/>`));
  book = await owner(source, native, () => assert.fail('whole-workbook writer'));
  try {
    await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 3 }]), /missing required sheetData/);
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1, undefined);
  } finally {
    await book.close();
  }
});

test('traced native dynamic formulas coexist with bounded compatibility dependencies', async () => {
  const rows = Array.from(
    { length: 2000 },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c><c r="B${i + 1}"><f>A${i + 1}*2</f></c>${i === 0 ? '<c r="C1"><f>INDIRECT("A1")</f></c><c r="D1"><f>C1+1</f></c><c r="E1"><f>LEN(LEFT("abcd",3))</f></c>' : ''}</row>`
  ).join('');
  const { source } = await fixture(sheet(rows));
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
    assert(diagnostic.inputCells < 20);
    const result = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1000', value: 7 }]);
    assert(diagnostic.nativePlanHit);
    assert.equal(diagnostic.evaluatedFormulaCells, 1);
    assert(diagnostic.inputCells < 10);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'C1:E1')).cells;
    assert.equal(cells.C1.calculation.value, 1);
    assert.equal(cells.D1.calculation.value, 2);
    assert.equal(cells.E1.calculation.value, 3);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'B1000')).cells.B1000.calculation.value, 14);
    assert.equal((await book.read('0', 'Sheet1', 'B1000')).cells.B1000.calculation.value, 2000);
    const distant = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'Z10000', value: 9 }]);
    assert.equal(diagnostic.nativePlanHit, true, 'Sheet growth preserves dimension-independent dependency plans');
    const next = await book.apply(distant.view.storage.state, [{ sheet: 'Sheet1', cell: 'Y9999', value: 8 }]);
    assert(diagnostic.nativePlanHit);
    assert.equal(diagnostic.evaluatedFormulaCells, 0);
    assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 2);
    assert.equal((await book.read(distant.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 1);
    const saved = await book.export(distant.view.storage.state);
    assert.equal(readWorkbook(await fs.readFile(saved.path)).sheets[0].cells.B1000.cachedValue, 14);
  } finally {
    await book.close();
  }
});
test('whole shared and array formula groups clear natively; partial clears remain atomic', async () => {
  for (const kind of ['shared', 'array']) {
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><f t="${kind}" ${kind === 'shared' ? 'si="0"' : ''} ref="A1:A2">B1+1</f><v>2</v></c><c r="B1"><v>1</v></c><c r="D1"><f>SUM(A1:A2)</f></c></row><row r="2"><c r="A2">${kind === 'shared' ? '<f t="shared" si="0"/>' : ''}<v>3</v></c><c r="B2"><v>2</v></c></row>`
      )
    );
    const book = await owner(source, native, () => assert.fail('writer fallback'));
    try {
      if (kind === 'array')
        await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'A2', value: 9 }]), {
          code: 'fill-unsupported',
        });
      if (kind === 'array')
        for (const range of ['A1', 'A2'])
          await assert.rejects(book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range }]), {
            code: 'fill-unsupported',
          });
      const cleared = await book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range: 'A1:A2' }]);
      const cells = (await book.read(cleared.view.storage.state, 'Sheet1', 'A1:D2')).cells;
      assert.equal(cells.A1.formulaType, undefined);
      assert.equal(cells.A2.formulaType, undefined);
      assert.equal(cells.A1.value, null);
      assert.equal(cells.A2.value, null);
      assert.equal(cells.D1.calculation.value, 0);
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.formulaType, kind);
      const saved = await book.export(cleared.view.storage.state),
        reopened = await owner(saved.path);
      try {
        assert.equal((await reopened.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 0);
      } finally {
        await reopened.close();
      }
    } finally {
      await book.close();
    }
  }
});

test('mixed sorting matches the saved Excel 16.113 oracle in both directions', async () => {
  const values = [
    10,
    '2',
    2,
    '10',
    'a',
    'A',
    'b',
    false,
    true,
    null,
    '#N/A',
    '#DIV/0!',
    -3,
    0,
    'z',
    null,
    'é',
    'e',
    '中',
    '文',
    'a-b',
    'ab',
    "a'b",
  ];
  const rows = values
    .map(
      (v, i) =>
        `<row r="${i + 1}"><c r="A${i + 1}"${i === 10 || i === 11 ? ' t="e"' : typeof v === 'string' ? ' t="inlineStr"' : typeof v === 'boolean' ? ' t="b"' : ''}>${v === null ? '' : i === 10 || i === 11 ? `<v>${v}</v>` : typeof v === 'string' ? `<is><t>${v}</t></is>` : `<v>${Number(v)}</v>`}</c><c r="B${i + 1}"><v>${i + 1}</v></c></row>`
    )
    .join('');
  const { source } = await fixture(sheet(rows)),
    oracle = JSON.parse(await fs.readFile(new URL('../fixtures/excel-b5-oracle.json', import.meta.url), 'utf8'));
  let projectedKeys = 0;
  const book = await owner(
    source,
    async (input) => {
      if (input.operation === 'range') projectedKeys++;
      return native(input);
    },
    () => assert.fail('writer fallback')
  );
  try {
    for (const [direction, key] of [
      ['asc', 'SortAsc'],
      ['desc', 'SortDesc'],
    ]) {
      projectedKeys = 0;
      const result = await book.apply('0', [
        { operation: 'sortRows', sheet: 'Sheet1', range: 'A1:B23', column: 'A', direction },
      ]);
      assert.equal(projectedKeys, 0);
      const cells = (await book.read(result.view.storage.state, 'Sheet1', 'B1:B23')).cells;
      assert.deepEqual(
        Array.from({ length: 23 }, (_, i) => cells[`B${i + 1}`].value),
        oracle.values[key].map((row) => row[1])
      );
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 10);
    }
  } finally {
    await book.close();
  }
});

test('native chart sheet copying clones relationship ownership and preserves images and literal cells', async () => {
  const source = fileURLToPath(new URL('../fixtures/chart-source.xlsx', import.meta.url)),
    original = unzipSync(await fs.readFile(source));
  let projected = 0;
  const book = await owner(
    source,
    async (input) => {
      if (input.operation === 'range') projected++;
      return native(input);
    },
    () => assert.fail('writer fallback')
  );
  try {
    const copy = await book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Data', newName: 'Data Copy' }]);
    assert.equal(projected, 0);
    const data = unzipSync(await fs.readFile((await book.export(copy.view.storage.state)).path));
    assert.deepEqual(data['xl/worksheets/sheet1.xml'], original['xl/worksheets/sheet1.xml']);
    assert.deepEqual(data['xl/worksheets/sheet2.xml'], original['xl/worksheets/sheet1.xml']);
    for (const name of Object.keys(original).filter((n) => n.startsWith('xl/media/')))
      assert.deepEqual(data[name], original[name]);
    const chart = Object.keys(data).find((n) => /^xl\/charts\/chart1-copy\d+\.xml$/.test(n));
    assert(chart);
    assert.match(strFromU8(data[chart]), /'Data Copy'!\$B\$2:\$B\$4/);
    assert.equal(strFromU8(data['xl/charts/chart1.xml']), strFromU8(original['xl/charts/chart1.xml']));
    assert.match(strFromU8(data['xl/worksheets/_rels/sheet2.xml.rels']), /drawing1-copy1.xml/);
    assert.match(strFromU8(data['xl/drawings/_rels/drawing1-copy1.xml.rels']), /chart1-copy1.xml/);
    assert.match(strFromU8(data['xl/drawings/_rels/drawing1-copy1.xml.rels']), /image1.png/);
    assert.equal((await book.read(copy.view.storage.state, 'Data Copy', 'B4')).cells.B4.value, 5);
    assert.equal(book.view('0').sheets.length, 1);
    const again = await book.apply(copy.view.storage.state, [
      { operation: 'duplicateSheet', sheet: 'Data', newName: 'Second' },
    ]);
    assert.equal(again.view.sheets.length, 3);
    assert.equal((await book.read(again.view.storage.state, 'Second', 'G1')).cells.G1.value, 'Merged label');
  } finally {
    await book.close();
  }
});
test('deleting every table column removes its package references and rewrites structured references', async () => {
  const source = fileURLToPath(new URL('../fixtures/delete-table-source.xlsx', import.meta.url)),
    book = await owner(source, native, () => assert.fail('writer fallback'));
  try {
    const changed = await book.apply('0', [{ operation: 'deleteColumn', sheet: 'Data', index: 0, count: 2 }]);
    const saved = await book.export(changed.view.storage.state),
      bytes = await fs.readFile(saved.path),
      parts = unzipSync(bytes),
      view = readWorkbook(bytes);
    assert.equal(view.tables.length, 0);
    assert.equal(view.sheets[0].cells.B1.formula, 'SUM(#REF!)');
    assert.equal(view.sheets[0].cells.B2.formula, 'COUNTA(#REF!)');
    assert.equal(view.namedRanges[0].formula, '#REF!');
    assert(!Object.keys(parts).some((n) => n.startsWith('xl/tables/')));
    assert(!strFromU8(parts['[Content_Types].xml']).includes('/xl/tables/'));
    assert(!strFromU8(parts['xl/worksheets/sheet1.xml']).includes('tableParts'));
    assert(!strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']).includes('/table'));
    assert.equal(book.view('0').tables.length, 1);
    assert.equal((await book.read('0', 'Data', 'B3')).cells.B3.value, 3);
  } finally {
    await book.close();
  }
});

test('malformed shared groups cannot orphan followers when clearing their declared range', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">B1+1</f><v>2</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>4</v></c></row>'
    )
  );
  const book = await owner(source);
  try {
    await assert.rejects(
      book.apply('0', [{ type: 'clearRange', sheet: 'Sheet1', range: 'A1:A2' }]),
      /orphan|read-only|protected/
    );
    assert.equal(book.view('0').readOnly, false);
    assert.equal(book.view('0').calculationPolicy.preserveSource, true);
    assert.deepEqual(await fs.readFile((await book.export('0')).path), await fs.readFile(source));
    assert.equal((await book.read('0', 'Sheet1', 'A3')).cells.A3.formulaType, 'shared');
  } finally {
    await book.close();
  }
});
test('external chart relationships reject and bare deleted-table bindings rewrite atomically', async () => {
  const { source } = await fixture(sheet(''));
  const parts = unzipSync(await fs.readFile(new URL('../fixtures/chart-source.xlsx', import.meta.url)));
  parts['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(
    strFromU8(parts['xl/drawings/_rels/drawing1.xml.rels']).replace(
      'Target="/xl/charts/chart1.xml"',
      'Target="https://example.invalid/chart.xml" TargetMode="External"'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const book = await owner(source);
  try {
    await assert.rejects(book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Data', newName: 'Refused' }]), {
      code: 'package-unsupported',
    });
    assert.equal(book.view('0').sheets.length, 1);
  } finally {
    await book.close();
  }
  const other = await owner(fileURLToPath(new URL('../fixtures/delete-table-source.xlsx', import.meta.url)));
  try {
    const edited = await other.apply('0', [{ sheet: 'Data', cell: 'D1', formula: 'SUM(Sales)' }]);
    const deleted = await other.apply(edited.view.storage.state, [
      { operation: 'deleteColumn', sheet: 'Data', index: 0, count: 2 },
    ]);
    assert.equal((await other.read(deleted.view.storage.state, 'Data', 'B1')).cells.B1.formula, 'SUM(#REF!)');
    assert.equal(other.view(edited.view.storage.state).tables.length, 1);
  } finally {
    await other.close();
  }
});

test('runtime targets replace old edges, preserve undo plans and survive compaction', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>B1</t></is></c><c r="B1"><v>2</v></c><c r="C1"><v>7</v></c><c r="D1"><f>INDIRECT(A1)</f></c><c r="E1"><f>D1*2</f></c><c r="F1"><v>1</v></c></row>'
    )
  );
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  await book.open(source);
  try {
    assert.equal(diagnostic.tracedFormulaCells, 1);
    const first = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 3 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 2);
    assert.equal((await book.read(first.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 6);
    const retarget = await book.apply(first.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 'C1' }]);
    assert.equal((await book.read(retarget.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 14);
    const ignored = await book.apply(retarget.view.storage.state, [{ sheet: 'Sheet1', cell: 'B1', value: 99 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 0);
    const changed = await book.apply(ignored.view.storage.state, [{ sheet: 'Sheet1', cell: 'C1', value: 8 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 2);
    assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 16);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 5 }]);
    assert.equal((await book.read(branch.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 10);
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 4);
    const failed = await book.apply(changed.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 'missing' }]);
    assert.equal(
      (await book.read(failed.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.status,
      'unresolved'
    );
    const recovered = await book.apply(failed.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 'B1' }]);
    assert.equal((await book.read(recovered.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 198);
    await book.compact(['0', recovered.view.storage.state]);
    const afterCompact = await book.apply(recovered.view.storage.state, [{ sheet: 'Sheet1', cell: 'B1', value: 100 }]);
    assert.equal((await book.read(afterCompact.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 200);
    const textTarget = await book.apply(recovered.view.storage.state, [
      { sheet: 'Sheet1', cell: 'C1', value: 'text' },
      { sheet: 'Sheet1', cell: 'A1', value: 'C1' },
    ]);
    assert.equal((await book.read(textTarget.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 'text');
    const repairedTarget = await book.apply(textTarget.view.storage.state, [{ sheet: 'Sheet1', cell: 'C1', value: 9 }]);
    assert.equal((await book.read(repairedTarget.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 18);
    const saved = await book.export(recovered.view.storage.state),
      again = await owner(saved.path);
    try {
      assert.equal((await again.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 198);
    } finally {
      await again.close();
    }
  } finally {
    await book.close();
  }
});
test('ordinary formula sheet copies rewrite only copied self references and preserve originals', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>Sheet1!A1+$A$1</f><v>4</v></c></row>')
  );
  const book = await owner(source, native, () => assert.fail('writer fallback'));
  try {
    const copied = await book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Sheet1', newName: 'Copy' }]);
    assert.equal((await book.read(copied.view.storage.state, 'Copy', 'B1')).cells.B1.formula, "'Copy'!A1+$A$1");
    assert.equal((await book.read(copied.view.storage.state, 'Sheet1', 'B1')).cells.B1.formula, 'Sheet1!A1+$A$1');
    const edited = await book.apply(copied.view.storage.state, [{ sheet: 'Copy', cell: 'A1', value: 7 }]);
    assert.equal((await book.read(edited.view.storage.state, 'Copy', 'B1')).cells.B1.calculation.value, 14);
    assert.equal((await book.read(edited.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 4);
    assert.equal(book.view('0').sheets.length, 1);
    const saved = await book.export(edited.view.storage.state),
      again = await owner(saved.path);
    try {
      assert.equal((await again.read('0', 'Copy', 'B1')).cells.B1.calculation.value, 14);
    } finally {
      await again.close();
    }
  } finally {
    await book.close();
  }
});

test('copied local names receive new scope and rewritten definitions without changing global names', async () => {
  const { source, parts } = await fixture(
    sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>LocalAmount+GlobalAmount</f></c></row>')
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      '<definedNames><definedName name="LocalAmount" localSheetId="0">Sheet1!$A$1</definedName><definedName name="GlobalAmount">Sheet1!$A$1</definedName></definedNames></workbook>'
    )
  );
  await fs.writeFile(source, zipSync(parts));
  const book = await owner(source, native, () => assert.fail('writer fallback'));
  try {
    const copied = await book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Sheet1', newName: 'Copy' }]);
    const local = copied.view.namedRanges.filter((n) => n.name === 'LocalAmount');
    assert.equal(local.length, 2);
    assert.equal(local.find((n) => n.sheet === 'Copy').formula, "'Copy'!$A$1");
    assert.equal(copied.view.namedRanges.find((n) => n.name === 'GlobalAmount').formula, 'Sheet1!$A$1');
    const edited = await book.apply(copied.view.storage.state, [{ sheet: 'Copy', cell: 'A1', value: 7 }]);
    assert.equal((await book.read(edited.view.storage.state, 'Copy', 'B1')).cells.B1.calculation.value, 9);
    assert.equal((await book.read(edited.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 4);
    assert.equal(book.view('0').namedRanges.length, 2);
  } finally {
    await book.close();
  }
});
test('formula copies still reject dynamic and special references without publishing partial metadata', async () => {
  for (const formula of ['INDIRECT("A1")', 'A1']) {
    const special = formula === 'A1';
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f${special ? ' t="array" ref="B1:B2"' : ''}>${formula}</f><v>2</v></c></row>`
      )
    );
    const book = await owner(source, native, () => assert.fail('writer fallback'));
    try {
      await assert.rejects(
        book.apply('0', [{ operation: 'duplicateSheet', sheet: 'Sheet1', newName: 'Copy' }]),
        /Dynamic|Special|qualification|qualified/
      );
      assert.equal(book.view('0').sheets.length, 1);
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 2);
    } finally {
      await book.close();
    }
  }
});

test('formula-derived text selectors calculate in SQLite, retarget transitively and preserve typed export', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>LEFT("Cxxx",1)&amp;A1</f></c><c r="D1"><f>INDIRECT(B1)</f></c><c r="E1"><f>D1+1</f></c><c r="F1"><f>IF(A1=2,"old","new")</f></c></row><row r="2"><c r="C2"><v>7</v></c></row><row r="3"><c r="C3"><v>11</v></c></row>'
    )
  );
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
    assert(diagnostic.nativeSqlite);
    assert.equal(diagnostic.inputCells, 0);
    assert.equal((await book.read('0', 'Sheet1', 'E1')).cells.E1.calculation.value, 8);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 3 }]);
    assert(diagnostic.nativeSqlite);
    assert.equal(diagnostic.evaluatedFormulaCells, 4);
    const cells = (await book.read(changed.view.storage.state, 'Sheet1', 'B1:F1')).cells;
    assert.equal(cells.B1.calculation.value, 'C3');
    assert.equal(cells.D1.calculation.value, 11);
    assert.equal(cells.E1.calculation.value, 12);
    assert.equal(cells.F1.calculation.value, 'new');
    const ignored = await book.apply(changed.view.storage.state, [{ sheet: 'Sheet1', cell: 'C2', value: 99 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 0);
    const retarget = await book.apply(ignored.view.storage.state, [{ sheet: 'Sheet1', cell: 'C3', value: 13 }]);
    assert.equal((await book.read(retarget.view.storage.state, 'Sheet1', 'E1')).cells.E1.calculation.value, 14);
    const saved = await book.export(retarget.view.storage.state),
      disk = readWorkbook(await fs.readFile(saved.path));
    assert.equal(disk.sheets[0].cells.B1.cachedValue, 'C3');
    assert.equal(disk.sheets[0].cells.E1.cachedValue, 14);
    assert.equal(disk.sheets[0].cells.F1.cachedValue, 'new');
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 'C2');
  } finally {
    await book.close();
  }
});
test('complete shared and array groups accept replacement values/formulas while partial edits stay atomic', async () => {
  for (const kind of ['shared', 'array']) {
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><f t="${kind}" ${kind === 'shared' ? 'si="0"' : ''} ref="A1:A2">B1+1</f><v>3</v></c><c r="B1"><v>2</v></c><c r="C1"><f>$B$1*2</f></c><c r="D1"><v>9</v></c></row><row r="2"><c r="A2">${kind === 'shared' ? '<f t="shared" si="0"/>' : ''}<v>4</v></c></row>`
      )
    );
    const book = await owner(source, native, () => assert.fail('fallback'));
    try {
      for (const type of ['pasteValues', 'fillRange']) {
        const command = { type, sheet: 'Sheet1', source: type === 'pasteValues' ? 'D1' : 'C1', range: 'A1:A2' };
        if (kind === 'array')
          await assert.rejects(book.apply('0', [{ ...command, range: 'A1' }]), { code: 'fill-unsupported' });
        const replaced = await book.apply('0', [command]);
        const cells = (await book.read(replaced.view.storage.state, 'Sheet1', 'A1:A2')).cells;
        for (const c of Object.values(cells)) {
          assert.equal(c.formulaType, undefined);
          if (type === 'pasteValues') assert.equal(c.value, 9);
          else {
            assert.equal(c.formula, '$B$1*2');
            assert.equal(c.calculation.value, 4);
          }
        }
        const saved = readWorkbook(await fs.readFile((await book.export(replaced.view.storage.state)).path));
        assert.equal(saved.sheets[0].cells.A1.formulaType, undefined);
        assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.formulaType, kind);
      }
    } finally {
      await book.close();
    }
  }
});
test('whole-table row deletion removes ownership metadata and rewrites bare/structured names', async () => {
  const book = await owner(
    fileURLToPath(new URL('../fixtures/delete-table-source.xlsx', import.meta.url)),
    native,
    () => assert.fail('fallback')
  );
  try {
    const populated = await book.apply('0', [
      { sheet: 'Data', cell: 'D5', formula: 'SUM(Sales)' },
      { sheet: 'Data', cell: 'E5', formula: 'SUM(Sales[Amount])' },
    ]);
    await assert.rejects(
      book.apply(populated.view.storage.state, [{ operation: 'deleteRow', sheet: 'Data', index: 0, count: 1 }]),
      /header/
    );
    const deleted = await book.apply(populated.view.storage.state, [
      { operation: 'deleteRow', sheet: 'Data', index: 0, count: 3 },
    ]);
    assert.equal(deleted.view.tables.length, 0);
    const cells = (await book.read(deleted.view.storage.state, 'Data', 'D2:E2')).cells;
    assert.equal(cells.D2.formula, 'SUM(#REF!)');
    assert.equal(cells.E2.formula, 'SUM(#REF!)');
    assert.equal(deleted.view.namedRanges.find((n) => n.name === 'SalesRange').formula, '#REF!');
    assert.equal(book.view('0').tables.length, 1);
    const parts = unzipSync(await fs.readFile((await book.export(deleted.view.storage.state)).path));
    assert(!parts['xl/tables/table1.xml']);
    assert(!strFromU8(parts['xl/worksheets/sheet1.xml']).includes('tablePart'));
  } finally {
    await book.close();
  }
});

test('partial shared edits expand anchored formulas with atomic undo, batch, bulk and native export', async () => {
  const rows = [1, 2, 3]
    .map(
      (n) =>
        `<row r="${n}"><c r="A${n}"><f t="shared" si="0"${n === 1 ? ' ref="A1:A3"' : ''}>${n === 1 ? '$B1+C$1' : ''}</f><v>999</v></c><c r="B${n}"><v>${n}</v></c>${n === 1 ? '<c r="C1"><v>10</v></c><c r="D1"><f>SUM(A1:A3)</f></c>' : ''}</row>`
    )
    .join('');
  const { source } = await fixture(sheet(rows));
  const book = await owner(source, native, () => assert.fail('writer fallback'));
  try {
    for (const commands of [
      [{ sheet: 'Sheet1', cell: 'A2', value: 50 }],
      [{ type: 'clearRange', sheet: 'Sheet1', range: 'A2' }],
      [{ type: 'fillRange', sheet: 'Sheet1', source: 'B1', range: 'A2' }],
      [
        { sheet: 'Sheet1', cell: 'A1', formula: 'B1*2' },
        { sheet: 'Sheet1', cell: 'A3', value: 8 },
      ],
    ]) {
      const changed = await book.apply('0', commands),
        cells = (await book.read(changed.view.storage.state, 'Sheet1', 'A1:D3')).cells;
      for (const ref of ['A1', 'A2', 'A3']) assert.equal(cells[ref]?.formulaType, undefined);
      if (commands.length === 1) {
        assert.equal(cells.A1.formula, '$B1+C$1');
        assert.equal(cells.A3.formula, '$B3+C$1');
        assert.equal(cells.A1.calculation.value, 11);
        assert.equal(cells.A3.calculation.value, 13);
        assert.equal(
          cells.D1.calculation.value,
          commands[0].type === 'clearRange' ? 24 : commands[0].type === 'fillRange' ? 25 : 74
        );
      } else {
        assert.equal(cells.A2.formula, '$B2+C$1');
        assert.equal(cells.D1.calculation.value, 22);
      }
      const bytes = await fs.readFile((await book.export(changed.view.storage.state)).path);
      const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
      assert(!xml.includes('t="shared"'));
      assert.equal(readWorkbook(bytes).sheets[0].cells.A2?.formulaType, undefined);
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.formulaType, 'shared');
    }
    const guarded = await book.apply('0', [
      { sheet: 'Sheet1', cell: 'A2', value: 17, expectedRawValue: '999', expectedFormula: null },
    ]);
    assert.equal((await book.read(guarded.view.storage.state, 'Sheet1', 'A2')).cells.A2.value, 17);
    await assert.rejects(
      book.apply('0', [{ sheet: 'Sheet1', cell: 'A2', value: 99, expectedFormula: 'wrong' }]),
      /Conflict/
    );
    await assert.rejects(
      book.apply('0', [
        { sheet: 'Sheet1', cell: 'A2', value: 50 },
        { sheet: 'Missing', cell: 'A1', value: 3 },
      ])
    );
    assert.equal((await book.read('0', 'Sheet1', 'A2')).cells.A2.formulaType, 'shared');
  } finally {
    await book.close();
  }
});
test('shared expansion refuses incomplete, conflicting, oversized and orphaned groups', async () => {
  for (const xml of [
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:XFD1048576">1</f></c></row>',
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A1" ca="1">1</f></c></row>',
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">1</f></c></row><row r="2"><c r="A2"><f t="shared" si="0" ca="1"/></c></row>',
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A2:A3">B1+1</f></c></row>',
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">B1+1</f></c></row><row r="2"><c r="A2"><f t="shared" si="0">B2+9</f></c></row>',
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">1</f></c></row><row r="3"><c r="A3"><f t="shared" si="0"/></c></row>',
  ]) {
    const { source } = await fixture(sheet(xml)),
      book = await owner(source);
    try {
      await assert.rejects(
        book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]),
        /shared|Shared|read-only|metadata|protected/
      );
      assert.deepEqual(await fs.readFile((await book.export('0')).path), await fs.readFile(source));
      assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.formulaType, 'shared');
    } finally {
      await book.close();
    }
  }
});
test('shared groups whose ranges overlap are edited like any other: each cell names its own group', async () => {
  // WPS and Excel both write these. A2 sits inside the first group's range but belongs to the second.
  const xml =
    '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A3">B1+1</f></c><c r="B1"><v>1</v></c></row>' +
    '<row r="2"><c r="A2"><f t="shared" si="1" ref="A2:A3">B2+2</f></c><c r="B2"><v>1</v></c></row>' +
    '<row r="3"><c r="A3"><f t="shared" si="1"/></c><c r="B3"><v>1</v></c></row>';
  const { source } = await fixture(sheet(xml)),
    book = await owner(source);
  try {
    assert.notEqual(book.view('0').calculationPolicy?.preserveSource, true);
    const edited = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]),
      state = edited.view.storage.state,
      cells = (await book.read(state, 'Sheet1', 'A1:A3')).cells;
    assert.equal(cells.A1.value, 9);
    assert.equal(cells.A2.formula, 'B2+2');
    const inserted = await book.apply(state, [{ operation: 'insertRow', sheet: 'Sheet1', index: 1 }]),
      moved = (await book.read(inserted.view.storage.state, 'Sheet1', 'A1:A4')).cells;
    assert.equal(moved.A3.formula, 'B3+2');
    assert.equal(moved.A4.formula, 'B4+2');
  } finally {
    await book.close();
  }
});
test('empty-table deletion preserves headers and relationships without absorbing the next row', async () => {
  const source = fileURLToPath(new URL('../fixtures/delete-table-source.xlsx', import.meta.url)),
    book = await owner(source, native, () => assert.fail('fallback'));
  try {
    const result = await book.apply('0', [{ operation: 'deleteRow', sheet: 'Data', index: 1, count: 2 }]);
    assert.equal(result.view.tables[0].ref, 'A1:B2');
    const cells = (await book.read(result.view.storage.state, 'Data', 'A1:D2')).cells;
    assert.equal(cells.A1.value, 'Item');
    assert.equal(cells.B1.value, 'Amount');
    assert.equal(cells.D1.calculation.value, 0);
    const bytes = await fs.readFile((await book.export(result.view.storage.state)).path),
      parts = unzipSync(bytes);
    assert.match(strFromU8(parts['xl/tables/table1.xml']), /insertRow="1"/);
    assert.match(strFromU8(parts['xl/tables/table1.xml']), /ref="A1:B2"/);
    assert.equal(readWorkbook(bytes).tables.length, 1);
    const occupied = await book.apply('0', [{ sheet: 'Data', cell: 'B4', value: 999 }]);
    await assert.rejects(
      book.apply(occupied.view.storage.state, [{ operation: 'deleteRow', sheet: 'Data', index: 1, count: 2 }]),
      /absorb/
    );
    assert.equal((await book.read(occupied.view.storage.state, 'Data', 'B4')).cells.B4.value, 999);
    await assert.rejects(book.apply('0', [{ operation: 'deleteRow', sheet: 'Data', index: 0, count: 1 }]), /header/);
  } finally {
    await book.close();
  }
});

test('What-If Data Table result groups clear or replace as a whole; partial and stale-cache copies reject', async () => {
  const source = fileURLToPath(new URL('../fixtures/excel-data-table23.xlsx', import.meta.url)),
    book = await owner(source, native, () => assert.fail('fallback'));
  try {
    for (const ref of ['D2', 'D3', 'D4']) {
      await assert.rejects(book.apply('0', [{ sheet: 'Data', cell: ref, value: 99 }]), /complete group|not qualified/);
      await assert.rejects(book.apply('0', [{ type: 'clearRange', sheet: 'Data', range: ref }]), {
        code: 'fill-unsupported',
      });
    }
    const copied = await book.apply('0', [{ type: 'pasteValues', sheet: 'Data', source: 'D2:D4', range: 'H2:H4' }]);
    const originalValues = (await book.read('0', 'Data', 'D2:D4')).cells,
      copyValues = (await book.read(copied.view.storage.state, 'Data', 'H2:H4')).cells;
    for (const row of [2, 3, 4]) assert.equal(copyValues['H' + row].value, originalValues['D' + row].calculation.value);
    const manual = await book.apply('0', [
        { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false },
      ]),
      pending = await book.apply(manual.view.storage.state, [{ sheet: 'Data', cell: 'A1', value: 7 }]);
    await assert.rejects(
      book.apply(pending.view.storage.state, [{ type: 'pasteValues', sheet: 'Data', source: 'D2:D4', range: 'H2:H4' }]),
      /current|pending|stored/i
    );
    for (const command of [
      { type: 'clearRange', sheet: 'Data', range: 'D2:D4' },
      { type: 'pasteValues', sheet: 'Data', range: 'D2:D4', source: 'F1' },
      { type: 'fillRange', sheet: 'Data', range: 'D2:D4', source: 'G1' },
    ]) {
      const result = await book.apply('0', [command]),
        cells = (await book.read(result.view.storage.state, 'Data', 'A1:G4')).cells;
      for (const ref of ['D2', 'D3', 'D4']) {
        assert.equal(cells[ref]?.formulaType, undefined);
        if (command.type === 'clearRange') assert.equal(cells[ref]?.value ?? null, null);
        else if (command.type === 'pasteValues') assert.equal(cells[ref].value, 9);
        else assert.equal(cells[ref].formula, `C${Number(ref.slice(1))}*2`);
      }
      assert.equal(cells.A1.value, 2);
      assert.equal(cells.C4.value, 3);
      assert.equal(cells.B1.calculation.value, 6);
      const bytes = await fs.readFile((await book.export(result.view.storage.state)).path);
      assert(!strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']).includes('t="dataTable"'));
      assert.equal((await book.read('0', 'Data', 'D2')).cells.D2.formulaType, 'dataTable');
    }
  } finally {
    await book.close();
  }
});

test('integer text arithmetic remains native through retargeted edits, old revisions and export', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>12</t></is></c><c r="B1"><f>A1+1</f></c><c r="C1"><f>B1*2</f></c></row>'
    )
  );
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
    assert.equal(diagnostic.inputCells, 0);
    const result = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: '24' }]);
    assert.equal(diagnostic.inputCells, 0);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 50);
    assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 26);
    const disk = readWorkbook(await fs.readFile((await book.export(result.view.storage.state)).path));
    assert.equal(disk.sheets[0].cells.C1.cachedValue, 50);
    const fraction = await book.apply(result.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: '0.1' }]);
    assert.equal((await book.read(fraction.view.storage.state, 'Sheet1', 'B1')).cells.B1.calculation.value, 1.1);
    assert.equal(diagnostic.inputCells, 0);
  } finally {
    await book.close();
  }
});

test('native calculation snapshot preserves cumulative export styles and accepted branch intent', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'workbook-calc-style-'));
  const source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>')
  );
  await fs.writeFile(source, zipSync(parts));
  const book = await owner(source);
  try {
    const styled = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', bold: true }]);
    const branch = await book.apply(styled.view.storage.state, [{ sheet: 'Sheet1', cell: 'B1', italic: true }]);
    const changed = await book.apply(branch.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
    const fork = await book.apply(styled.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
    for (const [state, value, italic] of [
      [changed.view.storage.state, 14, true],
      [fork.view.storage.state, 18, false],
    ]) {
      const saved = readWorkbook(await fs.readFile((await book.export(state)).path)).sheets[0].cells.B1;
      assert.equal(saved.cachedValue, value);
      assert.equal(saved.formula, 'A1*2');
      assert.equal(saved.bold, true);
      assert.equal(Boolean(saved.italic), italic);
    }
    assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.calculation.value, 4);
  } finally {
    await book.close();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('native shared calculation preserves compression metadata and refreshes caches without JavaScript projection', async () => {
  const rows = Array.from(
    { length: 1200 },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}"><f t="shared" si="0"${i === 0 ? ' ref="A1:A1200"' : ''}>${i === 0 ? '$C$1+1' : ''}</f><v>999</v></c>${i === 0 ? '<c r="C1"><v>2</v></c><c r="D1"><f>SUM(A1:A1200)</f></c>' : ''}</row>`
  ).join('');
  const { source } = await fixture(sheet(rows));
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  try {
    await book.open(source);
    assert.equal(diagnostic.inputCells, 0);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 3600);
    assert.equal((await book.read('0', 'Sheet1', 'A1200')).cells.A1200.formulaType, 'shared');
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'C1', value: 4 }]);
    assert.equal(diagnostic.inputCells, 0);
    assert.equal((await book.read(changed.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 6000);
    const bytes = await fs.readFile((await book.export(changed.view.storage.state)).path),
      saved = readWorkbook(bytes),
      xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
    assert.equal(saved.sheets[0].cells.A1200.cachedValue, 5);
    assert.equal(saved.sheets[0].cells.A1200.formulaType, 'shared');
    assert.equal((xml.match(/t="shared"/g) ?? []).length, 1200);
    assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 3600);
    const partial = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A600', value: 9 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 1);
    assert.equal((await book.read(partial.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 3606);
  } finally {
    await book.close();
  }
});

test('What-If column, row and two-input models match Excel, recalculate consumers and preserve inputs/undo/export', async () => {
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('fallback'),
    async (input) => (diagnostic = await calculateIndexedWorkbook(input, native))
  );
  await book.open(fileURLToPath(new URL('../fixtures/excel-scenarios-coercion24.xlsx', import.meta.url)));
  try {
    assert.equal(diagnostic.inputCells, 0);
    for (const [name, total] of [
      ['Column', 58],
      ['Row', 58],
      ['Two', 312],
    ])
      assert.equal((await book.read('0', name, 'J1')).cells.J1.calculation.value, total);
    const changed = await book.apply('0', [
      { sheet: 'Two', cell: 'E5', value: 30 },
      { sheet: 'Column', cell: 'A3', value: 4 },
      { sheet: 'Row', cell: 'G4', value: 7 },
    ]);
    assert.equal(diagnostic.inputCells, 0);
    const state = changed.view.storage.state;
    for (const [name, total] of [
      ['Column', 80],
      ['Row', 78],
      ['Two', 492],
    ])
      assert.equal((await book.read(state, name, 'J1')).cells.J1.calculation.value, total, name);
    const cells = (await book.read(state, 'Two', 'A1:H6')).cells;
    assert.equal(cells.A1.value, 2);
    assert.equal(cells.A2.value, 5);
    assert.equal(cells.B1.calculation.value, 21);
    assert.equal(cells.F5.calculation.value, 93);
    assert.equal(cells.G6.calculation.value, 66);
    await assert.rejects(book.apply(state, [{ sheet: 'Two', cell: 'G6', value: 3 }]), /complete group|partial/);
    const bytes = await fs.readFile((await book.export(state)).path),
      saved = readWorkbook(bytes);
    assert.equal(saved.sheets.find((s) => s.name === 'Two').cells.F5.cachedValue, 93);
    assert.equal(saved.sheets.find((s) => s.name === 'Two').cells.F5.formulaType, 'dataTable');
    assert.equal(saved.sheets.find((s) => s.name === 'Two').cells.G6.value, 66);
    assert.equal((await book.read('0', 'Two', 'J1')).cells.J1.calculation.value, 312);
    assert.equal((await book.selectionSummary(state, 'Two', 'F5:H6')).sum, 492);
    await assert.rejects(book.apply('0', [{ sheet: 'Two', cell: 'B1', formula: 'INDIRECT("A1")' }]), /What-If/);
    assert.equal((await book.read('0', 'Two', 'J1')).cells.J1.calculation.value, 312);
  } finally {
    await book.close();
  }
});

test('native whole-row copy and structural edits never project formulas into JavaScript', async () => {
  const { source } = await fixture(
    sheet('<row r="1"><c r="A1"><v>3</v></c></row><row r="2"><c r="B2"><f>SUM(1:1)</f></c></row>')
  );
  const calls = [];
  const book = await owner(
    source,
    async (i) => {
      calls.push(i);
      return native(i);
    },
    () => assert.fail('writer fallback')
  );
  try {
    calls.length = 0;
    const copied = await book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'B2', range: 'B3:B4' }]);
    const cells = (await book.read(copied.view.storage.state, 'Sheet1', 'B3:B4')).cells;
    assert.equal(cells.B3.formula, 'SUM(2:2)');
    assert.equal(cells.B4.calculation.value, 3);
    assert(
      !calls.some((i) => i.formulaOverrides || (i.operation === 'artifact-write' && i.name === 'bulk-formulas.jsonl'))
    );
    calls.length = 0;
    const moved = await book.apply(copied.view.storage.state, [{ operation: 'insertRow', sheet: 'Sheet1', index: 0 }]);
    assert.equal((await book.read(moved.view.storage.state, 'Sheet1', 'B3')).cells.B3.formula, 'SUM(2:2)');
    assert(!calls.some((i) => i.operation === 'range' && i.formulas));
    const saved = readWorkbook(await fs.readFile((await book.export(moved.view.storage.state)).path));
    assert.equal(saved.sheets[0].cells.B5.cachedValue, 3);
  } finally {
    await book.close();
  }
});
test('shared and array style-only changes preserve formula XML, caches, branches and reopen', async () => {
  for (const type of ['shared', 'array']) {
    const { source } = await fixture(
      sheet(
        `<row r="1"><c r="A1"><f t="${type}"${type === 'shared' ? ' si="0"' : ''} ref="A1:A2">1+1</f><v>2</v></c></row><row r="2"><c r="A2">${type === 'shared' ? '<f t="shared" si="0"/>' : ''}<v>2</v></c></row>`
      )
    );
    const book = await owner(source);
    try {
      const r = await book.apply('0', [
        { type: 'formatRange', sheet: 'Sheet1', range: 'A1:A2', format: { bold: true, numberFormat: 'currency' } },
      ]);
      const output = await fs.readFile((await book.export(r.view.storage.state)).path),
        xml = strFromU8(unzipSync(output)['xl/worksheets/sheet1.xml']);
      assert.match(xml, new RegExp(`t="${type}"`));
      assert.match(xml, /ref="A1:A2"/);
      const view = readWorkbook(output);
      assert.equal(view.sheets[0].cells.A1.bold, true);
      assert.equal(view.sheets[0].cells.A1.cachedValue, 2);
      assert(!(await book.read('0', 'Sheet1', 'A1')).cells.A1.bold);
      const branch = await book.apply('0', [
        { type: 'formatRange', sheet: 'Sheet1', range: 'A1', format: { italic: true } },
      ]);
      assert.equal(
        readWorkbook(await fs.readFile((await book.export(branch.view.storage.state)).path)).sheets[0].cells.A1.italic,
        true
      );
    } finally {
      await book.close();
    }
  }
});
test('formula replacement and scenario invalidation use precise versioned native dependency plans', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+1</f></c><c r="C1"><f>B1+1</f></c><c r="D1"><f>100+1</f></c></row>'
    )
  );
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (i) => (diagnostic = await calculateIndexedWorkbook(i, native))
  );
  try {
    await book.open(source);
    const r = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', formula: 'A1+10' }]);
    assert(diagnostic.nativePlanHit);
    assert.equal(diagnostic.evaluatedFormulaCells, 2);
    assert.equal((await book.read(r.view.storage.state, 'Sheet1', 'C1')).cells.C1.calculation.value, 13);
    assert.equal((await book.read('0', 'Sheet1', 'C1')).cells.C1.calculation.value, 4);
    const full = await book.apply(r.view.storage.state, [{ type: 'recalculate' }]);
    assert.equal((await book.read(full.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 101);
  } finally {
    await book.close();
  }
  const scenario = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (i) => (diagnostic = await calculateIndexedWorkbook(i, native))
  );
  try {
    await scenario.open(fileURLToPath(new URL('../fixtures/excel-scenarios-coercion24.xlsx', import.meta.url)));
    const unrelated = await scenario.apply('0', [{ sheet: 'Column', cell: 'Z100', value: 7 }]);
    assert.equal(diagnostic.evaluatedFormulaCells, 0);
    assert.equal((await scenario.read(unrelated.view.storage.state, 'Column', 'J1')).cells.J1.calculation.value, 58);
    const changed = await scenario.apply(unrelated.view.storage.state, [{ sheet: 'Column', cell: 'A3', value: 4 }]);
    assert(diagnostic.evaluatedFormulaCells < 20);
    assert.equal((await scenario.read(changed.view.storage.state, 'Column', 'J1')).cells.J1.calculation.value, 80);
  } finally {
    await scenario.close();
  }
});

test('native display search agrees with grid formatting across rounding, date, precision and exponent boundaries', async () => {
  const { displayCell } = await import('@office-editors/xlsx-editor/core/display.mjs');
  const parts = unzipSync(createWorkbook());
  const formats = [0, 1, 2, 3, 4, 9, 10, 14, 164, 165];
  const values = [0, -0.0001, 0.1, 1.005, 2.675, 0.615, -2.5, 1e-7, 1e21, 60, -1, 1000000000000000100];
  const xml = strFromU8(parts['xl/styles.xml'])
    .replace(
      /<cellXfs[^>]*>[\s\S]*?<\/cellXfs>/,
      `<cellXfs count="${formats.length}">${formats.map((numFmtId) => `<xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0"/>`).join('')}</cellXfs>`
    )
    .replace(/<numFmts[^>]*\/>|<numFmts[^>]*>[\s\S]*?<\/numFmts>/, '');
  parts['xl/styles.xml'] = strToU8(
    xml.replace(
      /(<styleSheet[^>]*>)/,
      '$1<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0"/><numFmt numFmtId="165" formatCode="$#,##0.00"/></numFmts>'
    )
  );
  const rows = [];
  let r = 0;
  for (const value of values)
    for (let style = 0; style < formats.length; style++) {
      r++;
      rows.push(`<row r="${r}"><c r="A${r}" s="${style}"><v>${value}</v></c></row>`);
    }
  parts['xl/worksheets/sheet1.xml'] = strToU8(sheet(rows.join('')));
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-display-native-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const cells = readWorkbook(await fs.readFile(source)).sheets[0].cells,
    calls = [],
    book = await owner(source, async (i) => {
      calls.push(i);
      return native(i);
    });
  try {
    const shown = Object.fromEntries(Object.entries(cells).map(([ref, cell]) => [ref, displayCell(cell, false)]));
    for (const needle of new Set(Object.values(shown))) {
      calls.length = 0;
      const actual = await book.query('0', 'Sheet1', needle, 'display');
      const expected = Object.entries(shown)
        .filter(([, text]) => text.toLowerCase().includes(needle.toLowerCase()))
        .map(([ref]) => ref);
      assert.deepEqual(Object.keys(actual.cells), expected, needle);
      assert(calls.some((i) => i.operation === 'query' && i.mode === 'display'));
      assert(!calls.some((i) => i.operation === 'range'));
    }
  } finally {
    await book.close();
  }
});

test('large formula replacements update dependency storage without the old 1024-position limit', async () => {
  const rows = Array.from(
    { length: 1500 },
    (_, i) =>
      `<row r="${i + 1}"><c r="A${i + 1}"><v>2</v></c><c r="B${i + 1}"><f>$A${i + 1}+1</f></c><c r="C${i + 1}"><f>B${i + 1}*2</f></c>${i === 0 ? '<c r="D1"><f>100+1</f></c><c r="E1"><f>$A1+5</f></c><c r="F1"><f>CONCAT(&quot;x&quot;,A1)</f></c>' : ''}</row>`
  ).join('');
  const { source } = await fixture(sheet(rows));
  let diagnostic;
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('writer fallback'),
    async (i) => (diagnostic = await calculateIndexedWorkbook(i, native))
  );
  try {
    await book.open(source);
    const result = await book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'E1', range: 'B1:B1200' }]);
    assert(diagnostic.nativePlanHit);
    assert.equal(diagnostic.evaluatedFormulaCells, 2400);
    assert.equal(diagnostic.inputCells, 0);
    const cells = (await book.read(result.view.storage.state, 'Sheet1', 'B1200:C1200')).cells;
    assert.equal(cells.C1200.calculation.value, 14);
    assert.equal((await book.read(result.view.storage.state, 'Sheet1', 'C1500')).cells.C1500.calculation.value, 6);
    assert.equal((await book.read('0', 'Sheet1', 'C1200')).cells.C1200.calculation.value, 6);
    const replaced = await book.apply(result.view.storage.state, [
      { sheet: 'Sheet1', cell: 'F1', formula: 'CONCAT("y",A2)' },
    ]);
    assert(diagnostic.nativePlanHit);
    assert.equal(diagnostic.evaluatedFormulaCells, 1);
    assert(diagnostic.inputCells < 5);
    assert.equal((await book.read(replaced.view.storage.state, 'Sheet1', 'F1')).cells.F1.calculation.value, 'y2');
  } finally {
    await book.close();
  }
});

test('native import batches isolate rich text, empty cells, formulas and styles after oversized fields', async () => {
  const long = '界'.repeat(20000),
    xml = sheet(
      `<row r="1"><c r="A1" t="inlineStr" s="0"><is><r><t>${long}</t></r><r><t>_x0041_</t></r><rPh><t>ignored</t></rPh></is></c><c r="B1"/><c r="C1"><f>1+1</f><v>2</v></c></row><row r="2"><c r="A2"><v>7</v></c><c r="B2" t="b"><v>0</v></c><c r="C2" t="inlineStr"><is><t>short &amp; final</t></is></c><c r="D2"/></row>`
    );
  const { source, parts } = await fixture(xml),
    expected = readWorkbook(await fs.readFile(source)),
    traces = [];
  const book = await owner(source, async (input) => {
    const result = await native({ ...input, diagnostics: true });
    if (input.operation === 'index') traces.push(result.diagnostics);
    delete result.diagnostics;
    return result;
  });
  try {
    assert.equal(traces[0].counters['import.large_cell_flushes'], 1);
    assert(traces[0].counters['import.insert_statements'] >= 2, 'An oversized field must be flushed separately');
    const cells = (await book.read('0', 'Sheet1', 'A1:D2')).cells;
    for (const ref of ['A1', 'A2', 'B2', 'C2']) assert.deepEqual(cells[ref], expected.sheets[0].cells[ref], ref);
    for (const ref of ['B1', 'D2']) assert.equal(cells[ref]?.value ?? null, null, ref);
    assert.equal(cells.C1.calculation.value, 2);
    assert.equal(cells.A1.value, long + 'A');
    const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'D2', value: 'only edit' }]),
      output = await book.export(next.view.storage.state),
      bytes = await fs.readFile(output.path),
      saved = readWorkbook(bytes);
    for (const ref of ['A1', 'A2', 'B2', 'C2'])
      assert.deepEqual(saved.sheets[0].cells[ref], expected.sheets[0].cells[ref], ref);
    assert.equal(saved.sheets[0].cells.B1?.value ?? null, null);
    assert.equal(saved.sheets[0].cells.D2.value, 'only edit');
    assert.deepEqual(unzipSync(bytes)['custom/opaque.bin'], parts['custom/opaque.bin']);
  } finally {
    await book.close();
  }
});

test('source-backed results preserve cached fields, styles, branches, compaction and legacy full records', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>999</v></c><c r="C1" t="str"><f>"text"</f><v>old text</v></c><c r="D1" t="b"><f>TRUE</f><v>0</v></c></row>'
    )
  );
  let root;
  const book = await owner(source, async (input) => {
    if (input.operation === 'index') root = input.root;
    return native(input);
  });
  const { DatabaseSync } = await import('node:sqlite');
  const inspect = (work) => {
    const db = new DatabaseSync(path.join(root, 'cells.sqlite'));
    try {
      return work(db);
    } finally {
      db.close();
    }
  };
  try {
    const baseline = (await book.read('0', 'Sheet1', 'A1:D1')).cells;
    const records = inspect((db) =>
      db
        .prepare('SELECT data FROM changes ORDER BY r,c')
        .all()
        .map((r) => JSON.parse(r.data))
    );
    assert(records.some((c) => c.$sourceFields === true));
    for (const record of records.filter((c) => c.$sourceFields))
      for (const key of ['value', 'dataType', 'rawValue', 'cachedValue', 'styleId'])
        assert.equal(Object.hasOwn(record, key), false);
    assert.equal(baseline.B1.value, 999);
    assert.equal(baseline.B1.cachedValue, 999);
    assert.equal(baseline.B1.rawValue, '999');
    assert.equal(baseline.B1.calculation.value, 4);
    assert.equal(baseline.C1.value, 'old text');
    assert.equal(baseline.C1.calculation.value, 'text');
    assert.equal(baseline.D1.value, false);
    assert.equal(baseline.D1.calculation.value, true);
    assert(!Object.values(baseline).some((c) => Object.hasOwn(c, '$sourceFields')));
    const styled = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', bold: true }]);
    const changed = await book.apply(styled.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]);
    const branch = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]);
    await book.compact(['0', styled.view.storage.state, changed.view.storage.state, branch.view.storage.state]);
    for (const [state, value, bold] of [
      [styled.view.storage.state, 4, true],
      [changed.view.storage.state, 14, true],
      [branch.view.storage.state, 18, false],
    ]) {
      const cells = (await book.read(state, 'Sheet1', 'A1:D1')).cells;
      assert.equal(cells.B1.calculation.value, value);
      assert.equal(Boolean(cells.B1.bold), bold);
      const saved = readWorkbook(await fs.readFile((await book.export(state)).path)).sheets[0].cells;
      assert.equal(saved.B1.cachedValue, value);
      assert.equal(Boolean(saved.B1.bold), bold);
      assert.equal(saved.C1.cachedValue, 'text');
      assert.equal(saved.D1.cachedValue, true);
    }
    assert.deepEqual((await book.read('0', 'Sheet1', 'A1:D1')).cells, baseline);
    assert(
      inspect(
        (db) =>
          db.prepare("SELECT count(*) AS n FROM changes WHERE json_extract(data,'$.$sourceFields') IS NULL").get().n
      ) > 0,
      'Full materialized edits coexist with source-backed results'
    );
  } finally {
    await book.close();
  }
});

test('native patches cannot forge the private source-backed calculation marker', async () => {
  const { source } = await fixture(sheet('<row r="1"><c r="A1"><v>2</v></c></row>'));
  let root;
  const book = await owner(source, async (input) => {
    if (input.operation === 'index') root = input.root;
    return native(input);
  });
  try {
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'patch',
        root,
        version: 0,
        changes: [
          {
            sheet: book.view('0').sheets[0].path,
            cell: 'A1',
            value: { $sourceFields: true, calculation: { status: 'calculated', value: 99 } },
            edits: [{ value: 99 }],
          },
        ],
      }),
      { code: /^(formula-unsupported|style-unsupported)$/ }
    );
    assert.equal((await book.read('0', 'Sheet1', 'A1')).cells.A1.value, 2);
  } finally {
    await book.close();
  }
});

test('source-backed formula inputs preserve aggregate types and recalculate selected dependencies', async () => {
  const { source } = await fixture(
    sheet(
      '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>999</v></c><c r="C1" t="str"><f>"word"</f><v>cached</v></c><c r="D1" t="b"><f>TRUE</f><v>0</v></c><c r="E1" t="str"><f>""</f><v>old</v></c><c r="F1"><f>SUM(B1:E1)</f><v>0</v></c><c r="G1"><f>COUNT(B1:E1)</f><v>0</v></c><c r="H1"><f>COUNTA(B1:E1)</f><v>0</v></c><c r="I1"><f>COUNTBLANK(B1:E1)</f><v>0</v></c></row>'
    )
  );
  const book = await owner(source);
  const check = async (state, sum) => {
    const cells = (await book.read(state, 'Sheet1', 'B1:I1')).cells;
    assert.equal(cells.F1.calculation.value, sum);
    assert.equal(cells.G1.calculation.value, 1);
    assert.equal(cells.H1.calculation.value, 4);
    assert.equal(cells.I1.calculation.value, 1);
    assert.equal(cells.C1.value, 'cached');
    assert.equal(cells.D1.value, false);
    assert.equal(cells.E1.value, 'old');
    assert.equal(book.diagnosticStats().lastProjectedBytes, 0);
  };
  try {
    await check('0', 4);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 3 }]);
    await check(changed.view.storage.state, 6);
    const scalar = await book.apply(changed.view.storage.state, [
      { sheet: 'Sheet1', cell: 'G2', formula: 'LEN(C1)+A1' },
      { sheet: 'Sheet1', cell: 'H2', formula: 'IF(D1,A1,0)' },
    ]);
    const cells = (await book.read(scalar.view.storage.state, 'Sheet1', 'G2:H2')).cells;
    assert.equal(cells.G2.calculation.value, 7);
    assert.equal(cells.H2.calculation.value, 3);
    assert.equal(book.diagnosticStats().lastProjectedBytes, 0);
    const zero = await book.apply(scalar.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 0 }]);
    await check(zero.view.storage.state, 0);
    await check('0', 4);
  } finally {
    await book.close();
  }
});

test('sparse shared groups preserve ordinary overrides, literals and absent gaps through calculation, edits and export', async () => {
  const { source } = await fixture(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A5">B1+1</f><v>99</v></c><c r="B1"><v>2</v></c><c r="D1"><f>SUM(A1:A5)</f></c></row><row r="2"><c r="A2"><f>10+1</f><v>99</v></c><c r="B2"><v>3</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>99</v></c><c r="B3"><v>4</v></c></row><row r="4"><c r="A4"><v>7</v></c></row>'
      )
    ),
    original = await fs.readFile(source),
    book = await owner(source);
  try {
    assert.equal(book.view('0').readOnly, false);
    const initial = (await book.read('0', 'Sheet1', 'A1:D5')).cells;
    assert.equal(initial.A1.calculation.value, 3);
    assert.equal(initial.A2.calculation.value, 11);
    assert.equal(initial.A3.calculation.value, 5);
    assert.equal(initial.A4.value, 7);
    assert.equal(initial.A5, undefined);
    assert.equal(initial.D1.calculation.value, 26);
    const changed = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B3', value: 9 }]);
    const cells = (await book.read(changed.view.storage.state, 'Sheet1', 'A1:D5')).cells;
    assert.equal(cells.A3.calculation.value, 10);
    assert.equal(cells.D1.calculation.value, 31);
    const saved = await fs.readFile((await book.export(changed.view.storage.state)).path),
      disk = readWorkbook(saved).sheets[0].cells;
    assert.equal(disk.A2.formula, '10+1');
    assert.equal(disk.A4.value, 7);
    assert.equal(disk.A5, undefined);
    assert.equal(disk.A3.formulaType, 'shared');
    assert.equal(disk.A3.cachedValue, 10);
    assert.deepEqual(await fs.readFile(source), original);
  } finally {
    await book.close();
  }
});
