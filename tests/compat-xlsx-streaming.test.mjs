import test from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook, editWorkbookResult } from '../core/workbook-package.mjs';
import { WorkbookSession, readWorkbookRange } from '../core/workbook-service.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const variant = (worksheet) => {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(worksheet);
  return parts;
};
const legacy = (bytes) => editWorkbookResult(bytes, [], undefined, true, undefined, false, { dom: true }).view;

test('streaming reader and targeted writer retain prefixed XML, rich strings, phonetics and exact text', async () => {
  const parts = variant(
    `<s:worksheet xmlns:s="${NS}" xmlns:z="urn:opaque"><s:sheetData><s:row r="1" z:keep="row"><s:c z:keep="cell" r="A1" t="s"><s:v>0</s:v><s:extLst><z:opaque z:attr="x &gt; y"/></s:extLst><!--retain--></s:c><s:c r="B1" t="inlineStr"><s:is><s:r><s:t>_x00</s:t></s:r><s:rPh><s:t>ignored</s:t></s:rPh><s:r><s:t>41_</s:t></s:r></s:is></s:c><s:c r="C1" t="e"><s:v>#VALUE!</s:v></s:c><s:c r="D1" t="inlineStr"><s:is><s:t>before&#13;after_x005F_x0041_</s:t></s:is></s:c><s:c r="E1" t="inlineStr"><s:is><s:t><![CDATA[old </s:c> >]]></s:t></s:is></s:c><s:c r="F1"/></s:row></s:sheetData><s:extLst><z:data><![CDATA[<s:c r="A1">not a cell</s:c>]]></z:data></s:extLst></s:worksheet>`
  );
  parts['xl/sharedStrings.xml'] = strToU8(
    `<s:sst xmlns:s="${NS}" count="1" uniqueCount="1"><s:si><s:r><s:t>_x00</s:t></s:r><s:rPh><s:t>phonetic</s:t></s:rPh><s:r><s:t>41_</s:t></s:r><s:r><s:t> café&#13;_x0002_</s:t></s:r></s:si></s:sst>`
  );
  const original = zipSync(parts),
    before = readWorkbook(original);
  assert.deepEqual(before, legacy(original));
  assert.equal(before.sheets[0].cells.A1.value, '_x0041_ café\r\u0002');
  assert.equal(before.sheets[0].cells.B1.value, '_x0041_');
  assert.equal(before.sheets[0].cells.F1.dataType, 'blank');
  assert.equal(before.sheets[0].cells.D1.value, 'before\rafter_x0041_');
  const result = editWorkbookResult(
    original,
    [{ sheet: 'Sheet1', cell: 'E1', expectedValue: 'old </s:c> >', value: 'new\rtext' }],
    before.revision
  );
  assert.deepEqual(readWorkbook(result.bytes), legacy(result.bytes));
  const saved = unzipSync(result.bytes),
    xml = strFromU8(saved['xl/worksheets/sheet1.xml']);
  for (const literal of ['z:keep="row"', 'z:keep="cell"', '<!--retain-->', '<![CDATA[<s:c r="A1">not a cell</s:c>]]>'])
    assert(xml.includes(literal), literal);
  for (const name of Object.keys(parts))
    if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(name))
      assert.deepEqual(saved[name], parts[name], name);
  assert.equal(readWorkbook(result.bytes).sheets[0].cells.E1.value, 'new\rtext');
  assert.equal(readWorkbook(result.bytes).sheets[0].cells.C1.dataType, 'error');
});

test('targeted insertion retains empty-row attributes, ordered new cells and metadata', () => {
  const parts = variant(
    `<w:worksheet xmlns:w="${NS}"><w:dimension ref="B2:D7"/><w:sheetData><w:row r="2" customHeight="1" ht="27"><w:c r="B2"><w:v>2</w:v></w:c><w:c r="D2"><w:v>4</w:v></w:c></w:row><w:row r="7" customHeight="1" ht="31"/></w:sheetData></w:worksheet>`
  );
  const original = zipSync(parts),
    commands = [
      { cell: 'A1', value: 1 },
      { cell: 'C2', value: 3 },
      { cell: 'E2', value: 5 },
      { cell: 'A7', value: 7 },
      { cell: 'A9', value: 9 },
    ].map((edit) => ({ sheet: 'Sheet1', ...edit }));
  const result = editWorkbookResult(original, commands),
    view = readWorkbook(result.bytes);
  assert.deepEqual(view, legacy(result.bytes));
  for (const edit of commands) assert.equal(view.sheets[0].cells[edit.cell].value, edit.value);
  assert.deepEqual(view.sheets[0].rowHeights, [
    { row: 1, height: 27 },
    { row: 6, height: 31 },
  ]);
  const xml = strFromU8(unzipSync(result.bytes)['xl/worksheets/sheet1.xml']);
  assert(!xml.includes(':dimension'));
  assert(xml.includes('customHeight="1" ht="31"'));
  assert(xml.indexOf('r="B2"') < xml.indexOf('r="C2"') && xml.indexOf('r="C2"') < xml.indexOf('r="D2"'));
});

test('bounded queries and key resolution accept a full valid worksheet area without enumerating it', () => {
  const view = readWorkbook(createWorkbook());
  view.sheets[0].cells = { A1: { value: 'first' }, A1048576: { value: 'last' }, XFD1048576: { value: 42 } };
  const result = readWorkbookRange(view, { sheet: 'Sheet1', range: 'A1:XFD1048576', offset: 17179869183, limit: 1 });
  assert.equal(result.total, 17179869184);
  assert.equal(result.items[0].cell, 'XFD1048576');
  assert.equal(result.items[0].value, 42);
  const keyed = readWorkbookRange(view, {
    sheet: 'Sheet1',
    range: 'A1:XFD1048576',
    keys: { column: 'A', values: ['last'] },
    limit: 2,
  });
  assert.equal(keyed.total, 16384);
  assert.equal(keyed.items[0].cell, 'A1048576');
  // Blank keys are duplicates in this deliberately sparse range and must refuse,
  // rather than silently selecting a row. The narrower real-data case follows.
  assert.throws(
    () =>
      resolveWorkbookEdits(view, [
        {
          type: 'updateRows',
          sheet: 'Sheet1',
          range: 'A1:XFD1048576',
          keyColumn: 'A',
          updates: [{ key: 'last', cells: [{ column: 'XFD', expectedValue: 42, value: 43 }] }],
        },
      ]),
    { code: /^(csv-duplicate|duplicates-duplicate|workbook-duplicate)$/ }
  );
  view.sheets[0].cells = Object.fromEntries(
    Array.from({ length: 10001 }, (_, row) => [`A${row + 1}`, { value: String(row + 1) }])
  );
  view.sheets[0].cells.XFD10001 = { value: 42 };
  const edits = resolveWorkbookEdits(view, [
    {
      type: 'updateRows',
      sheet: 'Sheet1',
      range: 'A1:XFD10001',
      keyColumn: 'A',
      updates: [{ key: '10001', cells: [{ column: 'XFD', expectedValue: 42, value: 43 }] }],
    },
  ]);
  assert.equal(edits[0].cell, 'XFD10001');
});

test('sequential guarded edits and keyed updates use current raw values without changing accepted input', async () => {
  const source = editWorkbookResult(createWorkbook(), [
    { sheet: 'Sheet1', cell: 'A1', value: 1 },
    { sheet: 'Sheet1', cell: 'B1', value: 5 },
  ]).bytes;
  const session = new WorkbookSession(source),
    before = await session.read();
  const result = await session.apply({
    operationId: 'sequential-key',
    expectedRevision: before.revision,
    commands: [
      { sheet: 'Sheet1', cell: 'A1', expectedRawValue: '1', value: 2 },
      {
        type: 'updateRows',
        sheet: 'Sheet1',
        range: 'A1:B1',
        keyColumn: 'A',
        updates: [{ key: 2, cells: [{ column: 'B', expectedValue: 5, value: 6 }] }],
      },
      { sheet: 'Sheet1', cell: 'A1', expectedRawValue: '2', value: 3 },
    ],
  });
  assert.equal(result.view.sheets[0].cells.A1.value, 3);
  assert.equal(result.view.sheets[0].cells.B1.value, 6);
  assert.equal(before.sheets[0].cells.A1.value, 1);
  assert.equal(before.sheets[0].cells.A1.rawValue, '1');
  const zeroView = readWorkbook(createWorkbook());
  zeroView.sheets[0].cells = { A1: { value: -0 }, A2: { value: 0 } };
  const query = readWorkbookRange(zeroView, { sheet: 'Sheet1', range: 'A1:A2', keys: { column: 'A', values: [-0] } });
  assert.equal(query.total, 1);
  assert.equal(query.items[0].cell, 'A1');
});

test('opaque cell-like extension payload stays metadata during literal and formula cache writes', async () => {
  const extension =
    '<s:extLst><s:ext uri="test"><z:row><s:c r="Z99"><s:f>OPAQUE()</s:f><s:v>7</s:v><s:is><s:t>opaque text</s:t></s:is></s:c></z:row></s:ext></s:extLst>';
  const parts = variant(
    `<s:worksheet xmlns:s="${NS}" xmlns:z="urn:opaque"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v>${extension}</s:c><s:c r="B1"><s:f>A1*2</s:f><s:v>2</s:v>${extension}</s:c></s:row></s:sheetData></s:worksheet>`
  );
  const original = zipSync(parts),
    session = new WorkbookSession(original),
    before = await session.read();
  assert.deepEqual(Object.keys(before.sheets[0].cells), ['A1', 'B1']);
  const result = await session.apply({
    operationId: 'opaque-cell-extension',
    expectedRevision: before.revision,
    commands: [{ sheet: 'Sheet1', cell: 'A1', expectedValue: 1, value: 3 }],
  });
  assert.equal(result.view.sheets[0].cells.A1.value, 3);
  assert.equal(result.view.sheets[0].cells.B1.value, 6);
  assert.deepEqual(Object.keys(result.view.sheets[0].cells), ['A1', 'B1']);
  const saved = unzipSync(result.bytes),
    xml = strFromU8(saved['xl/worksheets/sheet1.xml']);
  assert(!xml.includes('??'));
  assert.equal((xml.match(/r="Z99"/g) ?? []).length, 2);
  assert.equal((xml.match(/OPAQUE\(\)/g) ?? []).length, 2);
  assert.equal((xml.match(/<s:v>7<\/s:v>/g) ?? []).length, 2);
  assert.equal((xml.match(/opaque text/g) ?? []).length, 2);
  assert.equal(
    (xml.match(/<s:v>[36]<\/s:v><s:extLst>/g) ?? []).length,
    2,
    'authored value and calculated cache precede extLst'
  );
  assert.deepEqual(readWorkbook(result.bytes), legacy(result.bytes));
  assert.equal(before.sheets[0].cells.A1.value, 1);
  assert.equal(before.sheets[0].cells.B1.value, 2);
  for (const name of Object.keys(parts))
    if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(name))
      assert.deepEqual(saved[name], parts[name], name);
});

test('opaque worksheet-like metadata cannot redirect insertion spans or dimension cleanup', async () => {
  const opaque =
    '<s:extLst><s:ext uri="test"><z:data><s:dimension ref="Z99"/><s:sheetData><s:row r="1"><s:c r="B1"><s:v>7</s:v></s:c></s:row></s:sheetData></z:data></s:ext></s:extLst>';
  const parts = variant(
    `<s:worksheet xmlns:s="${NS}" xmlns:z="urn:opaque"><s:dimension ref="A1"/><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></s:sheetData>${opaque}</s:worksheet>`
  );
  const original = zipSync(parts),
    session = new WorkbookSession(original),
    before = await session.read();
  const result = await session.apply({
    operationId: 'opaque-sheet-location',
    expectedRevision: before.revision,
    commands: [
      { sheet: 'Sheet1', cell: 'B1', value: 3 },
      { sheet: 'Sheet1', cell: 'A2', value: 4 },
    ],
  });
  const reopened = readWorkbook(result.bytes),
    xml = strFromU8(unzipSync(result.bytes)['xl/worksheets/sheet1.xml']);
  assert.equal(result.view.sheets[0].cells.B1.value, 3);
  assert.equal(reopened.sheets[0].cells.B1.value, 3);
  assert.equal(reopened.sheets[0].cells.A2.value, 4);
  assert.deepEqual(Object.keys(reopened.sheets[0].cells), ['A1', 'B1', 'A2']);
  assert(xml.includes(opaque), 'opaque dimension, data and row bytes remain exact');
  assert(!xml.includes('<s:dimension ref="A1"/>'), 'only actual sheet dimension is invalidated');
  assert.deepEqual(reopened, legacy(result.bytes));
  assert.equal(before.sheets[0].cells.B1, undefined);
});
