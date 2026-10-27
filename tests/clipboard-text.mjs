import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { parseDelimited } from '../core/delimited.ts';
import { parseCellInput } from '../core/cell-input.mjs';
import { inspectPaste, pastePages } from '../ui/largePaste.ts';
import { PASTE_BYTE_LIMIT, pasteBytes } from '../core/paste-protocol.mjs';

// Exercise the actual browser bundle, including extensionless UI source imports.
const bundle = await build({
  entryPoints: [fileURLToPath(new URL('../ui/clipboardText.ts', import.meta.url))],
  bundle: true,
  write: false,
  platform: 'browser',
  format: 'esm',
  logLevel: 'silent',
});
const { serializeClipboardText, literalClipboardCells, CLIPBOARD_TEXT_POSITION_LIMIT } = await import(
  'data:text/javascript;base64,' + Buffer.from(bundle.outputFiles[0].contents).toString('base64')
);

const target = { sheet: "O'Brien Data", start: 'B3' };
function streamed(text, options = target) {
  const inspection = inspectPaste(text, { ...options, policy: 'text' });
  return { inspection, entries: [...pastePages(text, inspection)].flatMap((page) => page.cells) };
}

test('quoted TSV roundtrips separators, whitespace, Unicode and exact field contents through both parsers', () => {
  const rows = Object.freeze([
    Object.freeze(['tab\tinside', 'line\nnext', 'CR\rnext', 'CRLF\r\nnext']),
    Object.freeze(['"quoted"', '  spaces  ', '😀é', "'apostrophe"]),
    Object.freeze(['=SUM(A1)', 'TRUE', '00123', '9007199254740993']),
    Object.freeze(['#N/A', '1.0000000000000002', '12.5', '']),
  ]);
  const expected =
    '"tab\tinside"\t"line\nnext"\t"CR\rnext"\t"CRLF\r\nnext"\n' +
    '"""quoted"""\t"  spaces  "\t"😀é"\t"\'apostrophe"\n' +
    '"=SUM(A1)"\t"TRUE"\t"00123"\t"9007199254740993"\n' +
    '"#N/A"\t"1.0000000000000002"\t"12.5"\t""';
  const text = serializeClipboardText(rows);
  assert.equal(text, expected);
  assert.deepEqual(parseDelimited(text, '\t').rows, rows);
  const { inspection, entries } = streamed(text);
  assert.deepEqual(inspection.shape, { sheet: target.sheet, start: 'B3', rows: 4, columns: 4 });
  assert.equal(inspection.count, 16);
  assert.deepEqual(
    entries,
    rows.flatMap((row, r) => row.map((value, c) => [r * 4 + c, value]))
  );
  const result = literalClipboardCells(text, target);
  assert.equal(result.range, 'B3:E6');
  assert.deepEqual(result.cells, [
    { cell: 'B3', value: 'tab\tinside' },
    { cell: 'C3', value: 'line\nnext' },
    { cell: 'D3', value: 'CR\rnext' },
    { cell: 'E3', value: 'CRLF\r\nnext' },
    { cell: 'B4', value: '"quoted"' },
    { cell: 'C4', value: '  spaces  ' },
    { cell: 'D4', value: '😀é' },
    { cell: 'E4', value: "'apostrophe" },
    { cell: 'B5', value: '=SUM(A1)' },
    { cell: 'C5', value: 'TRUE' },
    { cell: 'D5', value: '00123' },
    { cell: 'E5', value: '9007199254740993' },
    { cell: 'B6', value: '#N/A' },
    { cell: 'C6', value: '1.0000000000000002' },
    { cell: 'D6', value: '12.5' },
    { cell: 'E6', value: '' },
  ]);
});

test('one copied blank cell and all final empty fields/rows remain explicit without a terminal separator', () => {
  for (const [rows, expected] of [
    [[['']], '""'],
    [[['', '']], '""\t""'],
    [[[''], [''], ['']], '""\n""\n""'],
    [
      [
        ['value', ''],
        ['', ''],
        ['', ''],
      ],
      '"value"\t""\n""\t""\n""\t""',
    ],
  ]) {
    const text = serializeClipboardText(rows);
    assert.equal(text, expected);
    assert(!text.endsWith('\n'));
    assert(!text.endsWith('\t'));
    assert.deepEqual(parseDelimited(text, '\t').rows, rows);
    const { inspection, entries } = streamed(text);
    assert.equal(inspection.shape.rows, rows.length);
    assert.equal(entries.length, rows.length * rows[0].length);
    assert.deepEqual(
      literalClipboardCells(text, target).cells.map((cell) => cell.value),
      rows.flat()
    );
  }
  assert.deepEqual(literalClipboardCells('""', target), { range: 'B3', cells: [{ cell: 'B3', value: '' }] });
  // Existing external-TSV terminal-newline semantics are retained. The UI owns empty-OS-clipboard no-op.
  for (const text of ['', '\n', '\r\n', '\n\n', 'a\t\n']) {
    const rows = parseDelimited(text, '\t').rows,
      { inspection, entries } = streamed(text);
    assert.equal(inspection.shape.rows, rows.length);
    assert.equal(entries.length, rows.flat().length);
    assert.deepEqual(
      literalClipboardCells(text, target).cells.map((cell) => cell.value),
      rows.flat()
    );
  }
});

test('ragged pasted rows explicitly clear rectangular holes while raw staging stays sparse', () => {
  const text = 'a\tb\nshort\n\t';
  assert.deepEqual(parseDelimited(text, '\t').rows, [['a', 'b'], ['short'], ['', '']]);
  const { inspection, entries } = streamed(text);
  assert.equal(inspection.count, 5);
  assert.deepEqual(entries, [
    [0, 'a'],
    [1, 'b'],
    [2, 'short'],
    [4, ''],
    [5, ''],
  ]);
  assert.deepEqual(literalClipboardCells(text, target), {
    range: 'B3:C5',
    cells: [
      { cell: 'B3', value: 'a' },
      { cell: 'C3', value: 'b' },
      { cell: 'B4', value: 'short' },
      { cell: 'C4', value: '' },
      { cell: 'B5', value: '' },
      { cell: 'C5', value: '' },
    ],
  });
});

test('literal policy preserves apostrophes, formulas, booleans, numbers, errors and precision strings as strings', () => {
  const fields = [
    "'00123",
    '=SUM(A1)',
    'TRUE',
    'false',
    '00123',
    '12.5',
    '-0',
    '#REF!',
    '9007199254740993',
    '1.0000000000000002',
    '1e999',
    ' 12 ',
    '',
  ];
  const result = literalClipboardCells(serializeClipboardText([fields]), { sheet: 'Sheet1', start: 'A1' });
  assert.deepEqual(
    result.cells.map((cell) => cell.value),
    fields
  );
  assert(result.cells.every((cell) => typeof cell.value === 'string' && !Object.hasOwn(cell, 'formula')));
  assert.deepEqual(parseCellInput('00123'), { value: '00123' });
  assert.deepEqual(parseCellInput('9007199254740993'), { value: '9007199254740993' });
  assert.deepEqual(parseCellInput('=SUM(A1)'), { formula: 'SUM(A1)' });
  assert.deepEqual(parseCellInput('TRUE'), { value: true });
  assert.deepEqual(parseCellInput('12.5'), { value: 12.5 });
});

test('10,000 positions, ragged footprint and final worksheet coordinates are bounded before returning edits', () => {
  assert.equal(CLIPBOARD_TEXT_POSITION_LIMIT, 10_000);
  const rows = Array.from({ length: 1000 }, (_, r) => Array.from({ length: 10 }, (_, c) => `${r}:${c}`));
  const text = serializeClipboardText(rows),
    { inspection, entries } = streamed(text, { sheet: 'Sheet1', start: 'A1' });
  assert.equal(inspection.count, 10000);
  assert.equal(entries.length, 10000);
  const result = literalClipboardCells(text, { sheet: 'Sheet1', start: 'A1' });
  assert.equal(result.range, 'A1:J1000');
  assert.deepEqual(result.cells.at(-1), { cell: 'J1000', value: '999:9' });
  assert.deepEqual(literalClipboardCells('""', { sheet: 'Sheet1', start: 'XFD1048576' }), {
    range: 'XFD1048576',
    cells: [{ cell: 'XFD1048576', value: '' }],
  });
  assert.equal(
    literalClipboardCells('a\tb\tc\nx\ty\tz', { sheet: 'Sheet1', start: 'XFB1048575' }).range,
    'XFB1048575:XFD1048576'
  );
  assert.throws(() => serializeClipboardText(Array.from({ length: 10001 }, () => ['x'])), /10,000/);
  assert.throws(() => literalClipboardCells('x\ty\n' + 'short\n'.repeat(5000), target), /10,000/);
  assert.throws(() => literalClipboardCells('x\ty', { sheet: 'Sheet1', start: 'XFD1' }), /limits/);
  assert.throws(() => literalClipboardCells('x\ny', { sheet: 'Sheet1', start: 'A1048576' }), /limits/);
});

test('invalid rectangles, malformed TSV, invalid XML and owner-length fields refuse without mutating inputs', () => {
  const rowHole = new Array(1),
    fieldHole = new Array(2);
  fieldHole[0] = 'a';
  for (const rows of [[], [[]], [['a'], ['b', 'c']], rowHole, [fieldHole], [[12]], [[null]]])
    assert.throws(() => serializeClipboardText(rows), /rectangle|text/);
  for (const text of ['"unfinished', '"closed" suffix', '\u0000', '\ud800', '\udc00'])
    assert.throws(() => literalClipboardCells(text, target), /Malformed|XML/);
  for (const start of ['b3', '$B$3', 'B0'])
    assert.throws(() => literalClipboardCells('x', { sheet: 'Sheet1', start }), /origin|limits|cell/i);
  assert.throws(() => literalClipboardCells('x', { ...target, extra: true }), /sheet and start/);
  const longest = 'x'.repeat(32767);
  assert.equal(literalClipboardCells(serializeClipboardText([[longest]]), target).cells[0].value, longest);
  assert.throws(() => serializeClipboardText([['x'.repeat(32768)]]), /32767/);
  assert.throws(() => literalClipboardCells('x'.repeat(32768), target), /32767/);
  const unchanged = Object.freeze([Object.freeze(['before', 'after'])]);
  serializeClipboardText(unchanged);
  assert.deepEqual(unchanged, [['before', 'after']]);
});

test('64 MiB accounting uses actual UTF8 and includes quoting expansion', () => {
  for (const value of ['ASCII', 'é', '😀', 'é', '"\t\r\n']) {
    const text = serializeClipboardText([[value]]);
    assert.equal(pasteBytes(text), Buffer.byteLength(text, 'utf8'));
  }
  // All individual fields are admitted; only total encoded bytes exceed the bound.
  const multibyte = 'é'.repeat(32767),
    quotes = '"'.repeat(32767);
  assert.equal(Buffer.byteLength(multibyte), 65534);
  assert.throws(() => serializeClipboardText(Array.from({ length: 1024 }, () => [multibyte])), /64 MiB/);
  assert.throws(() => serializeClipboardText(Array.from({ length: 1024 }, () => [quotes])), /64 MiB/);
  const over = 'é'.repeat(PASTE_BYTE_LIMIT / 2 + 1);
  assert(over.length < PASTE_BYTE_LIMIT);
  assert.throws(() => literalClipboardCells(over, target), /64 MiB/);
});
