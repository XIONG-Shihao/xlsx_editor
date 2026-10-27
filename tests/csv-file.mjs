import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync, readFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { gunzipSync } from 'node:zlib';
import { editCsvFile, writeCsvIndex } from '../core/csv-file-save.mjs';
import { readCsvFileTransfer, readCsvFileIndex } from '../core/csv-file.mjs';
import { readCsvWorkbook, editCsvBytes } from '../core/csv-document.mjs';
import { encodeCsvText } from '../core/csv-codec.mjs';
import {
  indexCsv,
  indexCsvChunks,
  csvReadRange,
  applyCsvCommands,
  csvTextEdits,
  csvCacheStatistics,
} from '../core/csv-workbook.mjs';

const cases = ['', 'a,b\r\n1,2\r\n', '"a""b",c\r\n"line\nline",🙂', '\r\n\n\r', ',,\n"",x,', '""', '"a"""'];
const metadata = ({ text, ...rest }) => rest;
test('long spans and short-field records preserve exact UTF-16 seams and malformed offsets', () => {
  const records = [
    ['x'.repeat(8193), '🙂'.repeat(4097), 'line\n"quote"', ''],
    Array(1000).fill('1'),
    ['tail', '中\r\nend'],
  ];
  for (const delimiter of [',', ';', '\t']) {
    const lines = records.map((row) => row.map((value) => '"' + value.replaceAll('"', '""') + '"').join(delimiter));
    const text = lines[0] + '\r\n' + lines[1] + '\n' + lines[2],
      expectedOffsets = [0, lines[0].length + 2, lines[0].length + 2 + lines[1].length + 1, text.length];
    for (const size of [1, 2, 7, 4095, 4096, 4097, 65536]) {
      function* chunks() {
        for (let i = 0; i < text.length; i += size) yield text.slice(i, i + size);
      }
      const index = { ...indexCsvChunks(chunks(), { delimiter }), text };
      assert.deepEqual([...index.rowOffsets], expectedOffsets);
      assert.deepEqual([...index.rowWidths], [4, 1000, 2]);
      assert.equal(index.nonemptyCells, 1005);
      assert.equal(index.newline, '\r\n');
      assert.equal(index.terminalNewline, false);
      assert.deepEqual(
        csvReadRange(index, 'A1:D1').items.map((c) => c.value),
        records[0]
      );
      assert.deepEqual(
        csvReadRange(index, 'A3:B3').items.map((c) => c.value),
        records[2]
      );
    }
  }
  for (const size of [1, 2, 4096, 16384]) {
    const text = '"ok"' + 'x'.repeat(8192);
    function* chunks() {
      for (let i = 0; i < text.length; i += size) yield text.slice(i, i + size);
    }
    assert.throws(() => indexCsvChunks(chunks()), /closing quote near offset 4$/);
  }
});
for (const text of cases)
  test(`chunk boundaries: ${JSON.stringify(text)}`, () => {
    const expected = metadata(indexCsv(text));
    for (let size = 1; size <= text.length + 1; size++) {
      function* chunks() {
        for (let at = 0; at < text.length; at += size) {
          yield '';
          yield text.slice(at, at + size);
        }
      }
      assert.deepEqual(indexCsvChunks(chunks()), expected);
    }
  });
test('known quote, CRLF and Unicode offsets retain literal cells', () => {
  const text = 'id;"a""b"\r\n001;"x\ny"\r\n🙂;""';
  const index = indexCsv(text, { delimiter: ';' });
  assert.deepEqual([...index.rowOffsets], [0, 11, 22, 27]);
  assert.deepEqual([...index.rowWidths], [2, 2, 2]);
  assert.equal(index.nonemptyCells, 5);
  assert.deepEqual(
    csvReadRange(index, 'A1:B3').items.map((x) => x.value),
    ['id', 'a"b', '001', 'x\ny', '🙂', '']
  );
});
test('invalid quote state rejects across every boundary', () => {
  assert.throws(() => indexCsv('a,b', { delimiter: null }), { code: 'csv-unsupported' });
  for (const text of ['"a', '"a"x', 'a,"b\n'])
    for (let n = 1; n <= text.length; n++) {
      const chunks = Array.from({ length: Math.ceil(text.length / n) }, (_, i) => text.slice(i * n, (i + 1) * n));
      assert.throws(() => indexCsvChunks(chunks), { code: 'csv-malformed' });
    }
});
test('file streaming preserves BOM, encoding, delimiter, text and index across byte boundaries', () => {
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-stream-'));
  try {
    const file = join(dir, 'test.csv');
    for (const encoding of ['utf-8', 'utf-16le', 'utf-16be'])
      for (const bom of [false, true]) {
        if (encoding !== 'utf-8' && !bom) continue;
        // Header exceeds a read block, escaped quotes/newlines and emoji cross
        // physical byte boundaries, and source characters remain UTF-16 indexed.
        for (const shift of [0, 1, 2, 3]) {
          const text = '"' + 'x'.repeat(1024 * 1024 - 4 + shift) + '🙂\n""tail";b\r\n001;中\r\n';
          const bytes = encodeCsvText(text, encoding, bom);
          writeFileSync(file, bytes);
          const expected = readCsvWorkbook(bytes),
            actual = readCsvFileTransfer(file, {}, expected.revision);
          const backed = readCsvFileIndex(file, {}, expected.revision);
          assert.deepEqual(metadata(backed), metadata(expected));
          for (const at of [0, 1024 * 1024 - 8, 1024 * 1024 - 2, text.length - 10])
            assert.equal(backed.text.slice(at, at + 12), text.slice(at, at + 12));
          const { transport, compressedText, blocks, textOffsets, ...view } = actual;
          assert.deepEqual(view, metadata(expected));
          const decoded = (blocks ?? [compressedText]).map((b) => gunzipSync(b).toString('utf8'));
          assert.equal(decoded.join(''), text);
          if (textOffsets) {
            assert.deepEqual(
              [...textOffsets],
              [0, ...decoded.map((_, i) => decoded.slice(0, i + 1).reduce((n, x) => n + x.length, 0))]
            );
            assert(decoded.every((x) => x.length <= 1024 * 1024 + 1));
          }
        }
      }
    writeFileSync(file, 'a,b\r\n1,2\r\n');
    assert.throws(() => readCsvFileTransfer(file, {}, 'stale'), { code: 'csv-stale' });
    for (const bytes of [
      Buffer.from([0xff]),
      Buffer.from([0xff, 0xfe, 0x61]),
      Buffer.from([0xef, 0xbb, 0xbf, 0xe2]),
      Buffer.from([0xff, 0xfe, 0, 0]),
    ]) {
      writeFileSync(file, bytes);
      assert.throws(() => readCsvFileTransfer(file), { code: 'csv-unsupported' });
    }
    writeFileSync(file, '');
    const empty = readCsvFileTransfer(file);
    assert.equal(empty.rows, 1);
    assert.equal(empty.revision, createHash('sha256').update('').digest('hex'));
    assert.equal(gunzipSync(empty.compressedText).length, 0);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('file-backed edits match byte writer across encodings, sparse gaps, deletion and append', () => {
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-save-'));
  try {
    const source = join(dir, 'source.csv'),
      target = join(dir, 'saved.csv');
    for (const encoding of ['utf-8', 'utf-16le', 'utf-16be'])
      for (const bom of [false, true]) {
        if (encoding !== 'utf-8' && !bom) continue;
        for (const terminal of ['', '\r\n']) {
          const text = 'id;value\r\n1;"🙂中文\r\nquote""x"\n2;keep\r3;tail' + terminal;
          const bytes = encodeCsvText(text, encoding, bom);
          writeFileSync(source, bytes);
          const base = readCsvWorkbook(bytes, { delimiter: ';' }),
            index = readCsvFileIndex(source, { delimiter: ';' });
          assert.deepEqual(metadata(index), metadata(base));
          assert.equal(index.text.slice(), text);
          assert.equal(index.text.slice(-5), text.slice(-5));
          for (let at = 0; at < text.length; at++) assert.equal(index.text.slice(at, at + 1), text.slice(at, at + 1));
          const prefix = bom ? (encoding === 'utf-8' ? 3 : 2) : 0;
          for (const at of [0, 1, 2, text.indexOf('中文'), text.length, text.indexOf('中文'), 1]) {
            const expected = prefix + (encoding === 'utf-8' ? Buffer.byteLength(text.slice(0, at), 'utf8') : at * 2);
            assert.equal(index.text.byteOffset(at), expected);
          }
          const commands = [
            { sheet: 'CSV', cell: 'B2', value: 'new;🙂' },
            { sheet: 'CSV', cell: 'B4', value: '終' },
            { sheet: 'CSV', cell: 'C6', value: 'last' },
          ];
          const expected = editCsvBytes(bytes, commands, base.revision, { delimiter: ';' });
          const result = editCsvFile(source, commands, base.revision, { delimiter: ';' }, target);
          assert.deepEqual(readFileSync(target), Buffer.from(expected.bytes));
          assert.equal(result.revision, expected.revision);
          assert.deepEqual(readFileSync(source), Buffer.from(bytes));
          rmSync(target);
          for (const edits of [[], [{ sheet: 'CSV', cell: 'B3', value: 'keep' }]]) {
            editCsvFile(source, edits, base.revision, { delimiter: ';' }, target);
            assert.deepEqual(readFileSync(target), Buffer.from(bytes));
            rmSync(target);
          }
          assert.throws(() => editCsvFile(source, commands, 'stale', {}, target), /stale/);
          assert(!existsSync(target));
          assert.throws(
            () =>
              editCsvFile(
                source,
                [...commands, { sheet: 'CSV', cell: 'B3', expectedValue: 'wrong', value: 'bad' }],
                base.revision,
                { delimiter: ';' },
                target
              ),
            /Conflict/
          );
          assert(!existsSync(target));
        }
      }
    for (const text of ['', 'a\n' + '🙂'.repeat(600000) + '\ntail\n', 'a,b\r\n' + 'x,y\r\n'.repeat(180000)]) {
      writeFileSync(source, text);
      const bytes = readFileSync(source),
        base = readCsvWorkbook(bytes);
      const commands = [
        { sheet: 'CSV', cell: 'A2', value: 'short' },
        { sheet: 'CSV', cell: 'B180004', value: 'last' },
      ];
      const result = editCsvFile(source, commands, base.revision, {}, target),
        expected = editCsvBytes(bytes, commands, base.revision);
      assert.deepEqual(readFileSync(target), Buffer.from(expected.bytes));
      assert.equal(result.revision, expected.revision);
      rmSync(target);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('existing destinations and changed snapshots survive rejected file saves', () => {
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-save-guards-'));
  try {
    const source = join(dir, 'source.csv'),
      target = join(dir, 'saved.csv');
    writeFileSync(source, 'a,b\n1,2\n');
    const index = readCsvFileIndex(source),
      edited = applyCsvCommands(index, [{ sheet: 'CSV', cell: 'B2', value: '3' }]).index;
    writeFileSync(target, 'keep destination');
    assert.throws(() => writeCsvIndex(edited, source, target), /exist/i);
    assert.equal(readFileSync(target, 'utf8'), 'keep destination');
    rmSync(target);
    writeFileSync(source, 'a,b\n8,2\n');
    assert.throws(() => writeCsvIndex(edited, source, target), { code: 'csv-stale' });
    assert(!existsSync(target));
    assert.equal(readFileSync(source, 'utf8'), 'a,b\n8,2\n');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('buffered CSV output handles partial writes and detects source mutation during final flush', async () => {
  const { default: fs } = await import('node:fs'),
    { syncBuiltinESMExports } = await import('node:module');
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-buffered-')),
    source = join(dir, 'source.csv'),
    target = join(dir, 'target.csv'),
    write = fs.writeSync;
  const text = 'id,value\r\n' + Array.from({ length: 1000 }, (_, i) => `${i},${'old'.repeat(24)}\r\n`).join(''),
    bytes = Buffer.from(text);
  const commands = Array.from({ length: 1000 }, (_, i) => ({
    sheet: 'CSV',
    cell: `B${i + 2}`,
    value: i % 7 ? 'changed' + i : 'quoted,\n🙂',
  }));
  try {
    writeFileSync(source, bytes);
    const expected = editCsvBytes(bytes, commands);
    fs.writeSync = (fd, data, offset, length, ...rest) => write(fd, data, offset, Math.min(length, 31), ...rest);
    syncBuiltinESMExports();
    editCsvFile(source, commands, undefined, {}, target);
    assert.deepEqual(readFileSync(target), Buffer.from(expected.bytes));
    assert.deepEqual(readFileSync(source), bytes);
    fs.writeSync = write;
    syncBuiltinESMExports();
    rmSync(target);
    let changed = false;
    fs.writeSync = (...args) => {
      const n = write(...args);
      if (!changed) {
        changed = true;
        writeFileSync(source, Buffer.concat([bytes, Buffer.from('tampered')]));
      }
      return n;
    };
    syncBuiltinESMExports();
    assert.throws(() => editCsvFile(source, commands, undefined, {}, target), /changed/);
    assert(changed);
    assert(!existsSync(target));
    fs.writeSync = write;
    syncBuiltinESMExports();
    writeFileSync(source, bytes);
    fs.writeSync = () => 0;
    syncBuiltinESMExports();
    assert.throws(() => editCsvFile(source, commands, undefined, {}, target), { code: 'csv-unsupported' });
    assert(!existsSync(target));
  } finally {
    fs.writeSync = write;
    syncBuiltinESMExports();
    rmSync(dir, { recursive: true, force: true });
  }
});

test('dense field-span saves preserve untouched bytes and coalesce adjacent edited fields', () => {
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-field-spans-'));
  try {
    const source = join(dir, 'source.csv'),
      target = join(dir, 'saved.csv');
    const rows = Array.from({ length: 160 }, (_, r) => Array.from({ length: 80 }, (_, c) => `"row${r};column${c}"`));
    const text = rows.map((r) => r.join(',')).join('\r\n') + '\r\n',
      commands = [];
    const expected = rows.map((r) => r.slice());
    for (let row = 0; row < rows.length; row++) {
      for (const [column, cell, value] of [
        [1, 'B', 'left,quoted'],
        [2, 'C', '🙂right'],
        [25, 'Z', ''],
      ]) {
        commands.push({ sheet: 'CSV', cell: cell + (row + 1), value });
        expected[row][column] = '"' + value + '"';
      }
    }
    for (const encoding of ['utf-8', 'utf-16le', 'utf-16be']) {
      writeFileSync(source, encodeCsvText(text, encoding, true));
      const index = readCsvFileIndex(source),
        edited = applyCsvCommands(index, commands).index,
        patches = csvTextEdits(edited);
      assert.equal(
        patches.length,
        rows.length * 2,
        'Adjacent changed fields should share a patch; distant changes should leave source spans intact'
      );
      assert(patches.every((p) => !p.text.includes('column0')));
      assert(patches.reduce((sum, p) => sum + p.text.length, 0) < text.length / 10);
      const digest = writeCsvIndex(edited, source, target),
        bytes = readFileSync(target);
      assert.deepEqual(
        bytes,
        Buffer.from(encodeCsvText(expected.map((r) => r.join(',')).join('\r\n') + '\r\n', encoding, true))
      );
      assert.equal(digest, createHash('sha256').update(bytes).digest('hex'));
      assert.deepEqual(readFileSync(source), Buffer.from(encodeCsvText(text, encoding, true)));
      rmSync(target);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('decoded row cache has a cumulative byte budget for wide records and versions share immutable source rows', () => {
  const text = Array.from({ length: 40 }, (_, row) => `row${row},"${'x'.repeat(32768)}"`).join('\n'),
    index = indexCsv(text);
  for (let row = 0; row < 40; row++) assert.equal(csvReadRange(index, `A${row + 1}`).items[0].value, `row${row}`);
  const cache = csvCacheStatistics(index);
  assert(cache.rows > 0);
  assert(cache.rows < 40);
  assert(cache.estimatedBytes <= cache.maxEstimatedBytes);
  assert.equal(cache.maxEstimatedBytes, 2 * 1024 * 1024);
  const next = applyCsvCommands(index, [{ sheet: 'CSV', cell: 'A1', value: 'changed' }]).index;
  assert.equal(csvReadRange(next, 'A1').items[0].value, 'changed');
  assert.equal(csvReadRange(index, 'A1').items[0].value, 'row0');
  assert.deepEqual(csvCacheStatistics(next), csvCacheStatistics(index));
  const oversized = indexCsv('"' + 'y'.repeat(600000) + '"');
  assert.equal(csvReadRange(oversized, 'A1').items[0].value.length, 600000);
  assert.equal(csvCacheStatistics(oversized).rows, 0);
});

test('uncached file windows reject changed or replaced private snapshot bytes', () => {
  const dir = mkdtempSync(join(tmpdir(), 'workbook-csv-private-source-'));
  try {
    const source = join(dir, 'source.csv'),
      text = 'head,value\n' + 'row,value\n'.repeat(250000);
    writeFileSync(source, text);
    const first = readCsvFileIndex(source);
    assert.equal(csvReadRange(first, 'A1').items[0].value, 'head');
    writeFileSync(source, text.replace(/value/g, 'other'));
    assert.throws(() => csvReadRange(first, 'B250001'), /source changed/);
    const second = readCsvFileIndex(source);
    rmSync(source);
    writeFileSync(source, text.replace(/value/g, 'other'));
    assert.throws(() => csvReadRange(second, 'B250001'), /source changed/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
