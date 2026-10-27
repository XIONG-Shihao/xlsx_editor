import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID, createHash } from 'node:crypto';
import { indexCsv, csvCell } from '../core/csv-workbook.mjs';
import { parseCellInput } from '../core/cell-input.mjs';
import {
  createBlobDelimitedSource,
  delimitedFields,
  inspectDelimitedSource,
  delimitedImportPages,
  stageDelimitedImport,
} from '../core/delimited-import.mjs';
import { PASTE_PAGE_BYTES, pasteBytes } from '../core/paste-protocol.mjs';

const options = { sheet: 'Sheet1', start: 'B2', delimiter: ',' };
function source(bytes, { chunk = 7, name = 'chosen.csv', change } = {}) {
  bytes = Buffer.from(bytes);
  let reads = 0;
  return {
    id: randomUUID(),
    name,
    size: bytes.length,
    async *chunks() {
      const selected = change?.(++reads) ?? bytes;
      for (let at = 0; at < selected.length; at += chunk) yield selected.subarray(at, at + chunk);
    },
  };
}
async function fields(input, delimiter = ',') {
  const out = [];
  for await (const field of delimitedFields(input, delimiter)) out.push(field);
  return out;
}
async function pages(input, inspection) {
  const out = [];
  for await (const page of delimitedImportPages(input, inspection)) out.push(...page.cells);
  return out;
}
function sinkFor(inspection, { upload, seal } = {}) {
  const id = randomUUID(),
    calls = [];
  return {
    calls,
    async begin() {
      calls.push('begin');
      return { id, sequence: 0 };
    },
    async upload(input) {
      calls.push('upload');
      return upload ? upload(input) : { id, sequence: input.sequence + 1, count: (input.sequence + 1) * 1000 };
    },
    async seal() {
      calls.push('seal');
      return seal
        ? seal()
        : {
            type: 'importCells',
            id,
            sha256: 'a'.repeat(64),
            bytes: 100,
            count: inspection.count,
            shape: inspection.shape,
            policy: 'text',
          };
    },
    async cancel() {
      calls.push('cancel');
    },
    async prune() {
      calls.push('prune');
    },
  };
}

test('field scanner agrees with existing CSV semantics at every byte boundary and keeps ragged holes', async () => {
  const cases = [
    '',
    'a',
    'a,',
    'a\r\n',
    '\r\n',
    'a\r\nb,c\rd\n',
    '01,=SUM(A1),  spaced  ,,""\nlast',
    '"a,b","two\r\nlines","double "" quote"\nlast,short',
    'bare"quote,x\n"😀",字',
  ];
  for (const text of cases)
    for (const delimiter of [',', ';', '\t']) {
      let expected;
      try {
        expected = indexCsv(text, { delimiter });
      } catch {
        for (const chunk of [1, 2, 3, 4, 7, 64])
          await assert.rejects(
            inspectDelimitedSource(source(Buffer.from(text), { chunk }), { ...options, delimiter }),
            { code: /^(csv-malformed|duplicates-malformed|table-malformed)$/ }
          );
        continue;
      }
      for (const chunk of [1, 2, 3, 4, 7, 64]) {
        const input = source(Buffer.from(text), { chunk }),
          inspection = await inspectDelimitedSource(input, { ...options, delimiter });
        assert.equal(inspection.shape.rows, expected.rows);
        assert.equal(inspection.shape.columns, expected.columns);
        const actual = await pages(input, inspection),
          wanted = [];
        for (let row = 0; row < expected.rows; row++)
          for (let col = 0; col < expected.rowWidths[row]; col++)
            wanted.push([row * expected.columns + col, csvCell(expected, row, col)]);
        assert.deepEqual(actual, wanted, `chunk ${chunk}: ${JSON.stringify(text)}`);
        assert.equal(inspection.source.sha256, createHash('sha256').update(text).digest('hex'));
      }
    }
  const input = source('01,=SUM(A1),,TRUE,9007199254740993\nragged');
  const inspection = await inspectDelimitedSource(input, options),
    actual = await pages(input, inspection);
  assert.deepEqual(actual, [
    [0, '01'],
    [1, '=SUM(A1)'],
    [2, ''],
    [3, 'TRUE'],
    [4, '9007199254740993'],
    [5, 'ragged'],
  ]);
  for (const [, text] of actual) assert.deepEqual(parseCellInput(text, true), { value: text });
});

test('auto separator and UTF-8/UTF-16 BOM decoding retain exact strings without replacement', async () => {
  const text = '"comma,inside";01;😀\r\nshort;""',
    le = Buffer.from(text, 'utf16le'),
    be = Buffer.from(le);
  be.swap16();
  const encodings = [
    ['utf-8', false, Buffer.from(text)],
    ['utf-8', true, Buffer.concat([Buffer.from([239, 187, 191]), Buffer.from(text)])],
    ['utf-16le', true, Buffer.concat([Buffer.from([255, 254]), le])],
    ['utf-16be', true, Buffer.concat([Buffer.from([254, 255]), be])],
  ];
  for (const [encoding, bom, bytes] of encodings) {
    const input = source(bytes, { chunk: 1 }),
      inspection = await inspectDelimitedSource(input, { ...options, delimiter: 'auto' });
    assert.equal(inspection.delimiter, ';');
    assert.equal(inspection.encoding, encoding);
    assert.equal(inspection.bom, bom);
    assert.deepEqual(await pages(input, inspection), [
      [0, 'comma,inside'],
      [1, '01'],
      [2, '😀'],
      [3, 'short'],
      [4, ''],
    ]);
  }
});

test('malformed grammar, encoding, XML scalars, geometry and source windows refuse before staging', async () => {
  for (const [bytes, pattern] of [
    [Buffer.from('"never ends'), /unterminated/],
    [Buffer.from('"closed"later'), /closing quote/],
    [Buffer.from([0xc3, 0x28]), /encoding/],
    [Buffer.from([255, 254, 0, 0]), /UTF-32|encoding/],
    [Buffer.from([255, 254, 0x41]), /encoding/],
    [Buffer.from('a\0b'), /XML|character/],
    [Buffer.from('x'.repeat(32769)), /text limits/],
  ])
    await assert.rejects(inspectDelimitedSource(source(bytes), options), pattern);
  await assert.rejects(inspectDelimitedSource(source('a,b'), { ...options, start: 'XFD1' }), /workbook|position/);
  await assert.rejects(inspectDelimitedSource(source('a'), { ...options, start: 'a1' }), /origin|position/);
  await assert.rejects(inspectDelimitedSource(source('a', { name: '../path.csv' }), options), /immutable/);
  await assert.rejects(inspectDelimitedSource(source('a', { name: 'chosen.exe' }), options), /immutable/);
  const hugeWindow = {
    id: randomUUID(),
    name: 'large.csv',
    size: 262145,
    async *chunks() {
      yield new Uint8Array(262145);
    },
  };
  await assert.rejects(inspectDelimitedSource(hugeWindow, options), { code: 'csv-limit' });
  const incomplete = source('a');
  incomplete.size = 2;
  await assert.rejects(inspectDelimitedSource(incomplete, options), { code: /^(csv-refused|session-unsupported)$/ });
  const stalled = {
    id: randomUUID(),
    name: 'empty.csv',
    size: 0,
    async *chunks() {
      for (let i = 0; i < 9; i++) yield new Uint8Array();
    },
  };
  await assert.rejects(inspectDelimitedSource(stalled, options), {
    code: /^(csv-bad-input|package-bad-input|paste-bad-input)$/,
  });
});

test('one million fields use bounded byte windows and pages without a whole document or cell matrix', async () => {
  const row = Buffer.from('x,'.repeat(499) + 'x\n'),
    rows = 2000,
    expected = createHash('sha256');
  let reads = 0;
  const input = {
    id: randomUUID(),
    name: 'million.csv',
    size: row.length * rows,
    async *chunks() {
      reads++;
      for (let i = 0; i < rows; i++) yield row;
    },
  };
  for (let i = 0; i < rows; i++) expected.update(row);
  const inspection = await inspectDelimitedSource(input, { sheet: 'Sheet1', start: 'A1', delimiter: ',' });
  assert.equal(inspection.count, 1000000);
  assert.deepEqual(inspection.shape, { sheet: 'Sheet1', start: 'A1', rows, columns: 500 });
  assert.equal(inspection.source.sha256, expected.digest('hex'));
  let count = 0,
    pageCount = 0,
    previous = -1;
  for await (const page of delimitedImportPages(input, inspection)) {
    assert.equal(page.sequence, pageCount++);
    assert(page.cells.length <= 1000);
    assert(page.cells.reduce((sum, entry) => sum + pasteBytes(JSON.stringify(entry) + '\n'), 0) <= PASTE_PAGE_BYTES);
    assert.equal(page.cells[0][0], previous + 1);
    previous = page.cells.at(-1)[0];
    count += page.cells.length;
  }
  assert.equal(count, 1000000);
  assert.equal(pageCount, 1000);
  assert.equal(reads, 2);
  const beyond = {
    ...input,
    size: row.length * (rows + 1),
    async *chunks() {
      for (let i = 0; i <= rows; i++) yield row;
    },
  };
  await assert.rejects(inspectDelimitedSource(beyond, { sheet: 'Sheet1', start: 'A1', delimiter: ',' }), /1,000,000/);
});

test('raw hash fences suppress the final page of same-sized changed source and cancellation stops reads', async () => {
  const original = Buffer.from('a\n'.repeat(2001)),
    mutated = Buffer.from(original);
  mutated[mutated.length - 2] = 98;
  const input = source(original, { change: (read) => (read === 2 ? mutated : original) }),
    inspection = await inspectDelimitedSource(input, options);
  let received = 0;
  await assert.rejects(
    async () => {
      for await (const page of delimitedImportPages(input, inspection)) received += page.cells.length;
    },
    { code: 'csv-stale' }
  );
  assert.equal(received, 2000);
  const controller = new AbortController(),
    aborted = source('a\n'.repeat(2001));
  let checks = 0;
  aborted.check = () => {
    if (++checks === 3) controller.abort(Error('stopped source'));
  };
  await assert.rejects(inspectDelimitedSource(aborted, { ...options, signal: controller.signal }), /stopped source/);
});

test('large legal fields honor the byte limit before the cell-count limit', async () => {
  const input = source(('"' + '😀'.repeat(16000) + '"\n').repeat(40), { chunk: 65536 }),
    inspection = await inspectDelimitedSource(input, options);
  let count = 0,
    received = 0;
  for await (const page of delimitedImportPages(input, inspection)) {
    count++;
    assert(page.cells.length < 1000);
    assert(page.cells.reduce((sum, entry) => sum + pasteBytes(JSON.stringify(entry) + '\n'), 0) <= PASTE_PAGE_BYTES);
    received += page.cells.length;
  }
  assert.equal(count, 3);
  assert.equal(received, 40);
});

test('staging awaits each acknowledgement, validates receipts and preserves cleanup on seal uncertainty', async () => {
  const input = source('x\n'.repeat(2001)),
    inspection = await inspectDelimitedSource(input, options);
  let outstanding = 0,
    total = 0;
  const sink = sinkFor(inspection, {
    upload: async (request) => {
      assert.equal(outstanding++, 0);
      await new Promise((resolve) => setImmediate(resolve));
      total += request.cells.length;
      outstanding--;
      return { sequence: request.sequence + 1, count: total };
    },
  });
  const ready = await stageDelimitedImport(input, inspection, { state: 'accepted', sink });
  assert.equal(ready.descriptor.count, 2001);
  assert.deepEqual(sink.calls, ['begin', 'upload', 'upload', 'upload', 'seal']);
  assert.equal(ready.source.sha256, inspection.source.sha256);
  const wrong = sinkFor(inspection, { upload: () => ({ sequence: 99, count: 1 }) });
  await assert.rejects(stageDelimitedImport(input, inspection, { state: 'accepted', sink: wrong }), {
    code: /^(csv-bad-input|query-stale)$/,
  });
  assert.equal(wrong.calls.at(-1), 'cancel');
  const malformed = sinkFor(inspection, {
    upload: (request) => ({ sequence: request.sequence + 1, count: Math.min((request.sequence + 1) * 1000, 2001) }),
    seal: () => ({ unexpected: true }),
  });
  await assert.rejects(stageDelimitedImport(input, inspection, { state: 'accepted', sink: malformed }), /descriptor/);
  assert.equal(malformed.calls.at(-1), 'prune');
  const uncertain = sinkFor(inspection, {
    upload: (request) => ({ sequence: request.sequence + 1, count: Math.min((request.sequence + 1) * 1000, 2001) }),
    seal: () => {
      throw Error('seal outcome unknown');
    },
  });
  uncertain.cancel = async () => {
    uncertain.calls.push('cancel');
    throw Error('already sealed');
  };
  await assert.rejects(
    stageDelimitedImport(input, inspection, { state: 'accepted', sink: uncertain }),
    /outcome unknown/
  );
  assert.deepEqual(uncertain.calls.slice(-2), ['cancel', 'prune']);
});

test('stale owner and cancellation after seal invoke trusted retained-history pruning', async () => {
  const input = source('one'),
    inspection = await inspectDelimitedSource(input, options),
    controller = new AbortController();
  const sink = sinkFor(inspection, { upload: () => ({ sequence: 1, count: 1 }) });
  const seal = sink.seal;
  sink.seal = async () => {
    const descriptor = await seal();
    controller.abort(Error('stopped after seal'));
    return descriptor;
  };
  await assert.rejects(
    stageDelimitedImport(input, inspection, { state: 'accepted', signal: controller.signal, sink }),
    /stopped after seal/
  );
  assert.equal(sink.calls.at(-1), 'prune');
  const stale = sinkFor(inspection, { upload: () => ({ sequence: 1, count: 1 }) });
  let valid = true;
  const upload = stale.upload;
  stale.upload = async (request) => {
    const receipt = await upload(request);
    valid = false;
    return receipt;
  };
  await assert.rejects(
    stageDelimitedImport(input, inspection, {
      state: 'accepted',
      sink: stale,
      checkCurrent: () => {
        if (!valid) throw Error('initiating tab changed');
      },
    }),
    /tab changed/
  );
  assert.equal(stale.calls.at(-1), 'cancel');
});

test('browser File/Blob reads bounded slices and exposes only opaque metadata', async () => {
  const blob = new Blob(['01\t=SUM(A1)\n😀\t']),
    whole = blob.arrayBuffer;
  blob.arrayBuffer = () => assert.fail('No whole-file read');
  const input = createBlobDelimitedSource(blob, { name: 'selected.tsv', chunkBytes: 1 }),
    inspection = await inspectDelimitedSource(input, { ...options, delimiter: '\t' });
  assert.deepEqual(await pages(input, inspection), [
    [0, '01'],
    [1, '=SUM(A1)'],
    [2, '😀'],
    [3, ''],
  ]);
  assert.equal(Object.hasOwn(input, 'path'), false);
  assert.equal(Object.hasOwn(inspection.source, 'path'), false);
  await input.dispose();
  await assert.rejects(inspectDelimitedSource(input, options), /closed/);
  blob.arrayBuffer = whole;
});
