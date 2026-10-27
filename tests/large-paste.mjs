import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID, createHash } from 'node:crypto';
import { inspectPaste, pastePages } from '../ui/largePaste.ts';
import { parseDelimited } from '../core/delimited.ts';
import { parseCellInput } from '../core/cell-input.mjs';
import { pasteBytes, PASTE_PAGE_BYTES, validatePasteDescriptor } from '../core/paste-protocol.mjs';
import { createPasteStager } from '../adapters/large-paste.mjs';

const options = { sheet: 'Sheet1', start: 'B3' };
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function staging(t, overrides = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-paste-staging-')),
    documentId = randomUUID();
  const stager = createPasteStager({ directory, documentId, ...overrides });
  t.after(async () => {
    await stager.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  return { directory, documentId, stager };
}
async function uploadText(stager, text, state = 'base') {
  const inspection = inspectPaste(text, options),
    { id } = await stager.begin({ ...inspection, state });
  for (const page of pastePages(text, inspection)) await stager.upload({ id, state, ...page });
  return stager.seal({ id, state });
}

test('streaming TSV matches the established parser for quotes, CRLF, whitespace, trailing fields and ragged holes', () => {
  for (const text of [
    '',
    '\n',
    '\r\n',
    '\t',
    'a\t\n',
    'a\tb\nshort\n',
    '""',
    '"x\ty"\t"line\r\nnext"\r\n"a""b"\t  001  ',
    'x"y\tz',
    '\n\n',
  ]) {
    const expected = parseDelimited(text, '\t').rows,
      inspection = inspectPaste(text, options),
      actual = [];
    for (const page of pastePages(text, inspection))
      for (const [offset, value] of page.cells)
        actual.push([Math.floor(offset / inspection.shape.columns), offset % inspection.shape.columns, value]);
    assert.deepEqual(
      actual,
      expected.flatMap((row, r) => row.map((value, c) => [r, c, value])),
      text
    );
    assert.equal(
      inspection.count,
      expected.reduce((n, row) => n + row.length, 0)
    );
  }
  for (const text of ['"unterminated', '"closed" bad', '"closed""'])
    assert.throws(() => inspectPaste(text, options), /Malformed/);
});

test('input policy preserves identifiers, precision, signed zero, formulas, booleans and literal target text', () => {
  for (const text of ['001', '1234567890123456', '1.0000000000000002', '9007199254740993', '1e999', ' 12 ', '12.'])
    assert.deepEqual(parseCellInput(text), { value: text });
  assert.deepEqual(parseCellInput('12.5'), { value: 12.5 });
  assert(Object.is(parseCellInput('-0').value, -0));
  assert.deepEqual(parseCellInput('=A1+1'), { formula: 'A1+1' });
  assert.deepEqual(parseCellInput(''), { value: null });
  assert.deepEqual(parseCellInput('false'), { value: false });
  assert.deepEqual(parseCellInput("'001"), { value: '001' });
  for (const text of ['=A1', 'FALSE', "'raw", '']) assert.deepEqual(parseCellInput(text, true), { value: text });
  for (const text of ['ASCII', 'é', '😀', '\ud800', '\udc00', '\"\n\t'])
    assert.equal(pasteBytes(text), Buffer.byteLength(text));
});

test('deterministic mixed-delimiter corpus agrees with established TSV values and malformed-quote refusals', () => {
  const alphabet = ['a', ' ', '"', '\t', '\r', '\n', 'é', '😀'];
  let seed = 0x6b33;
  for (let sample = 0; sample < 1000; sample++) {
    let text = '';
    for (let i = 0; i < sample % 23; i++) {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      text += alphabet[seed % alphabet.length];
    }
    let expected, error;
    try {
      expected = parseDelimited(text, '\t').rows;
    } catch (cause) {
      error = cause;
    }
    if (error) {
      assert.throws(() => inspectPaste(text, options), /Malformed/);
      continue;
    }
    const plan = inspectPaste(text, options),
      actual = [];
    for (const page of pastePages(text, plan))
      for (const [offset, value] of page.cells)
        actual.push([Math.floor(offset / plan.shape.columns), offset % plan.shape.columns, value]);
    assert.deepEqual(
      actual,
      expected.flatMap((row, r) => row.map((value, c) => [r, c, value])),
      JSON.stringify(text)
    );
  }
});

test('100,000 input cells and long fields stream through count and byte bounded pages without row padding', () => {
  const text = ('1\t'.repeat(99) + '1\n').repeat(1000),
    inspection = inspectPaste(text, { sheet: 'Sheet1', start: 'A1' });
  assert.deepEqual(inspection.shape, { sheet: 'Sheet1', start: 'A1', rows: 1000, columns: 100 });
  assert.equal(inspection.count, 100000);
  let count = 0,
    pages = 0;
  for (const page of pastePages(text, inspection)) {
    assert.equal(page.sequence, pages++);
    assert(page.cells.length <= 1000);
    for (const [offset, value] of page.cells) {
      assert.equal(offset, count++);
      assert.equal(value, '1');
    }
  }
  assert.equal(pages, 100);
  assert.equal(count, 100000);
  const long = ('"' + '\t'.repeat(32000) + '"\t').repeat(20) + 'last',
    longPlan = inspectPaste(long, { sheet: 'Sheet1', start: 'A1' });
  let longPages = 0;
  for (const page of pastePages(long, longPlan)) {
    longPages++;
    assert(page.cells.reduce((n, entry) => n + Buffer.byteLength(JSON.stringify(entry) + '\n'), 0) <= PASTE_PAGE_BYTES);
  }
  assert(longPages > 1);
  assert.throws(() => inspectPaste('a\n'.repeat(1001), { sheet: 'Sheet1', start: 'A1048000' }), /limits/);
  assert.throws(() => inspectPaste('a\t'.repeat(1000) + '\n' + 'b\n'.repeat(1000), options), /1,000,000/);
  assert.throws(() => inspectPaste('\ud800', options), /XML/);
  assert.throws(() => inspectPaste('\u0000', options), /XML/);
});

test('immutable upload replays its exact parent identity after owner restart and prunes only unreferenced attachments', async (t) => {
  const { directory, documentId, stager } = await staging(t),
    descriptor = await uploadText(stager, '001\tTRUE\nshort\t"line\nnext"');
  assert.equal(descriptor.type, 'importCells');
  assert.equal(
    Object.keys(descriptor).some((key) => /path|file/i.test(key)),
    false
  );
  const filename = path.join(directory, descriptor.id + '.paste.jsonl'),
    bytes = await fs.readFile(filename);
  assert.equal(descriptor.sha256, digest(bytes));
  assert.equal(descriptor.bytes, bytes.length);
  const replay = createPasteStager({ directory, documentId });
  t.after(() => replay.close());
  let seen = [];
  for await (const page of replay.pages(descriptor, { state: 'base' })) seen.push(...page.cells);
  assert.deepEqual(seen, [
    [0, '001'],
    [1, 'TRUE'],
    [2, 'short'],
    [3, 'line\nnext'],
  ]);
  await assert.rejects(
    async () => {
      for await (const _ of replay.pages(descriptor, { state: 'other' })) {
      }
    },
    { code: /^(internal-invalid-field|internal-mismatch)$/ }
  );
  const other = createPasteStager({ directory, documentId: randomUUID() });
  t.after(() => other.close());
  await assert.rejects(async () => {
    for await (const _ of other.pages(descriptor, { state: 'base' })) {
    }
  }, /document/);
  assert.deepEqual(await stager.prune([descriptor.id]), { removed: 0 });
  assert.deepEqual(await stager.prune([]), { removed: 1 });
  await assert.rejects(stager.cancel({ id: descriptor.id, state: 'base' }), { code: 'paste-stale' });
});

test('partial uploads reject sequence, duplicate offsets, shape overflow and quota without changing accepted files', async (t) => {
  const { directory, stager } = await staging(t, { maxBytes: 1500 }),
    inspection = inspectPaste('a\tb', options),
    upload = await stager.begin({ ...inspection, state: 'base' });
  await assert.rejects(stager.upload({ ...upload, state: 'base', sequence: 1, cells: [[0, 'a']] }), {
    code: /^(csv-refused|package-bad-input|paste-bad-input)$/,
  });
  await assert.rejects(stager.upload({ ...upload, state: 'other', sequence: 0, cells: [[0, 'a']] }), {
    code: 'paste-stale',
  });
  await stager.upload({ ...upload, state: 'base', sequence: 0, cells: [[0, 'a']] });
  await assert.rejects(stager.upload({ ...upload, state: 'base', sequence: 1, cells: [[0, 'b']] }), {
    code: 'paste-bad-input',
  });
  await assert.rejects(stager.upload({ ...upload, state: 'base', sequence: 1, cells: [[2, 'b']] }), {
    code: 'paste-bad-input',
  });
  await assert.rejects(
    stager.upload({ ...upload, state: 'base', sequence: 1, cells: [[1, 'x'.repeat(1400)]] }),
    /byte budget/
  );
  await assert.rejects(stager.seal({ ...upload, state: 'base' }), /incomplete/);
  await stager.cancel({ ...upload, state: 'base' });
  assert.equal(stager.stats().attachmentBytes, 0);
  assert.deepEqual(await fs.readdir(directory), []);
  const next = await uploadText(stager, 'accepted');
  assert(next.sha256);
  assert.equal((await fs.readdir(directory)).length, 1);
});

test('cancellation can clean a partial upload after the operation token is cancelled', async (t) => {
  let cancelled = false;
  const { directory, stager } = await staging(t, {
    checkCancelled: () => {
      if (cancelled) throw Error('cancelled');
    },
  });
  const upload = await stager.begin({ ...inspectPaste('a', options), state: 'base' });
  cancelled = true;
  await assert.rejects(stager.upload({ ...upload, state: 'base', sequence: 0, cells: [[0, 'a']] }), /cancelled/);
  await stager.cancel({ ...upload, state: 'base' });
  assert.deepEqual(await fs.readdir(directory), []);
  assert.equal(stager.stats().activeUploads, 0);
});

test('attachment tampering, symlinks and forged path descriptors fail before yielding cells', async (t) => {
  const { directory, stager } = await staging(t),
    descriptor = await uploadText(stager, 'kept'),
    filename = path.join(directory, descriptor.id + '.paste.jsonl');
  assert.throws(() => validatePasteDescriptor({ ...descriptor, path: '/tmp/other' }), /descriptor/);
  await fs.chmod(filename, 0o600);
  const original = await fs.readFile(filename);
  await fs.writeFile(filename, Buffer.from(original.toString().replace('kept', 'evil')));
  let yielded = 0;
  await assert.rejects(
    async () => {
      for await (const page of stager.pages(descriptor, { state: 'base' })) yielded += page.cells.length;
    },
    { code: 'paste-malformed' }
  );
  assert.equal(yielded, 0);
  const outside = path.join(directory, 'outside');
  await fs.writeFile(outside, original);
  await fs.rm(filename);
  await fs.symlink(outside, filename);
  await assert.rejects(
    async () => {
      for await (const _ of stager.pages(descriptor, { state: 'base' })) {
      }
    },
    { code: 'internal-missing-field' }
  );
});

test('recovery initialization removes incomplete private uploads but preserves sealed history data', async (t) => {
  const { directory, documentId, stager } = await staging(t),
    kept = await uploadText(stager, 'history');
  const partial = randomUUID() + '.paste.partial';
  await fs.writeFile(path.join(directory, partial), 'interrupted');
  const recovery = createPasteStager({ directory, documentId });
  t.after(() => recovery.close());
  const begin = await recovery.begin({ ...inspectPaste('new', options), state: 'base' });
  await recovery.cancel({ ...begin, state: 'base' });
  assert.deepEqual(await fs.readdir(directory), [kept.id + '.paste.jsonl']);
});

test('declared shape inflation and mutation during a suspended replay refuse final publication', async (t) => {
  const { directory, stager } = await staging(t);
  const forged = await stager.begin({
    shape: { sheet: 'Sheet1', start: 'A1', rows: 2, columns: 2 },
    count: 1,
    policy: 'grid',
    state: 'base',
  });
  await stager.upload({ ...forged, state: 'base', sequence: 0, cells: [[0, 'one']] });
  await assert.rejects(stager.seal({ ...forged, state: 'base' }), { code: 'paste-bad-input' });
  await stager.cancel({ ...forged, state: 'base' });
  const descriptor = await uploadText(stager, 'value\n'.repeat(1001)),
    filename = path.join(directory, descriptor.id + '.paste.jsonl');
  const iterator = stager.pages(descriptor, { state: 'base' })[Symbol.asyncIterator](),
    first = await iterator.next();
  assert.equal(first.value.cells.length, 1000);
  await fs.chmod(filename, 0o600);
  const bytes = await fs.readFile(filename);
  await fs.writeFile(filename, bytes);
  await assert.rejects(iterator.next(), { code: 'paste-stale' });
});
