import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { randomUUID, createHash } from 'node:crypto';
import { createPasteStager } from '../adapters/large-paste.mjs';
import { findImportReceipt } from '../adapters/import-receipt.mjs';
import {
  validatePastePlan,
  validatePasteDescriptor,
  pasteHeader,
  pasteDescriptorIdentity,
} from '../core/paste-protocol.mjs';

const shape = { sheet: 'Sheet1', start: 'B2', rows: 3, columns: 4 };
const conversion = {
  schema: 1,
  policies: ['text', 'number', 'boolean', 'iso-date'],
  date1904: false,
  keepHeader: true,
  outputRows: 2,
  outputColumns: 4,
};
const typed = () => ({
  shape: { ...shape },
  count: 12,
  policy: 'typed-text-v1',
  conversion: structuredClone(conversion),
});
const descriptor = () => ({
  type: 'importCells',
  id: randomUUID(),
  sha256: 'a'.repeat(64),
  bytes: 500,
  schema: 'workbook.xlsx.staged-paste.v2',
  ...typed(),
});
test('closed typed metadata binds dense shape, output bounds, all policies and explicit calendar', () => {
  const plan = typed(),
    accepted = validatePastePlan(plan);
  plan.conversion.policies[0] = 'number';
  assert.equal(accepted.conversion.policies[0], 'text');
  for (const mutate of [
    (p) => p.count--,
    (p) => delete p.conversion,
    (p) => (p.conversion.extra = 1),
    (p) => (p.conversion.schema = 2),
    (p) => (p.conversion.policies = ['number']),
    (p) => (p.conversion.policies[1] = 'guess'),
    (p) => (p.conversion.date1904 = undefined),
    (p) => (p.conversion.keepHeader = 0),
    (p) => (p.conversion.outputRows = 4),
    (p) => (p.conversion.outputColumns = 0),
    (p) => (p.conversion.policies = Array(257).fill('text')),
  ]) {
    const bad = typed();
    mutate(bad);
    assert.throws(() => validatePastePlan(bad), /Invalid/);
  }
  const maximum = {
    shape: { sheet: 'Sheet1', start: 'A1', rows: 3906, columns: 256 },
    count: 999936,
    policy: 'typed-text-v1',
    conversion: { ...conversion, policies: Array(256).fill('iso-date'), outputRows: 3906, outputColumns: 256 },
  };
  assert.equal(validatePastePlan(maximum).count, 999936);
  assert(Buffer.byteLength(JSON.stringify(pasteHeader(maximum, randomUUID(), 'base')) + '\n') < 4096);
  const empty = typed();
  empty.conversion.outputRows = 0;
  assert.equal(validatePastePlan(empty).conversion.outputRows, 0);
});
test('untyped descriptors keep field order, new headers use neutral tags and typed identity includes conversion', () => {
  const legacy = {
    type: 'importCells',
    id: randomUUID(),
    sha256: 'a'.repeat(64),
    bytes: 99,
    shape: { sheet: 'Sheet1', start: 'A1', rows: 1, columns: 2 },
    count: 2,
    policy: 'text',
  };
  const expected = {
    schema: 'workbook.xlsx.staged-paste.v1',
    documentId: 'document',
    state: 'base',
    shape: legacy.shape,
    count: 2,
    policy: 'text',
  };
  assert.equal(
    JSON.stringify(pasteHeader({ shape: legacy.shape, count: 2, policy: 'text' }, 'document', 'base')),
    JSON.stringify(expected)
  );
  assert.equal(
    pasteDescriptorIdentity(legacy),
    JSON.stringify([legacy.type, legacy.id, legacy.sha256, legacy.count, legacy.bytes, 'Sheet1', 'A1', 1, 2, 'text'])
  );
  assert.equal(Object.hasOwn(validatePasteDescriptor(legacy), 'schema'), false);
  assert.throws(() => validatePasteDescriptor({ ...legacy, conversion }), /descriptor/);
  const d = descriptor(),
    nodes = [{ head: 'accepted', parent: null, commands: [d] }],
    request = { operation: 'importReceipt', descriptor: d, parentState: 'base' },
    context = { baseState: 'base', currentState: 'accepted', tipState: 'accepted' };
  assert.equal(findImportReceipt(nodes, request, context).found, true);
  for (const patch of [
    { date1904: true },
    { keepHeader: false },
    { policies: ['text', 'text', 'boolean', 'iso-date'] },
    { outputRows: 1 },
  ])
    assert.throws(
      () =>
        findImportReceipt(
          nodes,
          { ...request, descriptor: { ...d, conversion: { ...d.conversion, ...patch } } },
          context
        ),
      /identity/
    );
  assert.throws(() => validatePasteDescriptor({ ...d, schema: 'workbook.xlsx.staged-paste.v1' }), /descriptor/);
});
test('typed sealed attachments retain raw strings and reject changed descriptor before yielding cells', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'typed-paste-')),
    documentId = randomUUID(),
    stager = createPasteStager({ directory, documentId });
  t.after(async () => {
    await stager.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  const plan = typed(),
    { id } = await stager.begin({ ...plan, state: 'base' }),
    raw = ['Name', 'Amount', 'Active', 'Date', '=SUM(A1)', '-0', 'FALSE', '2011-07-05', '', '', '', ''];
  await stager.upload({ id, state: 'base', sequence: 0, cells: raw.map((value, offset) => [offset, value]) });
  const d = await stager.seal({ id, state: 'base' }),
    bytes = await fs.readFile(path.join(directory, id + '.paste.jsonl'));
  assert.equal(d.sha256, createHash('sha256').update(bytes).digest('hex'));
  assert.deepEqual(JSON.parse(bytes.toString().split('\n')[0]), pasteHeader(plan, documentId, 'base'));
  const replay = createPasteStager({ directory, documentId });
  t.after(() => replay.close());
  const seen = [];
  for await (const page of replay.pages(d, { state: 'base' })) seen.push(...page.cells);
  assert.deepEqual(
    seen,
    raw.map((value, offset) => [offset, value])
  );
  for (const patch of [
    { date1904: true },
    { keepHeader: false },
    { outputRows: 1 },
    { policies: ['text', 'text', 'boolean', 'iso-date'] },
  ]) {
    let yielded = 0;
    await assert.rejects(
      async () => {
        for await (const page of replay.pages({ ...d, conversion: { ...d.conversion, ...patch } }, { state: 'base' }))
          yielded += page.cells.length;
      },
      { code: 'internal-mismatch' }
    );
    assert.equal(yielded, 0);
  }
  assert.deepEqual(await fs.readFile(path.join(directory, id + '.paste.jsonl')), bytes);
});
test('typed uploads preserve existing page and attachment cancellation limits', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'typed-paste-limits-')),
    stager = createPasteStager({ directory, documentId: randomUUID(), maxBytes: 4096 });
  t.after(async () => {
    await stager.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  const plan = {
      shape: { sheet: 'Sheet1', start: 'A1', rows: 1001, columns: 1 },
      count: 1001,
      policy: 'typed-text-v1',
      conversion: {
        schema: 1,
        policies: ['number'],
        date1904: false,
        keepHeader: false,
        outputRows: 1001,
        outputColumns: 1,
      },
    },
    { id } = await stager.begin({ ...plan, state: 'base' });
  await assert.rejects(
    stager.upload({ id, state: 'base', sequence: 0, cells: Array.from({ length: 1001 }, (_, i) => [i, '1']) }),
    /1,000/
  );
  await assert.rejects(stager.upload({ id, state: 'base', sequence: 0, cells: [[0, 'x'.repeat(4096)]] }), /budget/);
  await stager.cancel({ id, state: 'base' });
  assert.deepEqual(await fs.readdir(directory), []);
  assert.equal(stager.stats().attachmentBytes, 0);
});
