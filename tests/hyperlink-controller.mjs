import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHyperlinkController } from '../ui/hyperlinkController.ts';
const receipt = { documentId: 'doc-A', ownerId: 'owner-A', state: 'head-1', revision: 'revision-1' },
  raw = {
    schema: 'workbook.xlsx.hyperlink-definition.v1',
    sheet: 'Main',
    cell: 'A1',
    hyperlinkRevision: 'a'.repeat(64),
    link: { kind: 'url', target: 'https://example.test/' },
    tooltip: 'tip',
    canEdit: true,
    protectedReason: null,
  };
const frame = (patch) => ({
  key: 'A/1/epoch1',
  family: 'A/cell/epoch1',
  sheet: 'Main',
  cell: 'A1',
  receipt,
  active: true,
  writable: true,
  read: async () => ({ ...raw, ...receipt }),
  apply: async () => {},
  ...patch,
});
const tick = () => new Promise((r) => setTimeout(r, 0));
async function ready(c, f) {
  c.setFrame(f);
  await tick();
  assert(c.getSnapshot().definition);
}
test('one read lane joins retired inspection and accepts only the latest bound target', async () => {
  const c = createHyperlinkController();
  let release,
    active = 0,
    peak = 0;
  const old = frame({
    read: async () => {
      peak = Math.max(peak, ++active);
      await new Promise((r) => (release = r));
      active--;
      return { ...raw, ...receipt };
    },
  });
  c.setFrame(old);
  c.setFrame(
    frame({
      key: 'A/2/epoch2',
      receipt: { ...receipt, state: 'head-2' },
      read: async () => {
        peak = Math.max(peak, ++active);
        active--;
        return { ...raw, ...receipt, state: 'head-2' };
      },
    })
  );
  assert.equal(c.getSnapshot().definition, null);
  release();
  await tick();
  await tick();
  assert.equal(peak, 1);
  assert.equal(c.getSnapshot().definition.state, 'head-2');
  c.dispose();
});
test('one captured write keeps original command/accepted truth after normal head advance', async () => {
  const c = createHyperlinkController(),
    writes = [];
  let release;
  const original = frame({
    apply: async (command) => {
      writes.push(command);
      await new Promise((r) => (release = r));
    },
  });
  await ready(c, original);
  const promise = c.write(original, { kind: 'internal', sheet: 'Dest', cell: 'B2' }, 'Visit');
  assert.equal(c.getSnapshot().pending, true);
  await assert.rejects(c.write(original, null), /Wait/);
  c.setFrame(
    frame({
      key: 'A/2/epoch1',
      receipt: { ...receipt, state: 'head-2' },
      read: async () => ({ ...raw, ...receipt, state: 'head-2' }),
    })
  );
  release();
  const result = await promise;
  assert.equal(result.accepted, true);
  assert.equal(result.current, false);
  assert.deepEqual(writes, [
    [
      {
        type: 'setHyperlink',
        sheet: 'Main',
        cell: 'A1',
        expectedHyperlinkRevision: 'a'.repeat(64),
        link: { kind: 'internal', sheet: 'Dest', cell: 'B2' },
        tooltip: 'Visit',
      },
    ],
  ]);
  c.dispose();
});
test('retired failures, closed/reopened pane epochs and other owners cannot receive old errors or commands', async () => {
  const c = createHyperlinkController();
  let reject;
  const old = frame({ apply: () => new Promise((_, r) => (reject = r)) });
  await ready(c, old);
  const pending = c.write(old, null);
  c.setFrame({ ...old, active: false });
  c.setFrame(frame({ key: 'A/1/epoch2', family: 'A/cell/epoch2' }));
  reject(Error('old private failure'));
  await assert.rejects(pending);
  await tick();
  assert.equal(c.getSnapshot().error, '');
  await assert.rejects(c.write(old, null), /changed/);
  c.dispose();
});
test('latest readonly/busy state and imported protection refuse before the writer', async () => {
  const c = createHyperlinkController(),
    writes = [];
  const original = frame({ apply: async (command) => writes.push(command) });
  await ready(c, original);
  c.setFrame({ ...original, writable: false, reason: 'Read-only' });
  await assert.rejects(c.write(original, null), /Read-only/);
  await ready(
    c,
    frame({
      key: 'protected',
      read: async () => ({ ...raw, ...receipt, link: null, canEdit: false, protectedReason: 'Imported range' }),
    })
  );
  await assert.rejects(c.write(frame({ key: 'protected' }), null), /Imported range/);
  assert.deepEqual(writes, []);
  c.dispose();
});
test('follow is read-only and late internal results cannot navigate another active target', async () => {
  const c = createHyperlinkController(),
    targets = [];
  let release;
  const original = frame({
    follow: () => new Promise((r) => (release = r)),
    navigate: (...args) => targets.push(args),
  });
  await ready(c, original);
  const pending = c.follow(original);
  c.setFrame(frame({ key: 'other', family: 'other' }));
  release({ kind: 'internal', sheet: 'Dest', cell: 'B2' });
  assert.equal((await pending).current, false);
  assert.deepEqual(targets, []);
  await ready(
    c,
    frame({
      key: 'current',
      follow: async () => ({ kind: 'internal', sheet: 'Dest', cell: 'B2' }),
      navigate: (...args) => targets.push(args),
    })
  );
  await c.follow(
    frame({
      key: 'current',
      follow: async () => ({ kind: 'internal', sheet: 'Dest', cell: 'B2' }),
      navigate: (...args) => targets.push(args),
    })
  );
  assert.deepEqual(targets, [['Dest', 'B2']]);
  c.dispose();
});
test('mismatched native DTOs and readonly inactive frames cannot populate or mutate the form', async () => {
  const c = createHyperlinkController(),
    original = frame({ read: async () => ({ ...raw, ...receipt, cell: 'B1' }) });
  c.setFrame(original);
  await tick();
  assert.equal(c.getSnapshot().definition, null);
  assert.match(c.getSnapshot().error, /captured target/);
  c.setFrame({ ...original, active: false });
  assert.equal(c.getSnapshot().definition, null);
  await assert.rejects(c.write(original, null), /changed/);
  c.dispose();
});
test('new links in blank cells accept metadata and visible address in one batch; formulas keep their data', async () => {
  const c = createHyperlinkController(),
    writes = [];
  const blank = frame({
    read: async () => ({ ...raw, ...receipt, hyperlinkRevision: null, link: null }),
    readCell: async () => null,
    apply: async (commands) => writes.push(commands),
  });
  await ready(c, blank);
  await c.write(blank, { kind: 'url', target: 'https://example.test/new' }, '', '');
  assert.deepEqual(writes[0], [
    {
      type: 'setHyperlink',
      sheet: 'Main',
      cell: 'A1',
      expectedHyperlinkRevision: null,
      link: { kind: 'url', target: 'https://example.test/new' },
      tooltip: '',
    },
    { sheet: 'Main', cell: 'A1', value: 'https://example.test/new', expectedValue: null, expectedFormula: null },
  ]);
  const formula = frame({
    key: 'formula',
    readCell: async () => ({ value: 3, formula: 'B1+1', dataType: 'number' }),
    apply: async (commands) => writes.push(commands),
  });
  await ready(c, formula);
  await assert.rejects(
    c.write(formula, { kind: 'url', target: 'https://example.test/' }, '', 'Overwrite'),
    /worksheet/
  );
  assert.equal(writes.length, 1);
  c.dispose();
});
test('display labels stay literal and captured; same text adds no separate cell change', async () => {
  const c = createHyperlinkController(),
    writes = [],
    cell = { value: 'Original', dataType: 'text' };
  const original = frame({ readCell: async () => cell, apply: async (commands) => writes.push(commands) });
  await ready(c, original);
  cell.value = 'External change';
  await c.write(original, { kind: 'url', target: 'https://example.test/' }, '', '=SUM(A1)');
  assert.equal(writes[0][1].value, '=SUM(A1)');
  assert.equal(writes[0][1].expectedValue, 'Original');
  await c.write(original, { kind: 'url', target: 'https://example.test/' }, '', 'Original');
  assert.equal(writes[1].length, 1);
  c.dispose();
});
