import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createPageStylesController } from '../ui/pageStylesController.ts';
const revision = 'a'.repeat(64),
  frame = (patch) => ({
    key: 'owner/head/sheet/pane',
    family: 'owner/sheet/pane',
    sheet: 'Main',
    active: true,
    writable: true,
    theme: { revision: null, sample: null, canApply: true, protectedReason: null },
    background: undefined,
    apply: async () => {},
    ...patch,
  });
test('preview and Reset are local, including read-only sample browsing', () => {
  const c = createPageStylesController(),
    writes = [],
    f = frame({ writable: false, apply: async (v) => writes.push(v) });
  c.setFrame(f);
  c.choose(f, 'themes', 'forest');
  c.choose(f, 'backgrounds', 'soft-dots');
  assert.equal(c.getSnapshot().theme, 'forest');
  assert.equal(c.getSnapshot().background, 'soft-dots');
  c.reset(f);
  assert.equal(c.getSnapshot().theme, null);
  assert.equal(c.getSnapshot().background, null);
  assert.deepEqual(writes, []);
  c.dispose();
});
test('theme and background commands carry the accepted sample, original sheet and exact revision', async () => {
  const c = createPageStylesController(),
    writes = [],
    f = frame({
      theme: { revision, sample: 'neutral', canApply: true, protectedReason: null },
      background: {
        schema: 'workbook.xlsx.sheet-background.v1',
        sheet: 'xl/worksheets/sheet1.xml',
        revision,
        sample: 'graph-paper',
        canEdit: true,
        protectedReason: null,
      },
      apply: async (commands) => writes.push(commands),
    });
  c.setFrame(f);
  c.choose(f, 'themes', 'plum');
  await c.theme(f);
  c.choose(f, 'backgrounds', 'soft-dots');
  await c.background(f);
  await c.background(f, true);
  assert.deepEqual(writes, [
    [{ type: 'applyPageTheme', sample: 'plum', expectedThemeRevision: revision }],
    [{ type: 'setSheetBackground', sheet: 'Main', expectedBackgroundRevision: revision, sample: 'soft-dots' }],
    [{ type: 'removeSheetBackground', sheet: 'Main', expectedBackgroundRevision: revision }],
  ]);
  c.dispose();
});
test('one physical write lane survives owner/head changes without retargeting or late notices', async () => {
  const c = createPageStylesController(),
    writes = [];
  let release;
  const f = frame({
    apply: async (commands) => {
      writes.push(commands);
      await new Promise((r) => (release = r));
    },
  });
  c.setFrame(f);
  c.choose(f, 'themes', 'forest');
  const pending = c.theme(f);
  const other = frame({ key: 'other', family: 'other', sheet: 'Other', apply: async (x) => writes.push(x) });
  c.setFrame(other);
  await assert.rejects(c.background(other), /Wait/);
  assert.throws(() => c.choose(other, 'backgrounds', 'graph-paper'), /Wait/);
  release();
  assert.equal((await pending).current, false);
  assert.equal(c.getSnapshot().notice, '');
  assert.equal(c.getSnapshot().pending, false);
  assert.equal(writes.length, 1);
  assert.equal(writes[0][0].sample, 'forest');
  await assert.rejects(c.theme(f), /changed/);
  c.dispose();
});
test('pane retirement removes drafts and suppresses failed old-owner errors', async () => {
  const c = createPageStylesController();
  let reject;
  const f = frame({ apply: () => new Promise((_, r) => (reject = r)) });
  c.setFrame(f);
  c.choose(f, 'backgrounds', 'graph-paper');
  const pending = c.background(f);
  c.setFrame({ ...f, active: false });
  c.setFrame(frame({ key: 'new-pane', family: 'new-pane' }));
  reject(Error('old owner failure'));
  await assert.rejects(pending);
  assert.equal(c.getSnapshot().background, null);
  assert.equal(c.getSnapshot().error, '');
  c.dispose();
});
test('read-only, busy and imported protection remain independent of preview selection', async () => {
  const c = createPageStylesController(),
    writes = [];
  let f = frame({ apply: async (x) => writes.push(x) });
  c.setFrame(f);
  c.choose(f, 'themes', 'neutral');
  c.setFrame({ ...f, writable: false, reason: 'Busy owner' });
  await assert.rejects(c.theme(f), /Busy/);
  f = frame({
    key: 'protected',
    theme: { revision, sample: null, canApply: false, protectedReason: 'Imported theme' },
    background: { revision, sample: null, canEdit: false, protectedReason: 'Shared media' },
    apply: async (x) => writes.push(x),
  });
  c.setFrame(f);
  c.choose(f, 'themes', 'forest');
  c.choose(f, 'backgrounds', 'soft-dots');
  await assert.rejects(c.theme(f), /Imported/);
  await assert.rejects(c.background(f), /Shared/);
  await assert.rejects(c.background(f, true), /Shared/);
  assert.throws(() => c.choose(f, 'themes', 'unknown'));
  assert.deepEqual(writes, []);
  c.dispose();
});
