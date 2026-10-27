/** Closed metadata geometry only; no native helper, workbook files or app launch. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  CLEAR_HYPERLINKS_LIMITS,
  clearHyperlinksCommand,
  normalizeClearHyperlinksCommands,
} from '../core/clear-hyperlinks.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';
import { clearHyperlinksTargets } from '../ui/clearHyperlinksTargets.ts';
import {
  dispatchHomeCommand,
  homeCommandDefinition,
  homeCommandCapability,
  normalizeHomeArguments,
  captureHomeBinding,
} from '../ui/homeCommands.ts';
const command = (ranges = ['B2:D4']) => ({ type: 'clearHyperlinks', sheet: 'Sheet1', ranges });
const binding = (patch) => ({
  documentId: 'book-A',
  ownerIdentity: 'owner-A',
  state: 'accepted-1',
  sheet: 'Sheet1',
  selection: 'B2:D4',
  selectedCell: 'B2',
  activationEpoch: 1,
  navigationGeneration: 0,
  modeGeneration: 0,
  specialSelectionEpoch: 0,
  mode: 'normal',
  specialSelection: null,
  ...patch,
});
const frame = (patch) => ({
  binding: binding(),
  profile: 'xlsx',
  active: true,
  canEdit: true,
  busy: false,
  acceptBatch: async () => {},
  ...patch,
});
test('canonical Clear Hyperlinks accepts full-sheet/axis geometry without a position cap', () => {
  assert.deepEqual(CLEAR_HYPERLINKS_LIMITS, { ranges: 10000, actionBytes: 2097152 });
  for (const range of ['A1', 'A1:XFD1048576', 'A1:A1048576', 'A1:XFD1', 'XFD1048576'])
    assert.deepEqual(clearHyperlinksCommand(command([range])), command([range]));
  assert.deepEqual(clearHyperlinksCommand(command(['C20:D30', 'A2:B3', 'Z1'])).ranges, ['Z1', 'A2:B3', 'C20:D30']);
  const ranges = Array.from({ length: 10000 }, (_, i) => 'A' + (1 + i * 2));
  assert.equal(clearHyperlinksCommand(command(ranges)).ranges.length, 10000);
});
test('metadata command rejects aliases, forged fields, missing ranges and out-of-grid geometry', () => {
  for (const ranges of [
    [],
    ['a1'],
    ['$A$1'],
    ['A01'],
    ['A1:A1'],
    ['B2:A1'],
    ['XFE1'],
    ['A1048577'],
    ['Sheet1!A1'],
    ['A1 B2'],
    Array(10001).fill('A1'),
  ])
    assert.throws(() => clearHyperlinksCommand(command(ranges)));
  for (const extra of [
    { value: null },
    { format: { underline: false } },
    { fontColor: '000000' },
    { operation: 'clearRange' },
    { url: 'https://example.invalid/' },
    { relationshipId: 'rId1' },
    { mode: 'remove' },
  ])
    assert.throws(() => clearHyperlinksCommand({ ...command(), ...extra }));
  for (const value of [
    null,
    [],
    { ...command(), sheet: ' ' },
    { ...command(), sheet: 'Invalid:Sheet' },
    { ...command(), ranges: 'A1' },
  ])
    assert.throws(() => clearHyperlinksCommand(value));
});
test('geometric overlap guards include containment/corners and permit edge adjacency', () => {
  for (const ranges of [
    ['A1:B2', 'B2:C3'],
    ['A1:XFD1048576', 'XFD1048576'],
    ['A1:C3', 'B2'],
    ['C3', 'C3'],
    ['A1:A10', 'A5:B5'],
  ])
    assert.throws(() => clearHyperlinksCommand(command(ranges)), {
      code: /^(hyperlink-bad-input|internal-invalid-field)$/,
    });
  for (const ranges of [
    ['A1:B2', 'C1:D2'],
    ['A1:XFD5', 'A6:XFD1048576'],
    ['A1:A1048576', 'B1:XFD1048576'],
    ['A1:B1', 'A2:B2'],
    ['A1:C1', 'A2', 'C2', 'A3:C3'],
  ])
    assert.deepEqual(clearHyperlinksCommand(command(ranges)).ranges, ranges);
});
test('queued batch capture owns range arrays and enforces original worksheet/action bytes', () => {
  const original = [command(['C3', 'A1']), { sheet: 'Sheet1', cell: 'Z1', value: 'unchanged ordinary command' }],
    accepted = normalizeClearHyperlinksCommands(original, {
      sheets: [{ name: 'Sheet1', path: 'xl/worksheets/custom.xml' }],
    });
  original[0].ranges[0] = 'D4';
  original[1].value = 'retired';
  assert.deepEqual(accepted, [
    command(['A1', 'C3']),
    { sheet: 'Sheet1', cell: 'Z1', value: 'unchanged ordinary command' },
  ]);
  for (const sheets of [
    [],
    [{ name: 'Sheet1' }],
    [
      { name: 'Sheet1', path: 'one' },
      { name: 'Sheet1', path: 'two' },
    ],
  ])
    assert.throws(() => normalizeClearHyperlinksCommands([command()], { sheets }), /worksheet/);
  assert.throws(
    () => normalizeClearHyperlinksCommands([command(), { sheet: 'Sheet1', cell: 'Z1', value: 'x'.repeat(2097152) }]),
    /2 MiB/
  );
});
test('portable cell resolver explicitly refuses metadata removal before cell projection', () => {
  const cells = { A1: { value: 'linked text', styleId: 7, underline: true } },
    book = { sheets: [{ name: 'Sheet1', cells }] },
    before = structuredClone(book);
  assert.throws(() => resolveWorkbookEdits(book, [command(['A1'])]), /native XLSX owner/);
  assert.deepEqual(book, before);
});
test('captured sparse targets preserve holes and merge only exact cell runs', () => {
  const source = binding({ selection: 'A1:XFD1048576', specialSelection: ['C3', 'B2', 'A2', 'A1', 'B1', 'C1', 'A3'] }),
    captured = clearHyperlinksTargets(source);
  assert.deepEqual(captured, command(['A1:C1', 'A2:B2', 'A3', 'C3']));
  source.specialSelection.push('D4');
  assert.deepEqual(captured.ranges, ['A1:C1', 'A2:B2', 'A3', 'C3']);
  assert.deepEqual(clearHyperlinksTargets(binding({ specialSelection: ['B2', 'A2', 'A1', 'B1'] })), command(['A1:B2']));
  assert.deepEqual(clearHyperlinksTargets(binding({ selection: 'A1:XFD1048576' })), command(['A1:XFD1048576']));
  assert.deepEqual(clearHyperlinksTargets(binding({ selection: 'C3:C3' })), command(['C3']));
  for (const patch of [
    { ownerIdentity: null },
    { mode: 'pageLayout' },
    { specialSelection: ['A1', 'A1'] },
    { specialSelection: Array.from({ length: 10001 }, (_, i) => 'A' + (i + 1)) },
  ])
    assert.throws(() => clearHyperlinksTargets(binding(patch)));
});
test('Clear Hyperlinks registry is a closed native-XLSX editing route that permits sparse Clear', () => {
  const definition = homeCommandDefinition('home.clear-hyperlinks');
  assert.equal(definition.group, 'Editing');
  assert.equal(definition.requiresEdit, true);
  assert.equal(definition.requiresBatch, true);
  assert.equal(definition.allowSpecial, true);
  assert.deepEqual(definition.supportedProfiles, ['xlsx']);
  assert(definition.order > homeCommandDefinition('home.clear-all').order);
  assert.equal(normalizeHomeArguments('home.clear-hyperlinks'), undefined);
  assert.throws(() => normalizeHomeArguments('home.clear-hyperlinks', { ranges: ['foreign'] }));
  assert.equal(
    homeCommandCapability('home.clear-hyperlinks', frame({ binding: binding({ specialSelection: ['A1', 'C3'] }) }))
      .state,
    'enabled'
  );
  for (const patch of [
    { active: false },
    { busy: true },
    { canEdit: false },
    { acceptBatch: undefined },
    { binding: binding({ ownerIdentity: null }) },
    { binding: binding({ mode: 'pageLayout' }) },
  ])
    assert.equal(homeCommandCapability('home.clear-hyperlinks', frame(patch)).state, 'disabled');
  for (const profile of ['delimited', 'passive'])
    assert.equal(homeCommandCapability('home.clear-hyperlinks', frame({ profile })).state, 'hidden');
});
test('captured Home lease keeps one batch and refuses retired selection/head/owner', async () => {
  const original = frame();
  let current = original;
  const accepted = [];
  original.acceptBatch = async (commands) => accepted.push(structuredClone(commands));
  const scope = { captureFrame: () => original, currentFrame: () => current };
  const result = await dispatchHomeCommand('home.clear-hyperlinks', scope, (lease) =>
    lease.accept([clearHyperlinksTargets(lease.binding)])
  );
  assert.equal(result.outcome, 'completed');
  assert.deepEqual(accepted, [[command()]]);
  for (const patch of [
    { selection: 'A1' },
    { state: 'accepted-2' },
    { ownerIdentity: 'owner-B' },
    { navigationGeneration: 1 },
    { specialSelection: ['B2'], specialSelectionEpoch: 1 },
  ]) {
    current = { ...original, binding: captureHomeBinding(binding(patch)) };
    const stale = await dispatchHomeCommand('home.clear-hyperlinks', scope, (lease) => lease.accept([command()]));
    assert.equal(stale.outcome, 'stale');
  }
  assert.equal(accepted.length, 1);
});
