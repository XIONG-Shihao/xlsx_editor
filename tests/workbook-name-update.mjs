import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { indexedNamePlan } from '../core/indexed-names.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { checkXlsxPackage } from '../scripts/test/package-oracle.mjs';
after(() => closeNativeWorkbookWorker());
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'name-update-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['custom/opaque.bin'] = Uint8Array.from([0, 255, 1]);
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace('<sheets>', '<!-- preserved boundary -->\n<sheets>')
  );
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source);
  let root;
  const bridge = async (input) => {
    if (input.operation === 'index') root = input.root;
    return native(input);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No fallback writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(source);
  t.after(() => book.close());
  return { directory, source, parts, original, book, root: () => root };
}
const create = (scope) => ({
  operation: 'defineName',
  sheet: 'Sheet1',
  name: 'Inputs',
  range: 'A1:B4',
  nameScope: scope,
  expectedNameFormula: null,
});
const update = (scope) => ({
  operation: 'updateName',
  sheet: 'Sheet1',
  name: 'Inputs',
  range: 'C3:E9',
  nameScope: scope,
  expectedNameFormula: "'Sheet1'!$A$1:$B$4",
});
for (const scope of ['workbook', 'worksheet'])
  test(`isolated ${scope} static name range updates under prior guard and independently saved XML`, async (t) => {
    const f = await fixture(t),
      made = await f.book.apply('0', [create(scope)]),
      before = made.view.storage.state,
      changed = await f.book.apply(before, [update(scope)]),
      state = changed.view.storage.state;
    assert.equal(changed.view.namedRanges[0].formula, "'Sheet1'!$C$3:$E$9");
    assert.equal(f.book.view(before).namedRanges[0].formula, "'Sheet1'!$A$1:$B$4");
    assert.deepEqual(f.book.view('0').namedRanges, []);
    await assert.rejects(f.book.apply(state, [update(scope)]), /Conflict/);
    const replay = await f.book.apply(before, changed.edits);
    assert.deepEqual(replay.view.namedRanges, changed.view.namedRanges);
    const old = unzipSync(await fs.readFile((await f.book.export(before)).path)),
      saved = (await f.book.export(state)).path,
      bytes = await fs.readFile(saved),
      parts = unzipSync(bytes);
    assert.equal(
      strFromU8(parts['xl/workbook.xml']),
      strFromU8(old['xl/workbook.xml']).replace("'Sheet1'!$A$1:$B$4", "'Sheet1'!$C$3:$E$9")
    );
    for (const part of Object.keys(old)) if (part !== 'xl/workbook.xml') assert.deepEqual(parts[part], old[part], part);
    const beforeFile = path.join(f.directory, 'before.xlsx');
    await fs.writeFile(beforeFile, zipSync(old));
    assert.equal(checkXlsxPackage(beforeFile, saved, { expectedNames: { Inputs: "'Sheet1'!$C$3:$E$9" } }).passed, true);
    await fs.copyFile(saved, path.join(f.directory, 'saved.xlsx'));
    assert.deepEqual(await fs.readFile(f.source), f.original);
  });
test('static name update after ordinary value edits preserves those accepted values', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [create('workbook')]),
    edited = await f.book.apply(made.view.storage.state, [{ sheet: 'Sheet1', cell: 'D3', value: 25 }]),
    before = edited.view.storage.state,
    changed = await f.book.apply(before, [update('workbook')]);
  assert.equal((await f.book.read(changed.view.storage.state, 'Sheet1', 'D3')).cells.D3.value, 25);
  assert.equal((await f.book.read(before, 'Sheet1', 'D3')).cells.D3.value, 25);
  await f.book.export(changed.view.storage.state);
});
test('referenced names can retarget; used removal, reserved and extra fields refuse; neighbors survive', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [create('workbook')]),
    state = made.view.storage.state,
    meta = await fs.readFile(path.join(f.root(), 'metadata.xlsx'));
  for (const cmd of [
    { ...update('workbook'), expectedNameFormula: 'changed' },
    { ...update('workbook'), name: '_xlnm.Print_Area' },
    { ...update('workbook'), unsafe: 1 },
  ])
    await assert.rejects(f.book.apply(state, [cmd]));
  const formula = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', formula: 'SUM(Inputs)' }]);
  const retargeted = await f.book.apply(formula.view.storage.state, [update('workbook')]);
  assert.equal(retargeted.view.namedRanges[0].formula, "'Sheet1'!$C$3:$E$9");
  await assert.rejects(
    f.book.apply(retargeted.view.storage.state, [
      { operation: 'removeName', sheet: 'Sheet1', name: 'Inputs', expectedNameFormula: "'Sheet1'!$C$3:$E$9" },
    ]),
    { code: 'name-unsupported' }
  );
  const another = await f.book.apply(state, [{ ...create('workbook'), name: 'Other' }]);
  const resized = await f.book.apply(another.view.storage.state, [update('workbook')]);
  assert.equal(resized.view.namedRanges.find((n) => n.name === 'Inputs').formula, "'Sheet1'!$C$3:$E$9");
  assert.equal(resized.view.namedRanges.find((n) => n.name === 'Other').formula, "'Sheet1'!$A$1:$B$4");
  const parts = unzipSync(meta);
  parts['xl/drawings/unrelated.xml'] = strToU8('<drawing/>');
  await assert.rejects(
    Promise.resolve().then(() => indexedNamePlan(zipSync(parts), f.book.view(state), update('workbook'))),
    { code: 'name-unsupported' }
  );
  assert.equal(f.book.view(state).namedRanges[0].formula, "'Sheet1'!$A$1:$B$4");
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('native name-update guard independently refuses changed scope, prior formula and unrelated bytes', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [create('workbook')]),
    state = made.view.storage.state,
    root = f.root(),
    exported = await f.book.export(state),
    plan = indexedNamePlan(await fs.readFile(exported.path), f.book.view(state), update('workbook'));
  const attempt = async (value) => {
    await fs.writeFile(path.join(root, 'package-metadata.json'), JSON.stringify(value));
    return native({
      action: 'workbook',
      operation: 'package-edit',
      root,
      version: 0,
      source: path.basename(exported.path),
    });
  };
  await attempt(plan);
  for (const mutate of [
    (p) => {
      p.scope = 0;
    },
    (p) => {
      p.priorFormula = 'wrong';
    },
    (p) => {
      p.replacements['xl/workbook.xml'] = p.replacements['xl/workbook.xml'].replace(
        'preserved boundary',
        'forged boundary'
      );
    },
    (p) => {
      p.formula = "'Missing'!$C$3:$E$9";
      p.replacements['xl/workbook.xml'] = p.replacements['xl/workbook.xml'].replace("'Sheet1'!$C$3:$E$9", p.formula);
    },
    (p) => {
      p.replacements['custom/opaque.bin'] = 'changed';
    },
  ]) {
    const changed = structuredClone(plan);
    mutate(changed);
    await assert.rejects(attempt(changed), /name|Name|unrelated|Static|static/);
  }
  assert.equal(f.book.view(state).namedRanges[0].formula, "'Sheet1'!$A$1:$B$4");
});
test('durable name update preserves one-step Undo/Redo and clean-cursor recovery', async (t) => {
  const f = await fixture(t),
    recoveryDirectory = path.join(f.directory, 'recovery');
  let owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null });
  t.after(async () => owner.close());
  let status = await owner.open(f.source);
  status = await owner.request({ operation: 'apply', state: status.state, commands: [create('worksheet')] });
  status = await owner.request({ operation: 'apply', state: status.state, commands: [update('worksheet')] });
  assert.equal(status.view.namedRanges[0].formula, "'Sheet1'!$C$3:$E$9");
  status = await owner.request({ operation: 'undo', state: status.state });
  assert.equal(status.view.namedRanges[0].formula, "'Sheet1'!$A$1:$B$4");
  await owner.close();
  owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null });
  status = await owner.open(f.source);
  assert.equal(status.recovered, true);
  assert.equal(status.canRedo, true);
  status = await owner.request({ operation: 'redo', state: status.state });
  assert.equal(status.view.namedRanges[0].formula, "'Sheet1'!$C$3:$E$9");
  assert.deepEqual(await fs.readFile(f.source), f.original);
});

test('workbook/local shadows and several ordinary names resize/remove without rewriting neighbor XML', async (t) => {
  const f = await fixture(t);
  let state = '0';
  for (const cmd of [create('workbook'), create('worksheet'), { ...create('workbook'), name: 'Other', range: 'D1:D2' }])
    state = (await f.book.apply(state, [cmd])).view.storage.state;
  const before = state,
    prior = f.book.view(state).namedRanges,
    local = prior.find((n) => n.name === 'Inputs' && n.sheet === 'Sheet1');
  const changed = await f.book.apply(state, [{ ...update('worksheet'), expectedNameFormula: local.formula }]);
  state = changed.view.storage.state;
  assert.equal(changed.view.namedRanges.find((n) => n.name === 'Inputs' && !n.sheet).formula, "'Sheet1'!$A$1:$B$4");
  assert.equal(changed.view.namedRanges.find((n) => n.name === 'Inputs' && n.sheet).formula, "'Sheet1'!$C$3:$E$9");
  const removed = await f.book.apply(state, [
    {
      operation: 'removeName',
      sheet: 'Sheet1',
      name: 'inputs',
      nameScope: 'worksheet',
      expectedNameFormula: "'Sheet1'!$C$3:$E$9",
    },
  ]);
  state = removed.view.storage.state;
  assert.equal(removed.view.namedRanges.length, 2);
  assert(removed.view.namedRanges.some((n) => n.name === 'Inputs' && !n.sheet));
  assert.deepEqual(f.book.view(before).namedRanges, prior);
  const saved = await f.book.export(state),
    parts = unzipSync(await fs.readFile(saved.path)),
    xml = strFromU8(parts['xl/workbook.xml']);
  assert.match(xml, /preserved boundary/);
  assert.ok(!xml.includes('localSheetId='));
  assert.ok(xml.includes('<definedName name="Other">'));
  assert.deepEqual(parts['custom/opaque.bin'], f.parts['custom/opaque.bin']);
});
test('native removal plan binds name/scope/prior formula and refuses sibling/envelope changes before output', async (t) => {
  const f = await fixture(t);
  let state = (await f.book.apply('0', [create('workbook')])).view.storage.state;
  state = (await f.book.apply(state, [{ ...create('worksheet'), name: 'Other' }])).view.storage.state;
  const exported = await f.book.export(state),
    parts = unzipSync(await fs.readFile(exported.path)),
    prepared = indexedNamePlan(zipSync(parts), f.book.view(state), {
      operation: 'removeName',
      sheet: 'Sheet1',
      name: 'Inputs',
      nameScope: 'workbook',
      expectedNameFormula: "'Sheet1'!$A$1:$B$4",
    }),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('prior complete output');
  await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(prepared));
  await native({
    action: 'workbook',
    operation: 'package-edit',
    root: f.root(),
    version: 0,
    source: path.basename(exported.path),
  });
  const mutators = [
    (p) => (p.name = 'Other'),
    (p) => (p.scope = 0),
    (p) => (p.priorFormula = 'wrong'),
    (p) =>
      (p.replacements['xl/workbook.xml'] = p.replacements['xl/workbook.xml'].replace(
        'preserved boundary',
        'forged boundary'
      )),
    (p) =>
      (p.replacements['xl/workbook.xml'] = p.replacements['xl/workbook.xml'].replace('name="Other"', 'name="Changed"')),
    (p) => (p.formula = 'unexpected'),
  ];
  for (const mutate of mutators) {
    const plan = structuredClone(prepared);
    mutate(plan);
    await fs.writeFile(output, sentinel);
    await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(plan));
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version: f.book.view(state).storage.version ?? 0,
        source: path.basename(exported.path),
      }),
      /name|Name|unrelated|identity|Invalid/
    );
    assert.deepEqual(await fs.readFile(output), sentinel);
  }
});
test('multiple-name changes keep dependent or ambiguous neighbors protected', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [create('workbook')]),
    state = made.view.storage.state,
    parts = unzipSync(await fs.readFile(path.join(f.root(), 'metadata.xlsx')));
  for (const extra of [
    "<definedName name='Dependent'>SUM(Inputs)</definedName>",
    "<definedName name='Inputs'>'Sheet1'!$C$1:$C$2</definedName>",
    "<definedName name='Bad' localSheetId='999'>'Sheet1'!$C$1:$C$2</definedName>",
  ]) {
    const m = { ...parts };
    m['xl/workbook.xml'] = strToU8(
      strFromU8(parts['xl/workbook.xml']).replace('</definedNames>', extra.replaceAll("\\'", "'") + '</definedNames>')
    );
    assert.throws(
      () => indexedNamePlan(zipSync(m), f.book.view(state), update('workbook')),
      /Dependent|dependent|Ambiguous|scope/
    );
  }
});
test('multiple independent names have real durable one-step resize/remove Undo and recovery', async (t) => {
  const f = await fixture(t),
    recoveryDirectory = path.join(f.directory, 'multi-recovery');
  let owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null });
  t.after(async () => owner.close());
  let status = await owner.open(f.source);
  for (const cmd of [create('workbook'), { ...create('worksheet'), name: 'LocalInput' }])
    status = await owner.request({ operation: 'apply', state: status.state, commands: [cmd] });
  const before = status;
  status = await owner.request({ operation: 'apply', state: status.state, commands: [update('workbook')] });
  const resized = status;
  status = await owner.request({ operation: 'undo', state: status.state });
  assert.deepEqual(status.view.namedRanges, before.view.namedRanges);
  status = await owner.request({ operation: 'redo', state: status.state });
  assert.deepEqual(status.view.namedRanges, resized.view.namedRanges);
  status = await owner.request({
    operation: 'apply',
    state: status.state,
    commands: [
      {
        operation: 'removeName',
        sheet: 'Sheet1',
        name: 'LocalInput',
        nameScope: 'worksheet',
        expectedNameFormula: "'Sheet1'!$A$1:$B$4",
      },
    ],
  });
  assert.equal(status.view.namedRanges.length, 1);
  status = await owner.request({ operation: 'undo', state: status.state });
  assert.deepEqual(status.view.namedRanges, resized.view.namedRanges);
  await owner.close();
  owner = new SpreadsheetDocument({ recoveryDirectory, native, cacheIdentity: null });
  status = await owner.open(f.source);
  assert.equal(status.canRedo, true);
  status = await owner.request({ operation: 'redo', state: status.state });
  assert.equal(status.view.namedRanges.length, 1);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
