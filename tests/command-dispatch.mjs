import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { workbook } from '../fixtures/standalone.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { commandModules, directModules, guardedModules, routeCommands } from '../core/commands/index.mjs';
import { isIndexedObjectCommand } from '../core/indexed-object-authoring.mjs';
import { isIndexedPivotCommand } from '../core/indexed-pivots.mjs';
import { isIndexedRulesCommand } from '../core/indexed-rules.mjs';
import { isIndexedTableCommand } from '../core/indexed-tables.mjs';
import { isIndexedNameCommand } from '../core/indexed-names.mjs';
import { isIndexedPrintCommand } from '../core/indexed-print.mjs';
import { isNoteCommand } from '../core/notes.mjs';
after(() => closeNativeWorkbookWorker());

// A command of each kind, as the session sees it: by `type`, by `operation`, or neither.
const sample = (module, kind) => [module.handles.type ? { type: kind } : { operation: kind }];

test('every command kind is routed to exactly one module', () => {
  const seen = new Map();
  for (const module of commandModules) {
    const kinds = module.handles.type ?? module.handles.operation;
    assert.ok(kinds.length, `${module.name} handles something`);
    for (const kind of kinds) {
      const owners = commandModules.filter((candidate) => candidate.matches(sample(module, kind)));
      assert.deepEqual(
        owners.map((owner) => owner.name),
        [module.name],
        `${String(kind)} belongs to ${module.name} only`
      );
      const key = (module.handles.type ? 'type:' : 'operation:') + String(kind);
      assert.ok(!seen.has(key), `${key} is claimed by ${seen.get(key)} and ${module.name}`);
      seen.set(key, module.name);
    }
  }
  assert.equal(new Set(commandModules.map((module) => module.name)).size, commandModules.length);
  for (const module of commandModules) {
    assert.equal(typeof module.run, 'function', `${module.name} has a run function`);
    assert.equal(typeof module.matches, 'function', `${module.name} has a matcher`);
  }
});

test('the module lists agree with the command checks the session already uses', () => {
  const handled = (type) => (type ? (m) => m.handles.type : (m) => m.handles.operation);
  const operations = commandModules.flatMap((module) => handled(false)(module) ?? []);
  const types = commandModules.flatMap((module) => handled(true)(module) ?? []);
  // Each shared "is this one of mine" check from the feature files must accept exactly what a module lists.
  const byCheck = [
    [isIndexedObjectCommand, 'objects'],
    [isIndexedPivotCommand, 'pivots'],
    [isIndexedNameCommand, 'names'],
    [isIndexedPrintCommand, 'print'],
  ];
  for (const [check, name] of byCheck) {
    const module = commandModules.find((candidate) => candidate.name === name);
    for (const operation of module.handles.operation) assert.ok(check({ operation }), `${operation} passes ${name}`);
    for (const operation of operations)
      assert.equal(check({ operation }), module.handles.operation.includes(operation), `${operation} vs ${name}`);
  }
  const rules = commandModules.find((module) => module.name === 'rules');
  for (const operation of ['setConditionalRules', 'setValidationRules'])
    assert.ok(isIndexedRulesCommand({ operation }));
  assert.ok(!isIndexedRulesCommand({ operation: 'addConditionalRule' }));
  assert.ok(rules.handles.operation.includes('addConditionalRule'));
  const tables = commandModules.find((module) => module.name === 'tables');
  for (const operation of tables.handles.operation.filter((kind) => kind !== 'createTable'))
    assert.ok(isIndexedTableCommand({ operation }));
  for (const type of ['setNote', 'removeNote']) assert.ok(isNoteCommand({ type }) && types.includes(type));
});

test('ordinary cell edits go to the cell-edits module and nothing else claims them', () => {
  const plain = [
    [{ sheet: 'Sheet1', cell: 'A1', value: 1 }],
    [{ type: 'setCells', sheet: 'Sheet1', cells: [] }],
    [{ type: 'formatRange', sheet: 'Sheet1', range: 'A1', format: {} }],
    [{ type: 'pasteValues', sheet: 'Sheet1', range: 'A1:B2' }],
    [
      { sheet: 'Sheet1', cell: 'A1', value: 1 },
      { type: 'fillRange', sheet: 'Sheet1', range: 'A1:A3' },
    ],
  ];
  for (const commands of plain)
    assert.deepEqual(
      commandModules.filter((module) => module.matches(commands)).map((module) => module.name),
      ['cell-edits']
    );
});

test('the pivot and validation guards sit between the two groups of modules', () => {
  const direct = directModules.map((module) => module.name);
  assert.deepEqual(direct, [
    'sheet-background',
    'page-theme',
    'boolean-cell',
    'notes',
    'hyperlinks',
    'clear-hyperlinks',
    'calculation-options',
    'recalculate',
  ]);
  assert.equal(guardedModules.at(-1).name, 'cell-edits');
  assert.equal(routeCommands(directModules, [{ operation: 'sortRows' }]), undefined);
  assert.equal(routeCommands(guardedModules, [{ operation: 'sortRows' }]).name, 'sort-filter');
  assert.equal(routeCommands(commandModules, [{ type: 'notACommand' }]), undefined);
  assert.equal(routeCommands(commandModules, [{ operation: 'notAnOperation' }]), undefined);
  assert.equal(routeCommands(commandModules, [{ type: 'recalculate' }, { type: 'recalculate' }]), undefined);
});

test('an unknown command kind is refused and the workbook stays as it was', async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'xlsx-dispatch-')),
    filename = path.join(root, 'document.xlsx'),
    owner = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery') });
  t.after(async () => {
    await owner.close();
    await rm(root, { recursive: true, force: true });
  });
  await writeFile(filename, workbook());
  await owner.open(filename);
  const before = owner.status().state;
  for (const command of [
    { type: 'notACommand', sheet: 'Sheet1' },
    { operation: 'notAnOperation', sheet: 'Sheet1' },
  ])
    await assert.rejects(
      owner.request({ operation: 'apply', state: before, commands: [command] }),
      /no qualified indexed implementation; accepted workbook is unchanged/
    );
  assert.equal(owner.status().state, before);
});
