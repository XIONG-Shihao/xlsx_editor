import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const directory = path.resolve('.build/tests/w7-rules-component');
await fs.mkdir(directory, { recursive: true });
const output = path.join(directory, 'WorkbookRules.mjs');
await build({
  entryPoints: ['ui/WorkbookRules.tsx'],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: output,
  external: ['react', 'react/jsx-runtime'],
  logLevel: 'silent',
});
const { ConditionalRules, DataValidation } = await import(pathToFileURL(output));
const dom = new JSDOM('<!doctype html><div id="root"></div>', { url: 'http://localhost/' });
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.IS_REACT_ACT_ENVIRONMENT = true;
const { createRoot } = await import('react-dom/client');
after(() => {
  dom.window.close();
  delete globalThis.window;
  delete globalThis.document;
  delete globalThis.HTMLElement;
  delete globalThis.IS_REACT_ACT_ENVIRONMENT;
});
const empty = { sheet: 'Sheet1', revision: null, canEdit: true, protectedReasons: [], rules: [] },
  base = () => ({
    documentId: 'doc1',
    revision: 'r1',
    sheet: 'Sheet1',
    selection: 'B2:B8',
    inventory: async () => empty,
    onBatch: async () => {},
  }),
  button = (text) => Array.from(document.querySelectorAll('button')).find((node) => node.textContent === text);
async function click(text) {
  assert.ok(button(text), 'Missing ' + text);
  await act(async () => button(text).click());
}
async function mount(t, Component, props) {
  const root = createRoot(document.getElementById('root'));
  await act(async () => root.render(React.createElement(Component, props)));
  t.after(() => act(async () => root.unmount()));
  return async (next) => act(async () => root.render(React.createElement(Component, next)));
}
test('conditional Add, Edit, Delete and reorder are exact guarded collection commands', async (t) => {
  const first = { range: 'B2:B8', kind: 'greaterThan', values: [0], fillColor: '#FFF2CC', stopIfTrue: false },
    second = { ...first, kind: 'lessThan', values: [100], fillColor: '#FF0000' },
    commands = [];
  let current = empty;
  const props = {
    ...base(),
    inventory: async () => current,
    onBatch: async (batch) => {
      commands.push(batch[0]);
      current = { ...empty, revision: 'a'.repeat(64), rules: batch[0].rules };
    },
  };
  await mount(t, ConditionalRules, props);
  await click('Add rule');
  assert.deepEqual(commands[0], {
    operation: 'setConditionalRules',
    sheet: 'Sheet1',
    expectedRulesRevision: null,
    rules: [first],
  });
  await click('Reset form');
  const action = button('Edit rule 1');
  assert.ok(action);
  await click('Edit rule 1');
  await click('Apply rule');
  assert.equal(commands.at(-1).expectedRulesRevision, 'a'.repeat(64));
  await click('Delete rule 1');
  assert.equal(commands.at(-1).rules.length, 0);
});
test('conditional move actions change priority without changing the declared rule bodies', async (t) => {
  const rules = [
      { range: 'B2:B8', kind: 'equal', values: [1], fillColor: '#FF0000', stopIfTrue: true },
      { range: 'B2:B8', kind: 'between', values: [0, 5], fillColor: '#00FF00', stopIfTrue: false },
    ],
    commands = [];
  await mount(t, ConditionalRules, {
    ...base(),
    inventory: async () => ({ ...empty, revision: 'b'.repeat(64), rules }),
    onBatch: async (batch) => commands.push(batch[0]),
  });
  await act(async () => document.querySelector('[aria-label="Move rule up 2"]').click());
  assert.deepEqual(commands[0].rules, [rules[1], rules[0]]);
  assert.equal(commands[0].expectedRulesRevision, 'b'.repeat(64));
});
test('validation list authoring exposes prompts/errors and stale protected responses cannot enable a foreign document', async (t) => {
  const commands = [];
  let resolve;
  const pending = new Promise((done) => (resolve = done)),
    props = { ...base(), inventory: async () => empty, onBatch: async (batch) => commands.push(batch[0]) };
  const render = await mount(t, DataValidation, props);
  assert.equal(document.querySelector('[aria-label="List values (one per line)"]').value, 'Yes\nNo');
  assert.ok(document.querySelector('[aria-label="Input message"]'));
  assert.ok(document.querySelector('[aria-label="Error message"]'));
  await click('Add rule');
  assert.equal(commands[0].operation, 'setValidationRules');
  assert.deepEqual(commands[0].rules[0].values, ['Yes', 'No']);
  assert.equal(commands[0].rules[0].showDropdown, true);
  await render({ ...props, revision: 'r2', inventory: () => pending });
  await render({
    ...props,
    documentId: 'doc2',
    revision: 'r3',
    inventory: async () => ({
      ...empty,
      canEdit: false,
      protectedReasons: ['Imported custom validation is protected'],
    }),
  });
  await act(async () => resolve(empty));
  assert.equal(document.querySelector('fieldset').disabled, true);
  assert.match(document.body.textContent, /Imported custom validation/);
});
