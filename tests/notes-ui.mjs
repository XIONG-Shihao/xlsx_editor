import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { createNoteController } from '../ui/noteController.ts';
const out = await fs.mkdtemp(path.resolve('.build/tests/notes-ui-')),
  bundle = path.join(out, 'ui.mjs');
await build({
  stdin: { contents: 'export {WorkbookNotes} from "./ui/WorkbookNotes"', resolveDir: process.cwd(), loader: 'tsx' },
  bundle: true,
  outfile: bundle,
  format: 'esm',
  platform: 'node',
  packages: 'external',
  jsx: 'automatic',
  logLevel: 'silent',
});
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Element: dom.window.Element,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  { WorkbookNotes } = await import(pathToFileURL(bundle));
after(() => dom.window.close());
const receipt = { documentId: 'document', ownerId: 'owner', state: 'head', revision: 'revision' },
  base = () => ({
    schema: 'workbook.xlsx.legacy-notes.v1',
    sheet: 'Main',
    noteRevision: null,
    notes: [],
    canEdit: true,
    protectedReason: null,
    ...receipt,
  }),
  frame = (patch = {}) => ({
    key: 'owner/head/A1/pane',
    sheet: 'Main',
    cell: 'A1',
    receipt,
    active: true,
    writable: true,
    read: async () => base(),
    apply: async () => {},
    navigate: () => {},
    ...patch,
  }),
  props = (n) => Object.entries(n).find(([k]) => k.startsWith('__reactProps$'))[1],
  label = (s) => document.querySelector('[aria-label="' + s + '"]'),
  button = (s) => [...document.querySelectorAll('button')].find((n) => n.textContent === s),
  flush = async () => {
    for (let i = 0; i < 4; i++) await act(async () => new Promise((r) => setTimeout(r, 0)));
  },
  change = async (s, value) => {
    await act(async () => props(label(s)).onChange({ target: { value } }));
    await flush();
  },
  click = async (s) => {
    assert(button(s), s);
    await act(async () => button(s).click());
    await flush();
  },
  submit = async () => {
    await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
    await flush();
  },
  gate = () => {
    let resolve;
    const promise = new Promise((r) => (resolve = r));
    return { promise, resolve };
  };
async function mount(t, patch = {}) {
  const root = createRoot(document.getElementById('root')),
    writes = [],
    moves = [];
  let value = frame({ apply: async (c) => writes.push(...c), navigate: (...args) => moves.push(args), ...patch });
  const render = async (p) => {
    value = { ...value, ...p };
    await act(async () => root.render(React.createElement(WorkbookNotes, { frame: value })));
    await flush();
  };
  await render({});
  t.after(async () => act(async () => root.unmount()));
  return { writes, moves, render, frame: () => value };
}
const note = { cell: 'A1', author: 'Alice', text: 'Saved note' },
  saved = () => ({
    ...base(),
    noteRevision: 'a'.repeat(64),
    notes: [note, { cell: 'Z80', author: 'Bob', text: 'Off extent' }],
  });
test('draft/Reset is local; Add uses exact current cell, author/text and original graph revision', async (t) => {
  const ui = await mount(t);
  await change('Note author', 'Alice');
  await change('Note text', ' Draft <text>\n😀 ');
  assert.equal(ui.writes.length, 0);
  await click('Reset note form');
  assert.equal(label('Note text').value, '');
  await change('Note author', 'Bob');
  await change('Note text', 'New');
  await submit();
  assert.deepEqual(ui.writes, [
    { type: 'setNote', sheet: 'Main', cell: 'A1', expectedNoteRevision: null, note: { author: 'Bob', text: 'New' } },
  ]);
});
test('edit/delete retain the displayed native revision; location navigation changes no workbook content', async (t) => {
  const ui = await mount(t, { read: async () => saved() });
  assert.equal(label('Note text').value, 'Saved note');
  await change('Note author', 'Carol');
  await submit();
  assert.equal(ui.writes[0].expectedNoteRevision, 'a'.repeat(64));
  assert.equal(ui.writes[0].note.author, 'Carol');
  await click('Delete note');
  assert.deepEqual(ui.writes[1], {
    type: 'removeNote',
    sheet: 'Main',
    cell: 'A1',
    expectedNoteRevision: 'a'.repeat(64),
  });
  await act(async () => label('Go to note Z80').click());
  await flush();
  assert.deepEqual(ui.moves, [['Main', 'Z80']]);
  assert.equal(ui.writes.length, 2);
});
test('protected known notes remain readable/navigable; unknown imported graphs do not display a guessed empty list', async (t) => {
  const ui = await mount(t, {
    writable: false,
    reason: 'Read-only',
    read: async () => ({ ...saved(), canEdit: false, protectedReason: 'Protected sheet' }),
  });
  assert.equal(label('Note text').value, 'Saved note');
  assert(document.querySelector('fieldset').disabled);
  await submit();
  assert.equal(ui.writes.length, 0);
  await act(async () => label('Go to note Z80').click());
  assert.equal(ui.moves.length, 1);
  await ui.render({
    key: 'foreign',
    read: async () => ({ ...base(), notes: null, canEdit: false, protectedReason: 'Threaded graph' }),
  });
  assert.match(document.body.textContent, /Threaded graph/);
  assert.doesNotMatch(document.body.textContent, /No notes on this worksheet/);
});
test('owner/head/selection/pane retirement discards a draft and refuses a held old submit callback', async (t) => {
  const ui = await mount(t);
  await change('Note text', 'Old draft');
  const old = props(document.querySelector('form')).onSubmit;
  await ui.render({ key: 'new-cell', cell: 'B2' });
  assert.equal(label('Note text').value, '');
  await act(async () => old({ preventDefault() {} }));
  await flush();
  assert.equal(ui.writes.length, 0);
  await ui.render({ active: false });
  assert(document.querySelector('fieldset').disabled);
});
test('one physical pending write remains blocking across replacement frames and refuses duplicate submits', async (t) => {
  const hold = gate(),
    writes = [],
    ui = await mount(t, {
      apply: async (c) => {
        writes.push(...c);
        await hold.promise;
      },
    });
  await change('Note text', 'Pending');
  await act(async () => props(document.querySelector('form')).onSubmit({ preventDefault() {} }));
  await flush();
  assert.equal(writes.length, 1);
  await ui.render({ key: 'next', cell: 'B2' });
  assert(document.querySelector('fieldset').disabled);
  await submit();
  assert.equal(writes.length, 1);
  await act(async () => hold.resolve());
  await flush();
  assert(!document.querySelector('fieldset').disabled);
});
test('invalid text, reserved author labels and mismatched native replies never produce a command', async (t) => {
  const ui = await mount(t);
  await change('Note author', 'tc=thread');
  await change('Note text', 'Body');
  await submit();
  assert.equal(ui.writes.length, 0);
  assert.match(document.body.textContent, /reserved/);
  await ui.render({ key: 'bad-reply', read: async () => ({ ...base(), ownerId: 'other' }) });
  assert.match(document.body.textContent, /retired owner/);
  assert(document.querySelector('fieldset').disabled);
});
test('the source read lane joins retirement and only the latest view can expose note data or navigate', async () => {
  const hold = gate(),
    reads = [],
    moves = [],
    controller = createNoteController(),
    old = frame({
      read: async () => {
        reads.push('old');
        return hold.promise;
      },
    });
  controller.setFrame(old);
  const fresh = frame({
    key: 'latest',
    read: async () => {
      reads.push('latest');
      return saved();
    },
    navigate: (...v) => moves.push(v),
  });
  controller.setFrame(fresh);
  assert.deepEqual(reads, ['old']);
  assert.equal(controller.getSnapshot().inventory, null);
  hold.resolve(base());
  await flush();
  assert.deepEqual(reads, ['old', 'latest']);
  assert.equal(controller.getSnapshot().key, 'latest');
  assert.throws(() => controller.navigate(old, 'A1'));
  assert.throws(() => controller.navigate(fresh, 'B99'));
  controller.navigate(fresh, 'Z80');
  assert.deepEqual(moves, [['Main', 'Z80']]);
  controller.dispose();
});
