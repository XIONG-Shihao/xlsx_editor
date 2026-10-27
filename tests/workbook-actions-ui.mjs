import './support/scratch-dir.mjs';
/** Actual Local Actions React pane with controlled private-store promises.
 * No native/workbook execution, real private persistence or Electron claim. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/actions-ui-')),
  checks = [];
const bundle = path.join(output, 'pane.mjs');
await build({
  entryPoints: [path.join(component, 'ui/WorkbookActions.tsx')],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: bundle,
  external: ['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'],
  logLevel: 'silent',
});
const { WorkbookActions } = await import(pathToFileURL(bundle));
const dom = new JSDOM('<input id="parked" value="=SUM("><div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client');
const flush = () => act(async () => new Promise((resolve) => setTimeout(resolve, 0))),
  props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps$'))],
  field = (name) => document.querySelector(`[aria-label=${JSON.stringify(name)}]`),
  button = (name) =>
    [...document.querySelectorAll('button')].find(
      (node) => (node.getAttribute('aria-label') ?? node.textContent) === name
    );
const click = async (name) => {
  const node = typeof name === 'string' ? button(name) : name;
  assert(node);
  assert(!node.disabled);
  await act(async () => node.click());
  await flush();
  return node;
};
const change = async (name, value) => {
  const node = typeof name === 'string' ? field(name) : name;
  assert(node);
  await act(async () => props(node).onChange({ target: { value } }));
  await flush();
};
const submit = async () => {
  const form = document.querySelector('.sheet-actions form');
  assert(form);
  await act(async () => props(form).onSubmit({ preventDefault() {} }));
  await flush();
};
const defer = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};
const freeze = (name = 'Freeze headings', sheet = 'Report') => ({
  name,
  descriptor: { schema: 1, command: 'freeze', parameters: { sheet, rows: 1, columns: 0 } },
});
const list = (entries = [], revision = 0) => ({ schema: 1, scope: 'private-profile', revision, entries });
const entry = (draft = freeze(), revision = 1, id = randomUUID()) => ({
  ...structuredClone(draft),
  id,
  revision,
  author: 'local-profile',
  createdAt: 1000,
  updatedAt: 1000 + revision,
});
function store(initial = list()) {
  let value = structuredClone(initial);
  const calls = [],
    overrides = {};
  const controls = {
    list: async () => {
      calls.push({ type: 'list' });
      return overrides.list ? overrides.list() : structuredClone(value);
    },
    create: async (draft, revision) => {
      calls.push({ type: 'create', draft: structuredClone(draft), revision });
      if (overrides.create) return overrides.create(draft, revision);
      assert.equal(revision, value.revision);
      value = list([...value.entries, entry(draft, value.revision + 1)], value.revision + 1);
      return structuredClone(value);
    },
    update: async (id, draft, revision) => {
      calls.push({ type: 'update', id, draft: structuredClone(draft), revision });
      if (overrides.update) return overrides.update(id, draft, revision);
      assert.equal(revision, value.revision);
      value = list(
        value.entries.map((old) =>
          old.id === id
            ? { ...old, ...structuredClone(draft), revision: old.revision + 1, updatedAt: old.updatedAt + 1 }
            : old
        ),
        value.revision + 1
      );
      return structuredClone(value);
    },
    remove: async (id, revision) => {
      calls.push({ type: 'remove', id, revision });
      if (overrides.remove) return overrides.remove(id, revision);
      assert.equal(revision, value.revision);
      value = list(
        value.entries.filter((old) => old.id !== id),
        value.revision + 1
      );
      return structuredClone(value);
    },
  };
  return {
    controls,
    calls,
    overrides,
    value: () => structuredClone(value),
    replace: (next) => (value = structuredClone(next)),
  };
}
async function mount(t, s = store(), patch = {}) {
  const root = createRoot(document.getElementById('root'));
  let input = {
      frame: { key: 'owner-head-activation-A', sheet: 'Report', controls: s.controls },
      visible: true,
      ...patch,
    },
    mounted = true;
  const render = async (patch) => {
    input = { ...input, ...patch };
    await act(async () => root.render(React.createElement(WorkbookActions, input)));
    await flush();
  };
  const unmount = async () => {
    if (!mounted) return;
    mounted = false;
    await act(async () => root.unmount());
    await flush();
  };
  await render({});
  t.after(unmount);
  return { s, render, unmount, input: () => input };
}
const check = (name, fn) => {
  test(name, { timeout: 15000 }, async (t) => {
    try {
      await fn(t);
    } catch (error) {
      await fs.writeFile(path.join(output, 'failure-' + checks.length + '.html'), document.documentElement.outerHTML);
      throw error;
    }
  });
};
const writes = (s) => s.calls.filter((call) => call.type !== 'list');
after(async () => {
  dom.window.close();
});

check('Unavailable host has an honest library message and no fabricated create, code or Run control', async (t) => {
  const s = store(),
    h = await mount(t, s, { frame: { key: 'A' } });
  assert.match(field('Local Actions').textContent, /has not supplied a private action library/);
  assert.equal(button('New action'), undefined);
  assert.equal(document.querySelector('textarea'), null);
  assert.equal(button('Run'), undefined);
  assert.deepEqual(s.calls, []);
  assert.equal(h.input().frame.key, 'A');
});
check(
  'Opening/hidden lifetime uses explicit list reads only and pure form Cancel leaves parked draft untouched',
  async (t) => {
    const s = store(),
      parked = document.getElementById('parked');
    parked.focus();
    parked.setSelectionRange(5, 5);
    const h = await mount(t, s, { visible: false });
    assert.deepEqual(s.calls, []);
    await h.render({ visible: true });
    assert.equal(s.calls.length, 1);
    assert.equal(s.calls[0].type, 'list');
    assert(field('Local Actions').hasAttribute('data-formula-draft-zone'));
    await click('New action');
    await change('Action name', 'Draft');
    await change('Rows to freeze', '4');
    await click('Cancel action');
    assert.deepEqual(writes(s), []);
    assert.equal(parked.value, '=SUM(');
    assert.equal(parked.selectionStart, 5);
    assert.equal(document.activeElement, parked);
    assert.equal(field('Local Actions').dataset.actionMode, 'list');
  }
);
check(
  'Freeze form uses exact source-sheet default and bounded parameters; one create, edit and delete use list CAS pins',
  async (t) => {
    const s = store();
    await mount(t, s);
    await click('New action');
    assert.equal(field('Action worksheet').value, 'Report');
    await change('Action name', 'Freeze report');
    await change('Rows to freeze', '100');
    await change('Columns to freeze', '30');
    await submit();
    assert.deepEqual(writes(s)[0], {
      type: 'create',
      draft: {
        name: 'Freeze report',
        descriptor: { schema: 1, command: 'freeze', parameters: { sheet: 'Report', rows: 100, columns: 30 } },
      },
      revision: 0,
    });
    const id = s.value().entries[0].id;
    await click('Edit action Freeze report');
    await change('Action name', 'Freeze first row');
    await change('Rows to freeze', '1');
    await change('Columns to freeze', '0');
    await submit();
    assert.equal(writes(s)[1].type, 'update');
    assert.equal(writes(s)[1].id, id);
    assert.equal(writes(s)[1].revision, 1);
    await click('Delete action Freeze first row');
    assert.deepEqual(writes(s)[2], { type: 'remove', id, revision: 2 });
    assert.equal(field('Saved actions').children.length, 0);
    assert.equal(field('Local Actions').dataset.actionRegistryRevision, '3');
  }
);
check(
  'Format tri-state fields emit exact canonical parameters and preserve other existing saved formatting',
  async (t) => {
    // A JSON boundary may reorder object members; ordered sort arrays still remain ordered.
    const existing = entry({
        name: 'Saved style',
        descriptor: {
          schema: 1,
          command: 'formatRange',
          parameters: {
            sheet: 'Report',
            range: 'A1:B2',
            format: { fontFamily: 'Arial', fontSize: 12, underline: true, bold: true },
          },
        },
      }),
      s = store(list([existing], 1));
    await mount(t, s);
    await click('Edit action Saved style');
    assert.equal(field('Action bold').value, 'true');
    assert.equal(field('Action italic').value, '');
    assert.match(field('Local Actions').textContent, /Font family: Arial/);
    await change('Action bold', 'false');
    await change('Action italic', 'true');
    await change('Action number format', 'date');
    await change('Action range', '$a$2:$b$3');
    await submit();
    assert.deepEqual(writes(s)[0].draft.descriptor, {
      schema: 1,
      command: 'formatRange',
      parameters: {
        sheet: 'Report',
        range: 'A2:B3',
        format: { fontFamily: 'Arial', fontSize: 12, underline: true, bold: false, italic: true, numberFormat: 'date' },
      },
    });
    await click('Edit action Saved style');
    s.overrides.update = async (id, draft, revision) => {
      const reversed = JSON.parse(
        JSON.stringify(draft, (_key, value) =>
          value && typeof value === 'object' && !Array.isArray(value)
            ? Object.fromEntries(Object.entries(value).reverse())
            : value
        )
      );
      return list([{ ...s.value().entries[0], ...reversed, id, revision: 3, updatedAt: 1005 }], revision + 1);
    };
    await change('Action bold', 'true');
    await submit();
    assert.equal(field('Local Actions').dataset.actionMode, 'list');
    assert.equal(document.querySelector('[role=alert]'), null);
    assert.equal(writes(s).length, 2);
  }
);
check(
  'Ordered sort form adds, moves, removes and independently directs keys; duplicate/outside keys cannot save',
  async (t) => {
    const s = store();
    await mount(t, s);
    await click('New action');
    await change('Action name', 'Sort report');
    await change('Action command', 'sortRows');
    await change('Action range', 'A2:C10');
    await change('Action sort column 1', 'B');
    await click('Add action level');
    await change('Action sort column 2', 'A');
    await click('Add action level');
    await change('Action sort column 3', 'C');
    await change('Action sort order 3', 'desc');
    await click('Move action level 3 up');
    await click('Move action level 2 up');
    await click('Remove action level 2');
    await click('Add action level');
    await change('Action sort column 3', 'C');
    assert(button('Save action').disabled);
    assert.match(field('Action definition issue').textContent, /distinct|duplicate/i);
    await change('Action sort column 3', 'D');
    assert(button('Save action').disabled);
    await change('Action sort column 3', 'B');
    await change('Action sort order 3', 'desc');
    await submit();
    assert.deepEqual(writes(s)[0].draft.descriptor, {
      schema: 1,
      command: 'sortRows',
      parameters: {
        sheet: 'Report',
        range: 'A2:C10',
        keys: [
          { column: 'C', direction: 'desc' },
          { column: 'A', direction: 'asc' },
          { column: 'B', direction: 'desc' },
        ],
      },
    });
    assert.equal(writes(s).length, 1);
  }
);
check('Recalculate descriptor has no invented worksheet, scope, formula, host or permission parameter', async (t) => {
  const s = store();
  await mount(t, s);
  await click('New action');
  await change('Action name', 'Calculate workbook');
  await change('Action command', 'recalculate');
  assert.equal(field('Action worksheet'), null);
  assert.equal(field('Action range'), null);
  await submit();
  assert.deepEqual(writes(s)[0].draft, {
    name: 'Calculate workbook',
    descriptor: { schema: 1, command: 'recalculate', parameters: {} },
  });
  assert.equal(button('Run'), undefined);
  assert.match(field('Local Actions').textContent, /Running saved actions is unavailable/);
});
check('Blank names, numeric Freeze limits and empty formatting changes refuse before any store mutation', async (t) => {
  const s = store();
  await mount(t, s);
  await click('New action');
  assert(button('Save action').disabled);
  await change('Action name', 'Bounded');
  for (const [name, value] of [
    ['Rows to freeze', '101'],
    ['Rows to freeze', '-1'],
    ['Rows to freeze', '1.5'],
    ['Columns to freeze', '31'],
  ]) {
    await change(name, value);
    assert(button('Save action').disabled);
    await submit();
    assert.deepEqual(writes(s), []);
  }
  await change('Rows to freeze', '1');
  await change('Columns to freeze', '0');
  await change('Action command', 'formatRange');
  await change('Action range', 'A1:B2');
  assert(button('Save action').disabled);
  await change('Action bold', 'false');
  assert(!button('Save action').disabled);
  await change('Action range', 'A1:XFD1048576');
  assert(button('Save action').disabled);
  await submit();
  assert.deepEqual(writes(s), []);
});
check(
  'Duplicate/older open intents preserve drafts and same-key sheet changes cannot retarget an edited worksheet',
  async (t) => {
    const s = store(),
      h = await mount(t, s, { openIntent: { serial: 1, mode: 'new' } });
    await change('Action name', 'Retain me');
    await change('Rows to freeze', '4');
    await h.render({ openIntent: { serial: 1, mode: 'list' }, frame: { ...h.input().frame, sheet: 'Other' } });
    assert.equal(field('Action name').value, 'Retain me');
    assert.equal(field('Action worksheet').value, 'Report');
    await h.render({ openIntent: { serial: 0, mode: 'list' } });
    assert.equal(field('Action name').value, 'Retain me');
    await h.render({ openIntent: { serial: 2, mode: 'new' } });
    assert.equal(field('Action name').value, 'Retain me');
    assert.equal(field('Rows to freeze').value, '4');
    await click('Cancel action');
    await click('New action');
    assert.equal(field('Action worksheet').value, 'Other');
    assert.deepEqual(writes(s), []);
  }
);
check(
  'Late original-frame lists are masked and the next frame waits for the physical slot, then loads its own library',
  async (t) => {
    const held = defer(),
      old = store();
    old.overrides.list = () => held.promise;
    const h = await mount(t, old),
      fresh = store(list([entry(freeze('New owner library'))], 1));
    await h.render({ frame: { key: 'new-owner-head', sheet: 'Next', controls: fresh.controls } });
    assert.equal(fresh.calls.length, 0);
    assert(!document.body.textContent.includes('Old owner library'));
    await act(async () => held.resolve(list([entry(freeze('Old owner library'))], 1)));
    await flush();
    await flush();
    assert.equal(fresh.calls.length, 1);
    assert(field('Saved actions').textContent.includes('New owner library'));
    assert(!document.body.textContent.includes('Old owner library'));
    assert.deepEqual(writes(old), []);
    assert.deepEqual(writes(fresh), []);
  }
);
check(
  'Saving captures an immutable draft and original callback/CAS; duplicate submit and pending mode intent create no extra write',
  async (t) => {
    const held = defer(),
      s = store();
    s.overrides.create = () => held.promise;
    const h = await mount(t, s);
    await click('New action');
    await change('Action name', 'Captured');
    const form = document.querySelector('form'),
      submitHandler = props(form).onSubmit,
      oldChange = props(field('Action name')).onChange;
    await act(async () => {
      submitHandler({ preventDefault() {} });
      submitHandler({ preventDefault() {} });
      oldChange({ target: { value: 'Changed while pending' } });
    });
    assert.equal(writes(s).length, 1);
    assert.equal(field('Action name').value, 'Captured');
    assert(button('Save action').disabled);
    await h.render({ openIntent: { serial: 1, mode: 'list' } });
    assert.equal(field('Local Actions').dataset.actionMode, 'new');
    const submitted = writes(s)[0].draft;
    await act(async () => held.resolve(list([entry(submitted)], 1)));
    await flush();
    assert.match(field('Local Actions').textContent, /definition saved/);
    assert.equal(writes(s).length, 1);
  }
);
check(
  'Late save failures/results stay with original frame, and retained handlers cannot disturb a replacement draft',
  async (t) => {
    for (const failure of [false, true]) {
      const held = defer(),
        old = store();
      old.overrides.create = () => held.promise;
      const h = await mount(t, old);
      await click('New action');
      await change('Action name', 'Original');
      const retainedNew = props(button('New action')).onClick,
        retainedSave = props(document.querySelector('form')).onSubmit;
      await submit();
      const fresh = store();
      await h.render({ frame: { key: 'replacement-' + failure, sheet: 'Fresh', controls: fresh.controls } });
      await act(async () =>
        failure
          ? held.reject(Error('Obsolete private save failure'))
          : held.resolve(list([entry(writes(old)[0].draft)], 1))
      );
      await flush();
      await flush();
      await click('New action');
      await change('Action name', 'New draft');
      await act(async () => {
        retainedNew();
        retainedSave({ preventDefault() {} });
      });
      await flush();
      assert.equal(field('Action name').value, 'New draft');
      assert.equal(field('Action worksheet').value, 'Fresh');
      assert(!document.body.textContent.includes('Obsolete private save failure'));
      assert.deepEqual(writes(fresh), []);
      assert.equal(writes(old).length, 1);
      await h.unmount();
    }
  }
);
check(
  'Save conflict preserves exact editable draft; refreshed changed entry retires edit rather than rebasing it silently',
  async (t) => {
    const old = entry(),
      s = store(list([old], 1));
    s.overrides.update = async () => {
      throw Error('Private library changed on disk.');
    };
    await mount(t, s);
    await click('Edit action Freeze headings');
    await change('Action name', 'Unsaved draft');
    await change('Rows to freeze', '5');
    await submit();
    assert.equal(field('Action name').value, 'Unsaved draft');
    assert.equal(field('Rows to freeze').value, '5');
    assert.match(document.querySelector('[role=alert]').textContent, /changed on disk/);
    s.replace(list([{ ...old, name: 'External update', revision: 2, updatedAt: 1003 }], 2));
    await click('Reload library');
    assert.equal(field('Action name').value, 'Unsaved draft');
    assert(button('Save action').disabled);
    assert.match(field('Local Actions').textContent, /Cancel the edit and select its current definition/);
    await submit();
    assert.equal(writes(s).length, 1);
    await click('Cancel action');
    await click('Edit action External update');
    assert.equal(field('Action name').value, 'External update');
  }
);
check(
  'Closing masks a completed write reply; reopen reads authoritative registry and no hidden handler issues a request',
  async (t) => {
    const held = defer(),
      s = store();
    s.overrides.create = () => held.promise;
    const h = await mount(t, s);
    await click('New action');
    await change('Action name', 'Saved on original frame');
    await submit();
    await h.render({ visible: false });
    assert.equal(field('Saved actions'), null);
    assert.match(field('Local Actions').textContent, /paused/);
    const accepted = list([entry(writes(s)[0].draft)], 1);
    s.replace(accepted);
    await act(async () => held.resolve(accepted));
    await flush();
    assert(!document.body.textContent.includes('definition saved'));
    await h.render({ visible: true });
    assert.equal(s.calls.filter((call) => call.type === 'list').length, 2);
    assert.equal(field('Action name').value, 'Saved on original frame');
    await click('Cancel action');
    assert(field('Saved actions').textContent.includes('Saved on original frame'));
    assert.equal(writes(s).length, 1);
  }
);
check(
  'Malformed/unsupported lists and false mutation acknowledgements refuse without hiding the unsaved form',
  async (t) => {
    const s = store();
    s.overrides.list = async () => ({ ...list(), filename: '/private/forged-registry.json' });
    await mount(t, s);
    assert.match(document.querySelector('[role=alert]').textContent, /fields/);
    assert.equal(field('Saved actions'), null);
    delete s.overrides.list;
    await click('Reload library');
    await click('New action');
    await change('Action name', 'Still unsaved');
    s.overrides.create = async () => list([], 0);
    await submit();
    assert.equal(field('Action name').value, 'Still unsaved');
    assert.match(document.querySelector('[role=alert]').textContent, /did not advance/);
    assert.equal(writes(s).length, 1);
  }
);
