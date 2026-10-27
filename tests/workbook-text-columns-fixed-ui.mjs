import './support/scratch-dir.mjs';
/** Fixed-width TTC source scaffolding, not run until implementation release.
 * Actual future model/pane imports; controlled JS reads/acceptance only. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React, { act } from 'react';
import { JSDOM } from 'jsdom';
import { fixedWidthFixture as f } from './support/text-columns-fixed-width-fixture.mjs';
import { rangeBounds } from '../core/workbook-commands.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/fixed-ttc-ui-')),
  checks = [];
const entry = path.join(output, 'entry.ts'),
  bundle = path.join(output, 'test.mjs');
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/textColumns.ts'))};export {WorkbookTextColumns} from ${JSON.stringify(path.join(component, 'ui/WorkbookTextColumns.tsx'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  format: 'esm',
  platform: 'browser',
  jsx: 'automatic',
  outfile: bundle,
  external: ['react', 'react/jsx-runtime', 'react-dom', 'react-dom/client'],
  logLevel: 'silent',
});
const dom = new JSDOM('<div id="root"></div>', { url: 'http://localhost' });
Object.assign(globalThis, {
  window: dom.window,
  document: dom.window.document,
  HTMLElement: dom.window.HTMLElement,
  Node: dom.window.Node,
  IS_REACT_ACT_ENVIRONMENT: true,
});
const { createRoot } = await import('react-dom/client'),
  ui = await import(pathToFileURL(bundle));
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
after(async () => {
  dom.window.close();
});
const fixed = (breaks = f.breaks, policies) => ({
    mode: 'fixed-width',
    breaks: [...breaks],
    destination: 'B3',
    ...(policies ? { policies: [...policies] } : {}),
  }),
  cell = (value) => ({ value, dataType: 'text' }),
  binding = () => ({
    documentId: 'doc',
    grantDocumentId: 'grant',
    ownerIdentity: 'owner',
    ownerKey: 'owner:head',
    state: 'head',
    sheet: 'S',
    selection: 'A1:A2',
    navigationGeneration: 0,
    activationEpoch: 0,
  });
function fixture({ records = f.records, date1904 = false, selection = 'A1:A2', accept } = {}) {
  let current = true;
  const cells = Object.fromEntries(records.map((text, i) => ['A' + (i + 1), cell(text)])),
    accepted = [],
    context = {
      binding: { ...binding(), selection },
      date1904,
      isCurrent: () => current,
      isRuntimeCurrent: () => current,
      availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
      read: async (range) => {
        const b = rangeBounds(range);
        return Object.fromEntries(
          Object.entries(cells).filter(([ref]) => {
            const p = rangeBounds(ref).first;
            return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
          })
        );
      },
      inspect: async () => ({ reason: null }),
      accept: accept ?? (async (commands) => accepted.push(commands)),
    };
  return {
    context,
    cells,
    accepted,
    retire: () => {
      current = false;
    },
  };
}
const props = (node) => node[Object.keys(node).find((key) => key.startsWith('__reactProps'))],
  field = (name) => [...document.querySelectorAll('[aria-label]')].find((n) => n.getAttribute('aria-label') === name),
  button = (name) => [...document.querySelectorAll('button')].find((n) => n.textContent === name),
  flush = async () => act(async () => new Promise((r) => setTimeout(r, 4)));
async function change(node, value) {
  assert(node);
  await act(async () => props(node).onChange({ target: { value } }));
  await flush();
}
async function click(name) {
  const node = button(name);
  assert(node);
  assert.equal(node.disabled, false);
  await act(async () => node.click());
  await flush();
}
async function mount(t, context) {
  const root = createRoot(document.getElementById('root'));
  let p = { context, visible: true };
  await act(async () => root.render(React.createElement(ui.WorkbookTextColumns, p)));
  await flush();
  t.after(async () => act(async () => root.unmount()));
  return {
    render: async (next) => {
      p = { ...p, ...next };
      await act(async () => root.render(React.createElement(ui.WorkbookTextColumns, p)));
      await flush();
    },
  };
}
check(
  'practical fixed widths preserve every raw space, leading zero and formula-looking string under default Text',
  async () => {
    const h = fixture(),
      p = await ui.previewTextColumns(h.context, fixed());
    assert.deepEqual(
      p.rows.map((row) => [...row.fields]),
      f.fields.map((row) => [...row])
    );
    assert.equal(p.outputs.length, 10);
    assert(p.outputs.every((o) => o.conversion.kind === 'text'));
    assert.deepEqual(
      p.commands[0].cells.map((c) => c.value),
      f.fields.flat()
    );
    assert.equal(p.commands[0].cells[4].formula, undefined);
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'breaks count Unicode code points without splitting emoji; combining marks remain a separately counted character',
  async () => {
    const h = fixture({ records: [f.unicode.record], selection: 'A1' }),
      p = await ui.previewTextColumns(h.context, fixed(f.unicode.breaks));
    assert.deepEqual([...p.rows[0].fields], [...f.unicode.fields]);
    assert.deepEqual(
      p.outputs.map((o) => o.value),
      [...f.unicode.fields]
    );
    assert.equal(p.rows[0].fields[1], '😀');
    assert.equal(p.rows[0].fields[3], '\u0301');
  }
);
check(
  'break exactly at end yields ordinary empty text; starts beyond end are explicitly padded, without invented spaces or zeros',
  async () => {
    const h = fixture({ records: [f.short.record, f.empty.record] }),
      p = await ui.previewTextColumns(h.context, fixed(f.short.breaks, ['text', 'number', 'boolean']));
    assert.deepEqual(
      p.rows.map((row) => [...row.fields]),
      [[...f.short.fields], [...f.empty.fields]]
    );
    assert.deepEqual(
      p.outputs.map((o) => o.padded),
      [...f.short.padded, ...f.empty.padded]
    );
    assert(p.outputs.filter((o) => o.original === '').every((o) => o.conversion.kind === 'text' && o.value === ''));
    assert.equal(p.rejections, 0);
  }
);
check(
  'same explicit per-field policies preserve sign and both calendar serials in one guarded batch; non-Text changes remain disclosed',
  async () => {
    for (const date1904 of [false, true]) {
      const h = fixture({ date1904 }),
        p = await ui.previewTextColumns(h.context, fixed(f.breaks, f.policies));
      assert.deepEqual(
        p.outputs.map((o) => o.value),
        [
          '00123',
          12.5,
          true,
          f.calendars[date1904 ? 1904 : 1900],
          '=SUM(A1)',
          'abcde',
          -0,
          false,
          f.calendars[date1904 ? 1904 : 1900],
          '=1+2',
        ]
      );
      assert(Object.is(p.commands[0].cells[6].value, -0));
      assert.equal(p.commands[0].cells[3].numberFormat, 'date');
      assert.equal(p.outputs[1].conversion.notices[0], 'Outer whitespace removed.');
      await ui.applyTextColumns(h.context, p);
      assert.equal(h.accepted.length, 1);
      assert.equal(h.accepted[0].length, 1);
    }
  }
);
check(
  'invalid, duplicated, descending or mixed-mode breaks and total position bounds refuse without owner writes',
  async () => {
    const h = fixture();
    for (const options of [
      fixed([]),
      fixed([0]),
      fixed([1, 1]),
      fixed([3, 2]),
      fixed([1.5]),
      fixed([32768]),
      { ...fixed(), delimiter: ',' },
      { ...fixed(), quotes: 'none' },
      { ...fixed(), breaks: Array(2) },
    ])
      await assert.rejects(ui.previewTextColumns(h.context, options));
    const bounded = fixture({ records: Array.from({ length: 1000 }, () => f.records[0]), selection: 'A1:A1000' });
    await assert.rejects(
      ui.previewTextColumns(bounded.context, fixed(Array.from({ length: 10 }, (_, i) => i + 1))),
      /10,000|positions/
    );
    assert.deepEqual(h.accepted, []);
    assert.deepEqual(bounded.accepted, []);
  }
);
check(
  'fixed-width collisions require exact replacement acknowledgment; source edits and retired owner cannot reuse accepted preview',
  async () => {
    const h = fixture();
    h.cells.B3 = { value: 'Keep', dataType: 'text', rawValue: 'Keep' };
    const p = await ui.previewTextColumns(h.context, fixed());
    assert.equal(p.collisions, 1);
    await assert.rejects(ui.applyTextColumns(h.context, p), /occupied/);
    assert.deepEqual(h.accepted, []);
    h.cells.A1 = cell('Changed' + f.records[0]);
    await assert.rejects(ui.applyTextColumns(h.context, p, { replaceFingerprint: p.fingerprint }), /changed/);
    h.retire();
    await assert.rejects(
      ui.applyTextColumns(h.context, p, { replaceFingerprint: p.fingerprint }),
      ui.TextColumnsRetired
    );
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'actual split-method controls retire previews and retain typed controls; completed Cancel returns default Text and zero history',
  async (t) => {
    const h = fixture();
    await mount(t, h.context);
    await change(field('Split method'), 'fixed-width');
    await change(field('Split after characters'), '5,13,18,28');
    await change(field('Destination start'), 'B3');
    await click('Preview split');
    assert.equal(field('Delimiter'), undefined);
    assert.equal(field('Quote policy'), undefined);
    assert.equal(field('Output value B3').textContent, '"00123"');
    await change(field('Field 2 type'), 'number');
    assert.equal(field('Split preview'), undefined);
    assert(field('Field 3 type'));
    await change(field('Field 3 type'), 'boolean');
    await change(field('Field 4 type'), 'iso-date');
    await click('Preview split');
    assert.equal(field('Output value C3').textContent, '12.5');
    assert.equal(field('Output value E3').textContent, '40729');
    await click('Cancel split');
    assert.equal(field('Split preview'), undefined);
    assert.equal(field('Field 2 type').value, 'text');
    assert.deepEqual(h.accepted, []);
  }
);
check(
  'changing boundaries invalidates review and field map, leaving source uncommitted; hidden/head-retired pane cannot apply',
  async (t) => {
    const h = fixture(),
      m = await mount(t, h.context);
    await change(field('Split method'), 'fixed-width');
    await change(field('Split after characters'), '5,13,18,28');
    await change(field('Destination start'), 'B3');
    await click('Preview split');
    await change(field('Split after characters'), '5,18,28');
    assert.equal(field('Split preview'), undefined);
    assert.deepEqual(h.accepted, []);
    await click('Preview split');
    h.retire();
    await m.render({ context: h.context });
    assert.equal(field('Split preview'), undefined);
    assert.match(document.body.textContent, /changed. Reopen/);
    await m.render({ visible: false });
    assert.deepEqual(h.accepted, []);
  }
);
