import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook, readWorkbookFilterStates } from '../core/workbook-package.mjs';
import { WorkbookSession } from '../core/workbook-service.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
async function editCalculatedWorkbookResult(bytes, commands) {
  const session = new WorkbookSession(bytes),
    view = await session.read();
  return session.apply({ operationId: 'filter-test', expectedRevision: view.revision, commands });
}
const roots = [],
  clients = [],
  ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function check(name, fn) {
  test(name, async (t) => {
  });
}
after(async () => {
  for (const c of clients) c.close();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
function fixtureBytes({ zero = false, manual = false, unknown = false, error = false } = {}) {
  const parts = unzipSync(createWorkbook());
  const attrs = ['', ' hidden="1"', ' hidden="true"', ' hidden="0"', ' hidden="false"', '', '', '', ''];
  const vals = [
    '<c r="A1" t="inlineStr"><is><t>Key</t></is></c>',
    '<c r="A2"><v>1</v></c>',
    '<c r="A3"><v>1</v></c>',
    '<c r="A4" t="inlineStr"><is><t>1</t></is></c>',
    '<c r="A5" t="b"><v>1</v></c>',
    '<c r="A6"><v>1</v></c>',
    '',
    '<c r="A8" t="inlineStr"><is><t></t></is></c>',
    error ? '<c r="A9" t="e"><v>#N/A</v></c>' : '<c r="A9"><v>2</v></c>',
  ];
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetFormatPr defaultRowHeight="15"${zero ? ' zeroHeight="true"' : ''}/><sheetData>${vals.map((v, i) => `<row r="${i + 1}"${attrs[i]}${i === 8 ? ' ht="23" customHeight="1"' : ''}>${v}<c r="B${i + 1}" t="inlineStr"><is><t>row-${i + 1}</t></is></c></row>`).join('')}</sheetData>${unknown ? '<autoFilter ref="A1:B9"><filterColumn colId="0"><customFilters><customFilter operator="greaterThan" val="0"/></customFilters></filterColumn></autoFilter>' : ''}</worksheet>`
  );
  parts['custom/filter-preserved.bin'] = new Uint8Array([0, 14, 255, 2]);
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</workbook>',
      `<calcPr calcMode="${manual ? 'manual' : 'auto'}" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`
    )
  );
  return Buffer.from(zipSync(parts));
}
const lexical = (bytes) => {
  const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
  return Object.fromEntries(
    [...xml.matchAll(/<row\b([^>]*)\/?\s*>/g)].map((m) => [
      Number(/\br="(\d+)"/.exec(m[1])[1]),
      /\bhidden="([^"]*)"/.exec(m[1])?.[1] ?? null,
    ])
  );
};
async function fixture(bytes) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'filter-reapply-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx'),
    recovery = path.join(root, 'recovery');
  await fs.writeFile(source, bytes);
  return { root, source, recovery };
}
async function owner(f) {
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  clients.push(client);
  const calls = [];
  const native = (input) => {
    calls.push(structuredClone(input));
    return client.call(input);
  };
  const doc = new SpreadsheetDocument({ recoveryDirectory: f.recovery, native });
  await doc.open(f.source);
  return { doc, calls, native };
}
const apply = (doc, command) => doc.request({ operation: 'apply', state: doc.status().state, commands: [command] });
const filter = (doc) => doc.status().view.sheets[0].filter;
const act = { operation: 'filterRows', sheet: 'Sheet1', range: 'A1:B9', column: 'A', values: [1] };
const follow = (doc, operation) => ({ operation, sheet: 'Sheet1', expectedFilterRevision: filter(doc).revision });
async function copy(doc, f, name) {
  const out = path.join(f.root, name + '.xlsx');
  await doc.request({ operation: 'copy', destination: out, expectedRevision: null });
  return fs.readFile(out);
}
const plain = [null, '1', 'true', '0', 'false', null, null, null, null];
function expectClear(bytes) {
  assert.deepEqual(Object.values(lexical(bytes)), plain);
  const v = readWorkbook(bytes);
  assert(v.sheets[0].filter.owned);
  assert.equal(v.sheets[0].filter.canReapply, false);
  assert.equal(v.sheets[0].filter.canClear, false);
  assert.deepEqual(v.sheets[0].filter.values, []);
  assert.equal(v.sheets[0].filterRecord, undefined);
  assert.equal(v.sheets[0].rowHeights.find((r) => r.row === 8).height, 23);
}
check('portable actual-byte service persists criteria, reuses baseline and clears to inactive dropdown', async () => {
  let bytes = fixtureBytes();
  const first = await editCalculatedWorkbookResult(bytes, [act]);
  assert(first.view.sheets[0].filter.owned);
  assert.equal(first.view.sheets[0].filterRecord, undefined);
  assert.equal(readWorkbookFilterStates(first.bytes).Sheet1.model.baseline.rows[1][1], 'true');
  const changed = await editCalculatedWorkbookResult(first.bytes, [{ sheet: 'Sheet1', cell: 'A4', value: 1 }]);
  const reapply = await editCalculatedWorkbookResult(changed.bytes, [
    { operation: 'reapplyFilter', sheet: 'Sheet1', expectedFilterRevision: changed.view.sheets[0].filter.revision },
  ]);
  assert.equal(lexical(reapply.bytes)[4], '0');
  const clear = await editCalculatedWorkbookResult(reapply.bytes, [
    { operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: reapply.view.sheets[0].filter.revision },
  ]);
  expectClear(clear.bytes);
  const parts = unzipSync(fixtureBytes());
  let xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
  for (const r of [6, 7]) xml = xml.replace(new RegExp('<row r="' + r + '"[^>]*>[\\s\\S]*?<\\/row>'), '');
  xml = xml.replace(/<row r="8"[^>]*>[\s\S]*?<\/row>/, '<row r="8"/>');
  parts['xl/worksheets/sheet1.xml'] = strToU8(xml);
  const sparse = await editCalculatedWorkbookResult(zipSync(parts), [act]);
  assert.equal(lexical(sparse.bytes)[6], '1');
  assert.equal(lexical(sparse.bytes)[7], '1');
  const newCell = await editCalculatedWorkbookResult(sparse.bytes, [{ sheet: 'Sheet1', cell: 'A6', value: 'new' }]);
  const restored = await editCalculatedWorkbookResult(newCell.bytes, [
    { operation: 'clearFilter', sheet: 'Sheet1', expectedFilterRevision: newCell.view.sheets[0].filter.revision },
  ]);
  assert.equal(lexical(restored.bytes)[7], undefined);
  assert.equal(lexical(restored.bytes)[8], null);
  assert.equal(readWorkbook(restored.bytes).sheets[0].cells.A6.value, 'new');
  const zeroParts = unzipSync(fixtureBytes());
  zeroParts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(zeroParts['xl/worksheets/sheet1.xml'])
      .replace('<c r="A4" t="inlineStr"><is><t>1</t></is></c>', '<c r="A4"><v>0</v></c>')
      .replace('<c r="A5" t="b"><v>1</v></c>', '<c r="A5"><v>-0</v></c>')
  );
  const session = new WorkbookSession(zipSync(zeroParts)),
    base = await session.read(),
    request = {
      operationId: 'signed-zero-filter',
      expectedRevision: base.revision,
      commands: [{ ...act, values: [-0] }],
    },
    signed = await session.apply(request);
  assert(Object.is(signed.view.sheets[0].filter.values[0], -0));
  assert.equal(lexical(signed.bytes)[4], '1');
  assert.equal(lexical(signed.bytes)[5], 'false');
  await assert.rejects(session.apply({ ...request, commands: [{ ...act, values: [0] }] }), {
    code: /^(internal-invalid-field|workbook-duplicate)$/,
  });
  assert.deepEqual(session.bytes, signed.bytes);
  return {
    baselineBackendOnly: true,
    clearRetainsDropdown: true,
    sparseRowsRestoredWithoutDeletingEditedCells: true,
    typedSignedZeroReceipt: true,
  };
});
check(
  'actual Document edits reapply accepted criteria and exact baseline survives Copy, UndoRedo, Save and cold',
  async () => {
    for (const zero of [false, true]) {
      const original = fixtureBytes({ zero }),
        f = await fixture(original),
        { doc } = await owner(f);
      try {
        await apply(doc, act);
        const first = filter(doc);
        assert(first.owned && first.canReapply && first.canClear);
        const firstBytes = await copy(doc, f, 'first');
        assert.equal(lexical(firstBytes)[2], '1');
        assert.equal(lexical(firstBytes)[3], 'true');
        assert.equal(lexical(firstBytes)[4], '1');
        assert.equal(lexical(firstBytes)[5], '1');
        assert.equal(lexical(firstBytes)[6], null);
        await apply(doc, { sheet: 'Sheet1', cell: 'A4', value: 1 });
        const edited = doc.status().state;
        await apply(doc, follow(doc, 'reapplyFilter'));
        const reapplied = doc.status().state;
        assert.notEqual(reapplied, edited);
        assert.equal(filter(doc).revision, first.revision);
        assert.equal(lexical(await copy(doc, f, 'reapplied'))[4], '0');
        await doc.request({ operation: 'undo' });
        assert.equal(doc.status().state, edited);
        await doc.request({ operation: 'redo' });
        assert.equal(doc.status().state, reapplied);
        await apply(doc, follow(doc, 'clearFilter'));
        const cleared = await copy(doc, f, 'cleared');
        expectClear(cleared);
        assert.deepEqual(await fs.readFile(f.source), original);
        await doc.request({ operation: 'save' });
        await doc.close();
        const cold = await owner(f);
        try {
          assert.equal(cold.doc.status().dirty, false);
          assert.equal(filter(cold.doc).canReapply, false);
          expectClear(await copy(cold.doc, f, 'cold'));
          await apply(cold.doc, { ...act, values: [true, '1', null, ''] });
          assert.deepEqual(filter(cold.doc).values, [true, '1', null, '']);
          await apply(cold.doc, follow(cold.doc, 'clearFilter'));
          expectClear(await copy(cold.doc, f, 'replacement-clear'));
        } finally {
          await cold.doc.close();
        }
      } finally {
        await doc.close();
      }
    }
    const portable = await editCalculatedWorkbookResult(fixtureBytes(), [act]),
      portableFixture = await fixture(portable.bytes),
      imported = await owner(portableFixture);
    try {
      assert(filter(imported.doc).owned);
      await apply(imported.doc, follow(imported.doc, 'reapplyFilter'));
      await apply(imported.doc, follow(imported.doc, 'clearFilter'));
      expectClear(await copy(imported.doc, portableFixture, 'portable-native-clear'));
    } finally {
      await imported.doc.close();
    }
    return {
      defaultVisibilityProfiles: 2,
      oneUndoRedo: true,
      coldOwnedDefinition: true,
      portableToNativeProvenance: true,
    };
  }
);
check('stale/extra commands and forged native view/version preserve accepted head and source', async () => {
  const bytes = fixtureBytes(),
    f = await fixture(bytes),
    { doc, calls, native } = await owner(f);
  try {
    await apply(doc, act);
    const before = doc.status(),
      copyBefore = await copy(doc, f, 'negative-before');
    for (const bad of [
      { operation: 'reapplyFilter', sheet: 'Sheet1', expectedFilterRevision: '0'.repeat(64) },
      { ...follow(doc, 'clearFilter'), range: 'A1:B9' },
      { ...follow(doc, 'reapplyFilter'), values: [true] },
    ]) {
      await assert.rejects(apply(doc, bad));
      assert.equal(doc.status().state, before.state);
      assert.deepEqual(await fs.readFile(f.source), bytes);
    }
    const request = calls.findLast((c) => c.operation === 'transform');
    await assert.rejects(
      native({
        ...request,
        view: { ...request.view, sheets: request.view.sheets.map((s) => ({ ...s, name: 'forged' })) },
      }),
      /metadata|No such file/
    );
    await assert.rejects(
      native({ ...request, version: 999999 }),
      /version|preparation|No such file|writer input differs/
    );
    assert.deepEqual(await copy(doc, f, 'negative-after'), copyBefore);
    return { closedRefusals: 3, retiredNativeReplayWithAlteredViewAndVersionRefused: true };
  } finally {
    await doc.close();
  }
});
check('unknown imported criteria and literal errors refuse without accepted filter provenance', async () => {
  for (const variant of [{ unknown: true }, { error: true }, { mode: true }, { extension: true }]) {
    let bytes = fixtureBytes(variant);
    if (variant.mode || variant.extension) {
      const p = unzipSync(bytes);
      let xml = strFromU8(p['xl/worksheets/sheet1.xml']);
      xml = variant.mode
        ? xml.replace('<sheetFormatPr', '<sheetPr filterMode="true"/><sheetFormatPr')
        : xml.replace(
            '</worksheet>',
            '<autoFilter ref="A1:B9"><extLst><ext uri="foreign"/></extLst></autoFilter></worksheet>'
          );
      p['xl/worksheets/sheet1.xml'] = strToU8(xml);
      bytes = Buffer.from(zipSync(p));
    }
    const f = await fixture(bytes),
      { doc } = await owner(f);
    try {
      const state = doc.status().state;
      if (variant.unknown) {
        assert.equal(filter(doc).owned, false);
        assert.equal(filter(doc).canReapply, false);
      }
      await assert.rejects(apply(doc, act), /Imported|Unknown|typed errors|provenance/i);
      assert.equal(doc.status().state, state);
      assert.deepEqual(await fs.readFile(f.source), bytes);
    } finally {
      await doc.close();
    }
  }
  return { importedCriteriaProtected: true, typedErrorNotText: true };
});
check(
  'Manual pending key refuses Reapply while Clear preserves pending formula and stored-source honesty',
  async () => {
    const bytes = fixtureBytes({ manual: true }),
      f = await fixture(bytes),
      { doc, calls } = await owner(f);
    try {
      await apply(doc, act);
      await apply(doc, { sheet: 'Sheet1', cell: 'A6', formula: '1' });
      const before = doc.status().state;
      calls.length = 0;
      await assert.rejects(apply(doc, follow(doc, 'reapplyFilter')), /calculat|pending|resolved/i);
      assert.equal(doc.status().state, before);
      await apply(doc, follow(doc, 'clearFilter'));
      assert(!calls.some((c) => c.operation === 'calculate'));
      const cell = (await doc.request({ operation: 'read', state: doc.status().state, sheet: 'Sheet1', range: 'A6' }))
        .cells.A6;
      assert.equal(cell.formula, '1');
      assert.equal(cell.calculation.pending, true);
      assert.equal(Object.hasOwn(cell.calculation, 'value'), false);
      expectClear(await copy(doc, f, 'manual-clear'));
      assert.deepEqual(await fs.readFile(f.source), bytes);
      return { noImplicitCalculation: true, pendingClearHonest: true };
    } finally {
      await doc.close();
    }
  }
);
check('header is not a criterion: unsafe numeric and Manual pending header preserve exact source', async () => {
  for (const manual of [false, true]) {
    const p = unzipSync(fixtureBytes({ manual }));
    p['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(p['xl/worksheets/sheet1.xml']).replace(
        '<c r="A1" t="inlineStr"><is><t>Key</t></is></c>',
        manual ? '<c r="A1"><f>1</f><v>913</v></c>' : '<c r="A1"><v>9007199254740993</v></c>'
      )
    );
    const bytes = Buffer.from(zipSync(p)),
      f = await fixture(bytes),
      { doc, calls } = await owner(f);
    try {
      calls.length = 0;
      await apply(doc, act);
      await apply(doc, follow(doc, 'reapplyFilter'));
      const saved = await copy(doc, f, 'header');
      const xml = strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml']);
      assert(xml.includes(manual ? '<f>1</f><v>913</v>' : '<v>9007199254740993</v>'));
      if (manual) assert(!calls.some((c) => c.operation === 'calculate'));
      assert.deepEqual(await fs.readFile(f.source), bytes);
    } finally {
      await doc.close();
    }
  }
  return { excludedHeaderProfiles: 2 };
});
