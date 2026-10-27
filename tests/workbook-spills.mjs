import './support/scratch-dir.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import {
  createWorkbook,
  readWorkbookMetadata,
  indexedCalculationMetadata,
  readWorkbook,
} from '../core/workbook-package.mjs';
import { buildSpillMetadata } from '../core/spill-metadata.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const sheet = 'xl/worksheets/sheet1.xml';
async function fixture(t, name) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-export-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const bytes = await fs.readFile(new URL(`../fixtures/${name}`, import.meta.url));
  await fs.writeFile(path.join(root, 'source.xlsx'), bytes);
  const index = await native({ action: 'workbook', operation: 'index', root });
  const metadata = await fs.readFile(path.join(root, 'metadata.xlsx'));
  const view = readWorkbookMetadata(metadata).view;
  return {
    root,
    bytes,
    index,
    view,
    call: (operation, extra = {}) => native({ action: 'workbook', operation, root, ...extra }),
  };
}
async function save(owner, version) {
  const plan = await owner.call('export-plan', { version, view: owner.view }),
    bytes = await fs.readFile(path.join(owner.root, 'metadata.xlsx')),
    metadata = indexedCalculationMetadata(bytes, plan.incomplete),
    parts = unzipSync(bytes);
  for (const [part, xml] of Object.entries(metadata.replacements)) parts[part] = strToU8(xml);
  for (const part of metadata.removed) delete parts[part];
  metadata.spillMetadata = buildSpillMetadata(zipSync(parts), plan.spills);
  metadata.styleIds = [];
  assert.equal(plan.styleCount, 0);
  await owner.call('artifact-write', { name: 'export-metadata.json', data: JSON.stringify(metadata) });
  await owner.call('export', { version, revision: owner.index.revision, calculationMetadata: true });
  return { plan, metadata, output: path.join(owner.root, 'export.xlsx') };
}

for (const name of ['excel-dynamic-source.xlsx', 'excel-dynamic-reference.xlsx'])
  test(`retained qualified spill metadata and complete typed outputs export/reopen from ${name}`, async (t) => {
    const owner = await fixture(t, name),
      stage = await owner.call('spill-stage', { version: 0, view: owner.view, revision: 'retained-actual-reference' });
    assert.equal(stage.groups, 6);
    assert.equal(stage.outputPositions, 17);
    const { plan, metadata, output } = await save(owner, stage.version);
    assert.equal(plan.incomplete, true);
    assert.equal(plan.spills.length, 6);
    const oracle = JSON.parse(
      await fs.readFile(new URL('../fixtures/excel-dynamic-reference.json', import.meta.url), 'utf8')
    );
    const copy = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-reopen-'));
    t.after(() => fs.rm(copy, { recursive: true, force: true }));
    await fs.copyFile(output, path.join(copy, 'source.xlsx'));
    await native({ action: 'workbook', operation: 'index', root: copy });
    const view = readWorkbookMetadata(await fs.readFile(path.join(copy, 'metadata.xlsx'))).view;
    const again = await native({
      action: 'workbook',
      operation: 'spill-stage',
      root: copy,
      version: 0,
      view,
      revision: 'native-reopen',
    });
    assert.equal(again.groups, 6);
    for (const group of plan.spills) {
      const page = await native({
        action: 'workbook',
        operation: 'range',
        root: copy,
        version: again.version,
        sheet,
        range: group.range,
      });
      for (const [ref, cell] of page.cells) {
        if (group.status === 'error') assert.equal(cell.calculation.error, oracle.cells[ref].expected, ref);
        else assert.equal(cell.calculation.value, oracle.cells[ref].expected, ref);
        assert.equal(cell.spill.master, group.master);
      }
    }
    const expected = {};
    for (const group of plan.spills) {
      const page = await owner.call('range', { version: stage.version, sheet, range: group.range });
      assert.equal(oracle.cells[group.master].formulaAttributes.ref, group.range);
      for (const [ref] of page.cells)
        expected[ref] = {
          value: oracle.cells[ref].expected,
          ...(ref === group.master ? { formula: oracle.authoredFormulas[ref] } : {}),
        };
    }
    const specification = path.join(owner.root, 'spill-export-specification.json'),
      report = fileURLToPath(new URL(`../.build/tests/spill-export/${name}.json`, import.meta.url));
    await fs.writeFile(
      specification,
      JSON.stringify({ masters: metadata.spillMetadata.masters, cells: { [sheet]: expected } })
    );
    const audit = (target) =>
      spawnSync(
        'python3',
        [
          fileURLToPath(new URL('./support/check-spill-export.py', import.meta.url)),
          path.join(owner.root, 'source.xlsx'),
          target,
          specification,
          report,
        ],
        { encoding: 'utf8' }
      );
    const checked = audit(output);
    assert.equal(checked.status, 0, checked.stderr);
    assert.equal(JSON.parse(await fs.readFile(report, 'utf8')).qualifiedCells, 17);
    const acceptedOutput = await fs.readFile(output);
    for (const mutate of [
      (candidate) => {
        candidate.spillMetadata.masters[`${sheet}\0A1`].range = 'A1:C3';
      },
      (candidate) => {
        delete candidate.spillMetadata.masters[`${sheet}\0A1`];
      },
      (candidate) => {
        candidate.spillMetadata.replacements['xl/charts/foreign.xml'] = '<chart/>';
      },
      (candidate) => {
        candidate.spillMetadata.masters[`${sheet}\0A1`].cm = 0;
      },
    ]) {
      const candidate = structuredClone(metadata);
      mutate(candidate);
      await owner.call('artifact-write', { name: 'export-metadata.json', data: JSON.stringify(candidate) });
      await assert.rejects(
        owner.call('export', { version: stage.version, revision: owner.index.revision, calculationMetadata: true })
      );
      assert.deepEqual(await fs.readFile(output), acceptedOutput);
      assert.deepEqual(await fs.readFile(path.join(owner.root, 'source.xlsx')), owner.bytes);
    }
    if (name === 'excel-dynamic-source.xlsx') {
      const savedParts = unzipSync(await fs.readFile(output)),
        text = new TextDecoder().decode(savedParts[sheet]);
      for (const [label, mutate] of [
        [
          'wrong-value',
          (parts) => {
            parts[sheet] = strToU8(
              text.replace(/<c\b[^>]*r="B3"[^>]*>[\s\S]*?<\/c>/, (cell) =>
                cell.replace(/(<v(?:\s[^>]*)?>)[^<]*/, (_full, start) => `${start}999`)
              )
            );
          },
        ],
        [
          'unaddressed-value',
          (parts) => {
            parts[sheet] = strToU8(text.replace('r="F1"', 'r="F999"'));
          },
        ],
        [
          'wrong-geometry',
          (parts) => {
            parts[sheet] = strToU8(text.replace('ref="A1:B3"', 'ref="A1:C3"'));
          },
        ],
        [
          'metadata-index',
          (parts) => {
            parts[sheet] = strToU8(text.replace('cm="1"', 'cm="0"'));
          },
        ],
      ]) {
        const parts = structuredClone(savedParts);
        mutate(parts);
        assert.notDeepEqual(parts[sheet], savedParts[sheet], label);
        const target = path.join(owner.root, `${label}.xlsx`);
        await fs.writeFile(target, zipSync(parts));
        assert.notEqual(audit(target).status, 0, label);
      }
    }
    assert.deepEqual(await fs.readFile(path.join(owner.root, 'source.xlsx')), owner.bytes);
  });

async function authored(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-owner-feature-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const parts = unzipSync(createWorkbook()),
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const literal = (ref, value) => `<c r="${ref}"><v>${value}</v></c>`,
    formula = (ref, value) => `<c r="${ref}"><f>${value}</f><v>777</v></c>`;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1">${formula('A1', 'SEQUENCE($J$1,2)')}${formula('D1', 'SUM(A1:B3)')}${formula('E1', 'IFERROR(A1,99)')}${literal('J1', 3)}</row><row r="2"><c r="B2" s="1"/>${formula('D2', 'B3+1')}</row><row r="3">${formula('D3', 'EndCell+1')}</row><row r="5">${formula('D5', 'SUM(_xlfn.ANCHORARRAY(A1))')}</row></sheetData></worksheet>`
  );
  const workbook = new TextDecoder().decode(parts['xl/workbook.xml']);
  parts['xl/workbook.xml'] = strToU8(
    workbook.replace(
      /(<calcPr\b|<\/workbook>)/,
      '<definedNames><definedName name="EndCell">Sheet1!$B$3</definedName></definedNames>$1'
    )
  );
  const styles = new TextDecoder().decode(parts['xl/styles.xml']);
  parts['xl/styles.xml'] = strToU8(
    styles
      .replace('<cellXfs count="1">', '<cellXfs count="2">')
      .replace(
        '</cellXfs>',
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>'
      )
  );
  const source = path.join(root, 'source.xlsx'),
    bytes = zipSync(parts);
  await fs.writeFile(source, bytes);
  const calls = [],
    transport = async (input) => {
      calls.push(input);
      return native(input);
    },
    make = () =>
      new IndexedWorkbook(
        transport,
        () => assert.fail('No legacy writer'),
        (input) => calculateIndexedWorkbook(input, transport)
      );
  return { root, source, bytes, calls, make };
}
test('indexed qualified SEQUENCE owns complete outputs and recomputes scalar dependents through shrink/growth and reopen', async (t) => {
  const f = await authored(t),
    book = f.make();
  await book.open(f.source);
  t.after(() => book.close());
  let state = '0';
  for (const rows of [3, 2, 3, 2]) {
    if (state !== '0' || rows !== 3) {
      const changed = await book.apply(state, [{ sheet: 'Sheet1', cell: 'J1', value: rows }]);
      state = changed.view.storage.state;
    }
    const cells = (await book.read(state, 'Sheet1', 'A1:J5')).cells;
    assert.equal(cells.A1.spill.role, 'master');
    assert.equal(cells.A1.formula, 'SEQUENCE($J$1,2)');
    assert.equal(cells.A1.calculation.value, 1);
    assert.equal(cells.B2.value, 4);
    assert.equal(cells.B2.spill.role, 'follower');
    assert.equal(cells.B2.styleId, 1);
    assert.equal(cells.B2.horizontalAlign, 'center');
    assert.equal(cells.B3?.value ?? null, rows === 3 ? 6 : null);
    assert.equal(cells.B3?.spill?.master ?? null, rows === 3 ? 'A1' : null);
    assert.equal(cells.D1.calculation.value, rows === 3 ? 21 : 10);
    assert.equal(cells.D2.calculation.value, rows === 3 ? 7 : 1);
    assert.equal(cells.D3.calculation.value, rows === 3 ? 7 : 1);
    assert.equal(cells.E1.calculation.value, 1);
    assert.equal(
      cells.D5.calculation.value,
      rows === 3 ? 21 : 10,
      'Complete owned scalar ANCHORARRAY consumers are now qualified'
    );
    const output = await book.export(state),
      saved = await fs.readFile(output.path);
    assert.equal(readWorkbook(saved).sheets[0].cells.A1.formula, '_xlfn.SEQUENCE($J$1,2)');
    const reopened = f.make();
    await reopened.open(output.path);
    try {
      const after = (await reopened.read('0', 'Sheet1', 'A1:J5')).cells;
      assert.equal(after.D1.calculation.value, cells.D1.calculation.value);
      assert.equal(after.B2.styleId, 1);
      assert.equal(after.B3?.value ?? null, rows === 3 ? 6 : null);
      assert.equal(after.D5.calculation.value, cells.D5.calculation.value);
    } finally {
      await reopened.close();
    }
  }
  const original = (await book.read('0', 'Sheet1', 'A1:J5')).cells;
  assert.equal(original.B3.value, 6);
  assert.equal(original.D1.calculation.value, 21);
  assert.deepEqual(await fs.readFile(f.source), Buffer.from(f.bytes));
});
test('master replacement/release, follower refusal and unsupported argument edits retain old state and replay one user command', async (t) => {
  const f = await authored(t),
    book = f.make();
  await book.open(f.source);
  t.after(() => book.close());
  await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 9 }]), /follower/);
  await assert.rejects(
    book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', formula: 'SEQUENCE(2,2)', expectedFormula: 'wrong' }]),
    /formula/i
  );
  await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 1 }]), { code: 'spill-unsupported' });
  assert.equal((await book.read('0', 'Sheet1', 'J1')).cells.J1.value, 3);
  assert.equal((await book.read('0', 'Sheet1', 'B3')).cells.B3.value, 6);
  const replaced = await book.apply('0', [
    { sheet: 'Sheet1', cell: 'A1', formula: 'SEQUENCE(2,2,10,2)', expectedFormula: 'SEQUENCE($J$1,2)' },
  ]);
  assert.equal(replaced.edits.length, 1);
  let state = replaced.view.storage.state;
  let cells = (await book.read(state, 'Sheet1', 'A1:J5')).cells;
  assert.equal(cells.A1.value, 10);
  assert.equal(cells.B2.value, 16);
  assert.equal(cells.B3.value, null);
  assert.equal(cells.D1.calculation.value, 52);
  const cleared = await book.apply(state, [
    { sheet: 'Sheet1', cell: 'A1', value: null, expectedFormula: 'SEQUENCE(2,2,10,2)' },
  ]);
  assert.equal(cleared.edits.length, 1);
  state = cleared.view.storage.state;
  cells = (await book.read(state, 'Sheet1', 'A1:J5')).cells;
  for (const ref of ['A1', 'B1', 'A2', 'B2']) {
    assert.equal(cells[ref].value, null);
    assert.equal(cells[ref].spill, undefined);
    assert.equal(cells[ref].formula, undefined);
  }
  assert.equal(cells.B2.styleId, 1);
  assert.equal(cells.D1.calculation.value, 0);
  const output = await book.export(state),
    saved = readWorkbook(await fs.readFile(output.path));
  assert.equal(saved.sheets[0].cells.A1?.formula, undefined);
  assert.equal(saved.sheets[0].cells.B2.styleId, 1);
  const replay = f.make();
  await replay.open(f.source);
  try {
    const replayed = await replay.apply('0', [...replaced.edits, ...cleared.edits]);
    assert.equal((await replay.read(replayed.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 0);
    assert.equal((await replay.read(replayed.view.storage.state, 'Sheet1', 'B2')).cells.B2.styleId, 1);
  } finally {
    await replay.close();
  }
  assert.equal((await book.read('0', 'Sheet1', 'B3')).cells.B3.value, 6);
  assert.deepEqual(await fs.readFile(f.source), Buffer.from(f.bytes));
});

test('newly authored qualified master extends public geometry and clear removes complete ownership without other formulas', async (t) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-new-spill-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const source = path.join(root, 'source.xlsx'),
    bytes = createWorkbook();
  await fs.writeFile(source, bytes);
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  await book.open(source);
  t.after(() => book.close());
  const added = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', formula: 'SEQUENCE(3,2)' }]),
    state = added.view.storage.state;
  assert(added.view.sheets[0].rows >= 3);
  assert(added.view.sheets[0].columns >= 2);
  assert.equal((await book.read(state, 'Sheet1', 'B3')).cells.B3.value, 6);
  const output = await book.export(state),
    saved = readWorkbook(await fs.readFile(output.path));
  assert.equal(saved.sheets[0].cells.A1.formula, '_xlfn.SEQUENCE(3,2)');
  assert.equal(saved.sheets[0].cells.B3.value, 6);
  const cleared = await book.apply(state, [{ sheet: 'Sheet1', cell: 'A1', value: null }]),
    empty = await book.export(cleared.view.storage.state),
    again = readWorkbook(await fs.readFile(empty.path));
  for (const ref of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3']) {
    assert.equal(again.sheets[0].cells[ref]?.value ?? null, null);
    assert.equal(again.sheets[0].cells[ref]?.formula, undefined);
  }
  assert.equal((await book.read(state, 'Sheet1', 'B3')).cells.B3.value, 6);
  assert.deepEqual(await fs.readFile(source), Buffer.from(bytes));
});

test('actual literal and zero collision masters carry typed SPILL metadata and recover when the blocker is cleared', async (t) => {
  const f = await fixture(t, 'excel-dynamic-source.xlsx'),
    book = new IndexedWorkbook(
      native,
      () => assert.fail('No legacy writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
  await book.open(path.join(f.root, 'source.xlsx'));
  t.after(() => book.close());
  let state = '0';
  for (const [master, blocker] of [
    ['A41', 'B42'],
    ['A60', 'B61'],
  ]) {
    let cells = (await book.read(state, 'Cases', `${master}:${blocker}`)).cells;
    assert.equal(cells[master].calculation.status, 'error');
    assert.equal(cells[master].calculation.error, '#SPILL!');
    assert.equal(cells[master].rawValue, '#VALUE!');
    assert.equal(cells[blocker].spill, undefined);
    const clear = await book.apply(state, [{ sheet: 'Cases', cell: blocker, value: null }]);
    state = clear.view.storage.state;
    cells = (await book.read(state, 'Cases', `${master}:${blocker}`)).cells;
    assert.equal(cells[master].calculation.value, 1);
    assert.equal(cells[blocker].value, 4);
    assert.equal(cells[blocker].spill.master, master);
    const output = await book.export(state),
      saved = readWorkbook(await fs.readFile(output.path));
    assert.equal(saved.sheets[0].cells[master].value, 1);
    assert.equal(saved.sheets[0].cells[blocker].value, 4);
  }
  const original = (await book.read('0', 'Cases', 'A41:B61')).cells;
  assert.equal(original.A41.calculation.error, '#SPILL!');
  assert.equal(original.B61.value, 0);
  assert.deepEqual(await fs.readFile(path.join(f.root, 'source.xlsx')), f.bytes);
});

test('a calculation fault after private spill staging publishes no group or stale follower and permits retry', async (t) => {
  const f = await authored(t);
  let fail = false;
  const transport = async (input) => {
    const result = await native(input);
    if (fail && input.operation === 'spill-stage') {
      fail = false;
      throw Error('Injected post-stage calculation cancellation');
    }
    return result;
  };
  const book = new IndexedWorkbook(
    transport,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, transport)
  );
  await book.open(f.source);
  t.after(() => book.close());
  fail = true;
  await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]), /post-stage/);
  assert.equal(book.diagnosticStats().retainedStates, 1);
  assert.equal((await book.read('0', 'Sheet1', 'B3')).cells.B3.value, 6);
  assert.equal((await book.read('0', 'Sheet1', 'D1')).cells.D1.calculation.value, 21);
  const retried = await book.apply('0', [{ sheet: 'Sheet1', cell: 'J1', value: 2 }]);
  assert.equal((await book.read(retried.view.storage.state, 'Sheet1', 'B3')).cells.B3.value, null);
  assert.equal((await book.read(retried.view.storage.state, 'Sheet1', 'D1')).cells.D1.calculation.value, 10);
  assert.deepEqual(await fs.readFile(f.source), Buffer.from(f.bytes));
});
