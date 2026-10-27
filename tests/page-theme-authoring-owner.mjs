/** Actual owner theme writes/history/calculation, not gallery qualification. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { pageThemeSamples, pageThemeXml } from '../core/page-style-samples.mjs';
import { pageThemeCommand } from '../core/page-theme-command.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
const roots = [],
  hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
after(async () => {
  await closeNativeWorkbookWorker();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
function source({ foreign = false, protectedSheet = false } = {}) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Body</t></is></c><c r="B1"><f>SUM(D1,1)</f><v>777</v></c><c r="C1" s="1" t="inlineStr"><is><t>Literal</t></is></c><c r="D1"><v>5</v></c></row></sheetData>' +
      (protectedSheet ? '<sheetProtection sheet="1"/>' : '') +
      '</worksheet>'
  );
  let styles = strFromU8(parts['xl/styles.xml']);
  styles = styles
    .replace('fonts count="1"', 'fonts count="2"')
    .replace('</fonts>', '<font><name val="Courier New"/><sz val="12"/><color rgb="FF11AA44"/></font></fonts>')
    .replace('cellXfs count="1"', 'cellXfs count="2"')
    .replace('</cellXfs>', '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs>');
  parts['xl/styles.xml'] = strToU8(styles);
  parts['custom/opaque.bin'] = new Uint8Array([1, 3, 5, 255]);
  if (foreign) {
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
        '</Relationships>',
        '<Relationship Id="importedTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/foreign.xml"/></Relationships>'
      )
    );
    parts['xl/theme/foreign.xml'] = strToU8(pageThemeXml('plum'));
  }
  return { parts, bytes: zipSync(parts) };
}
async function fixture(t, options) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'theme-authoring-'));
  roots.push(root);
  const input = source(options),
    filename = path.join(root, 'source.xlsx'),
    recoveryDirectory = path.join(root, 'recovery');
  await fs.writeFile(filename, input.bytes);
  const owners = [];
  const open = async (file = filename) => {
    const owner = new SpreadsheetDocument({ recoveryDirectory: path.join(recoveryDirectory, String(owners.length)) });
    owners.push(owner);
    await owner.open(file, { fresh: true });
    return owner;
  };
  const recover = async (id) => {
    const owner = new SpreadsheetDocument({ recoveryDirectory: path.join(recoveryDirectory, '0') });
    owners.push(owner);
    await owner.recover(id);
    return owner;
  };
  t.after(async () => {
    for (const owner of owners) await owner.close();
  });
  return { root, filename, ...input, open, recover };
}
const request = (owner, operation, rest = {}) => owner.request({ operation, state: owner.status().state, ...rest });
const apply = (owner, commands) =>
  request(owner, 'apply', { commands: Array.isArray(commands) ? commands : [commands] });
const read = async (owner) => (await request(owner, 'read', { sheet: 'Sheet1', range: 'A1:D1' })).cells;
const command = (owner, id) => ({
  type: 'applyPageTheme',
  sample: id,
  expectedThemeRevision: owner.status().view.pageTheme.revision,
});
const calculation = (cell) =>
  Object.fromEntries(Object.entries(cell.calculation ?? {}).filter(([key]) => key !== 'revision'));

test('closed commands and native catalog assets match the real three samples', async () => {
  for (const sample of pageThemeSamples) {
    assert.equal(
      await fs.readFile(new URL('../native/assets/page-themes/' + sample.id + '.xml', import.meta.url), 'utf8'),
      pageThemeXml(sample.id)
    );
    assert.equal(
      pageThemeCommand({ type: 'applyPageTheme', sample: sample.id, expectedThemeRevision: null }).sample,
      sample.id
    );
  }
  for (const value of [
    { type: 'applyPageTheme', sample: 'Office', expectedThemeRevision: null },
    { type: 'applyPageTheme', sample: 'forest', expectedThemeRevision: null, xml: 'caller' },
    { type: 'applyPageTheme', sample: 'forest', expectedThemeRevision: 'bad' },
    { type: 'applyPageTheme', sample: 'forest' },
  ])
    assert.throws(() => pageThemeCommand(value));
});

test('all three actual theme writes resolve Normal cells, preserve literals, Undo/Redo and cold saved output', async (t) => {
  for (const sample of pageThemeSamples) {
    const f = await fixture(t),
      owner = await f.open(),
      before = await read(owner),
      start = owner.status().state;
    await apply(owner, command(owner, sample.id));
    const accepted = owner.status().state;
    assert.notEqual(accepted, start);
    assert.equal(owner.status().view.pageTheme.sample, sample.id);
    assert.equal(owner.status().view.pageTheme.canApply, true);
    assert.equal((await read(owner)).A1.fontFamily, sample.minorFont);
    assert.equal((await read(owner)).A1.fontColor, sample.colors.dk1);
    assert.equal((await read(owner)).C1.fontFamily, 'Courier New');
    assert.equal((await read(owner)).C1.fontColor, '11AA44');
    assert.deepEqual(calculation((await read(owner)).B1), calculation(before.B1));
    await request(owner, 'undo');
    assert.equal(owner.status().state, start);
    assert.equal((await read(owner)).A1.fontFamily, 'Calibri');
    await request(owner, 'redo');
    assert.equal(owner.status().state, accepted);
    const saved = path.join(f.root, sample.id + '-saved.xlsx');
    await request(owner, 'copy', { destination: saved, expectedRevision: null });
    const parts = unzipSync(await fs.readFile(saved));
    assert.equal(strFromU8(parts['xl/theme/workbook-page-theme.xml']), pageThemeXml(sample.id));
    assert.deepEqual(parts['custom/opaque.bin'], f.parts['custom/opaque.bin']);
    assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
    const independent = readWorkbook(await fs.readFile(saved));
    assert.equal(independent.sheets[0].cells.A1.fontFamily, sample.minorFont);
    assert.equal(independent.sheets[0].cells.C1.fontFamily, 'Courier New');
    const recoveryId = owner.status().documentId;
    await owner.close();
    const recovered = await f.recover(recoveryId);
    assert.equal(recovered.status().view.pageTheme.sample, sample.id);
    await request(recovered, 'undo');
    assert.equal((await read(recovered)).A1.fontFamily, 'Calibri');
    await request(recovered, 'redo');
    assert.equal((await read(recovered)).A1.fontFamily, sample.minorFont);
    await recovered.close();
    const cold = await f.open(saved);
    assert.equal((await read(cold)).A1.fontFamily, sample.minorFont);
    assert.equal(cold.status().view.pageTheme.sample, sample.id);
    await cold.close();
  }
});

test('same-theme no-op retains head/Redo/recovery and stale/late mixed failures roll back', async (t) => {
  const f = await fixture(t),
    owner = await f.open();
  await apply(owner, command(owner, 'forest'));
  const first = owner.status().state;
  await apply(owner, command(owner, 'plum'));
  await request(owner, 'undo');
  assert.equal(owner.status().state, first);
  const before = structuredClone(owner.status()),
    cells = await read(owner),
    manifestPath = path.join(f.root, 'recovery/0', 'draft-' + before.documentId, 'current.json'),
    manifest = await fs.readFile(manifestPath);
  await apply(owner, command(owner, 'forest'));
  assert.deepEqual(owner.status(), before);
  assert.deepEqual(await read(owner), cells);
  assert.deepEqual(await fs.readFile(manifestPath), manifest);
  for (const value of [
    { type: 'applyPageTheme', sample: 'neutral', expectedThemeRevision: null },
    { ...command(owner, 'neutral'), xml: 'caller XML' },
  ])
    await assert.rejects(apply(owner, value));
  await assert.rejects(
    apply(owner, [command(owner, 'neutral'), { sheet: 'missing', cell: 'A1', value: 'late failure' }])
  );
  assert.deepEqual(owner.status(), before);
  assert.deepEqual(await read(owner), cells);
});

test('Manual pending formula state survives a theme change and remains pending on saved reopen', async (t) => {
  const f = await fixture(t),
    owner = await f.open();
  await apply(owner, { type: 'setCalculationOptions', mode: 'manual', recalculateBeforeSave: false });
  await apply(owner, { sheet: 'Sheet1', cell: 'D1', value: 8 });
  const before = await read(owner),
    policy = structuredClone(owner.status().view.calculationPolicy);
  assert(before.B1.calculation.pending);
  await apply(owner, command(owner, 'forest'));
  const after = await read(owner);
  assert.deepEqual(calculation(after.B1), calculation(before.B1));
  for (const key of ['mode', 'recalculateBeforeSave', 'pendingFormulaCells', 'pendingCacheObjects'])
    assert.equal(owner.status().view.calculationPolicy[key], policy[key]);
  const saved = path.join(f.root, 'manual-saved.xlsx');
  await request(owner, 'copy', { destination: saved, expectedRevision: null });
  await owner.close();
  const cold = await f.open(saved);
  assert((await read(cold)).B1.calculation.pending);
  assert.equal((await read(cold)).A1.fontFamily, 'Verdana');
});

test('imported theme and readonly worksheet sources stay unchanged after authoring refusal', async (t) => {
  for (const options of [{ foreign: true }, { protectedSheet: true }]) {
    const f = await fixture(t, options),
      owner = await f.open(),
      before = structuredClone(owner.status()),
      cells = await read(owner);
    await assert.rejects(
      apply(owner, { type: 'applyPageTheme', sample: 'forest', expectedThemeRevision: before.view.pageTheme.revision })
    );
    assert.deepEqual(owner.status(), before);
    assert.deepEqual(await read(owner), cells);
    assert.equal(hash(await fs.readFile(f.filename)), hash(f.bytes));
    await owner.close();
  }
});
