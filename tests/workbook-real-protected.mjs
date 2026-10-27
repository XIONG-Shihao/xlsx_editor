import './support/scratch-dir.mjs';
import { assertRetainedPart } from '../scripts/test/package-oracle.mjs';
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const root = fileURLToPath(new URL('../', import.meta.url)),
  // The supplied original is not in Git; XLSX_EDITOR_FINANCIAL_WORKBOOK names a local copy.
  source = process.env.XLSX_EDITOR_FINANCIAL_WORKBOOK ?? '',
  fixture = path.join(root, 'fixtures/financial-project-cost-opaque-filter.xlsx'),
  manifest = path.join(root, 'fixtures/financial-project-cost-opaque-filter.json'),
  sheet = '项目成本',
  part = 'xl/worksheets/sheet10.xml';
const sha = (value) => createHash('sha256').update(value).digest('hex');
const book = () =>
  new IndexedWorkbook(
    native,
    () => assert.fail('Intentional protection cannot invoke a legacy writer'),
    () => assert.fail('No formula calculation is needed for this minimized literal case')
  );

// Attribution to the supplied original needs that local workbook; the minimized
// fixture below is self-contained and always verifies the protection behavior.
test(
  'real financial source reduction is byte reproducible and independently attributed to original cost/filter/style XML',
  { skip: existsSync(source) ? false : 'Optional supplied original: set XLSX_EDITOR_FINANCIAL_WORKBOOK' },
  async (t) => {
    const record = JSON.parse(await fs.readFile(manifest, 'utf8'));
    assert.equal(sha(await fs.readFile(source)), record.source.sha256);
    const audit = spawnSync(
      'python3',
      [path.join(root, 'tests/support/check-minimized-financial.py'), source, fixture, manifest],
      { encoding: 'utf8' }
    );
    assert.equal(audit.status, 0, audit.stderr);
    const proof = JSON.parse(audit.stdout);
    assert.equal(proof.retainedCells, 1);
    assert.equal(proof.externalRuntimeClaim, false);
    const scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-real-reduction-'));
    t.after(() => fs.rm(scratch, { recursive: true, force: true }));
    const regenerated = path.join(scratch, path.basename(fixture)),
      regeneratedManifest = path.join(scratch, path.basename(manifest));
    const generated = spawnSync(
      'python3',
      [path.join(root, 'tests/support/minimize-financial-filter.py'), source, regenerated, regeneratedManifest],
      { encoding: 'utf8' }
    );
    assert.equal(generated.status, 0, generated.stderr);
    assert.deepEqual(await fs.readFile(regenerated), await fs.readFile(fixture));
    assert.deepEqual(JSON.parse(await fs.readFile(regeneratedManifest, 'utf8')), record);
    const realRoot = path.join(scratch, 'original-index');
    await fs.mkdir(realRoot);
    await fs.copyFile(source, path.join(realRoot, 'source.xlsx'));
    const indexed = await native({ action: 'workbook', operation: 'index', root: realRoot });
    assert.equal(indexed.sheets.find((item) => item.path === part).extended, true);
    const value = await native({
      action: 'workbook',
      operation: 'range',
      root: realRoot,
      version: 0,
      sheet: part,
      range: 'D4',
    });
    assert.equal(value.cells[0][1].rawValue, '254063.62');
    assert.equal(value.cells[0][1].styleId, 127);
    assert.equal(sha(await fs.readFile(source)), record.source.sha256);
    await fs.mkdir(path.join(root, '.build/tests/real-protected'), { recursive: true });
    await fs.writeFile(
      path.join(root, '.build/tests/real-protected/attribution.json'),
      JSON.stringify({ ...proof, regeneratedByteIdentical: true, originalIndexedSourceFeature: true }, null, 2) + '\n'
    );
  }
);

test('minimized real opaque filter refuses structural change atomically while removing only that feature qualifies the control', async (t) => {
  const original = await fs.readFile(fixture),
    session = book();
  await session.open(fixture);
  t.after(() => session.close());
  const view = session.view('0'),
    before = (await session.read('0', sheet, 'D4')).cells.D4;
  assert.equal(view.sheets[0].capabilities.insertRows, false);
  assert.equal(before.rawValue, '254063.62');
  assert.equal(before.styleId, 127);
  await assert.rejects(
    session.apply('0', [{ operation: 'insertRow', sheet, index: 3, count: 1 }]),
    /extended|qualification|qualified/i
  );
  assert.equal(session.diagnosticStats().retainedStates, 1);
  assert.deepEqual((await session.read('0', sheet, 'D4')).cells.D4, before);
  const refused = await session.export('0');
  assert.deepEqual(await fs.readFile(refused.path), original);
  assert.deepEqual(await fs.readFile(fixture), original);
  const scratch = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-real-protected-control-'));
  t.after(() => fs.rm(scratch, { recursive: true, force: true }));
  const parts = unzipSync(original),
    xml = strFromU8(parts[part]),
    filter = xml.match(/<autoFilter\b[\s\S]*?<\/autoFilter>/)?.[0];
  assert(filter);
  parts[part] = strToU8(xml.replace(filter, ''));
  const control = path.join(scratch, 'control.xlsx');
  await fs.writeFile(control, zipSync(parts));
  const plain = book();
  await plain.open(control);
  t.after(() => plain.close());
  assert.equal(plain.view('0').sheets[0].capabilities.insertRows, true);
  const changed = await plain.apply('0', [{ operation: 'insertRow', sheet, index: 3, count: 1 }]),
    cell = (await plain.read(changed.view.storage.state, sheet, 'D5')).cells.D5;
  assert.equal(cell.rawValue, before.rawValue);
  assert.equal(cell.styleId, 127);
  assert.equal((await plain.read(changed.view.storage.state, sheet, 'D4')).cells.D4?.value ?? null, null);
  const exported = unzipSync(await fs.readFile((await plain.export(changed.view.storage.state)).path));
  for (const name of Object.keys(parts)) if (name !== part) assertRetainedPart(name, parts[name], exported[name]);
  assert.deepEqual(await fs.readFile(fixture), original);
});
