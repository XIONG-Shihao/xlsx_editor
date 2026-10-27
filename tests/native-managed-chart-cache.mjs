import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111';
async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-native-chart-cache-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'input.xlsx');
  await fs.copyFile(path.resolve('tests/support/w7-reference-data/source.xlsx'), source);
  let exported;
  const bridge = async (input) => {
    const result = await native(input);
    if (input.operation === 'export') exported = { ...input };
    return result;
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No compatibility writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(source);
  t.after(() => book.close());
  const made = await book.apply('0', [
      {
        operation: 'createChart',
        sheet: 'Sales',
        objectId: id,
        expectedObjectRevision: null,
        chart: { kind: 'column', range: 'A1:C4', title: 'Monthly sales' },
      },
    ]),
    changed = await book.apply(made.view.storage.state, [{ sheet: 'Sales', cell: 'B2', value: 21 }]);
  const saved = await book.export(changed.view.storage.state),
    root = exported.root,
    original = await fs.readFile(saved.path),
    metadata = JSON.parse(await fs.readFile(path.join(root, 'export-metadata.json'), 'utf8')),
    part = Object.keys(metadata.managedChartCaches)[0];
  return { root, exported, original, metadata, part, saved };
}
test('native cache proof refuses stale source, graph/path claims, changed non-cache XML and wrong accepted values before output publication', async (t) => {
  const f = await fixture(t),
    cases = [],
    clone = () => structuredClone(f.metadata),
    bad = (change) => {
      const value = clone();
      change(value);
      cases.push(value);
    };
  bad((value) => (value.managedChartCaches[f.part].sourceSha256 = '0'.repeat(64)));
  bad((value) => (value.managedChartCaches['xl/charts/chart1.xml'] = value.managedChartCaches[f.part]));
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        'Monthly sales',
        'Changed title'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:v>21</c:v>',
        '<c:v>22</c:v>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:v>21</c:v>',
        '<c:v>NaN</c:v>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:v>Sales</c:v>',
        '<c:v>Wrong series</c:v>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:v>Jan</c:v>',
        '<c:v>Wrong category</c:v>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        'idx="1"><c:v>3',
        'idx="0"><c:v>3'
      ))
  );
  bad((value) => (value.replacements[f.part] = value.managedChartCaches[f.part].xml));
  bad((value) => value.removed.push(f.part));
  bad((value) => (value.managedChartCaches[f.part].extra = 'unqualified'));
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:ptCount val="3"/>',
        '<c:ptCount val="3"><c:v>unqualified</c:v></c:ptCount>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:pt idx="0"><c:v>21</c:v></c:pt>',
        '<c:pt idx="0"><c:v>21</c:v><c:layout/></c:pt>'
      ))
  );
  bad(
    (value) =>
      (value.managedChartCaches[f.part].xml = value.managedChartCaches[f.part].xml.replace(
        '<c:numCache>',
        '<c:numCache>unexpected'
      ))
  );
  bad((value) => (value.managedChartCaches[f.part].xml += ' '.repeat(4 * 1024 * 1024)));
  for (const metadata of cases) {
    await fs.writeFile(path.join(f.root, 'export-metadata.json'), JSON.stringify(metadata));
    await assert.rejects(native(f.exported), /chart|Chart|managed|Managed/);
    assert.deepEqual(await fs.readFile(f.saved.path), f.original);
    await assert.rejects(fs.stat(path.join(f.root, 'export.partial')), (error) => error.code === 'ENOENT');
  }
  await fs.writeFile(path.join(f.root, 'export-metadata.json'), JSON.stringify(f.metadata));
  const valid = await native(f.exported);
  assert.equal(valid.file, 'export.xlsx');
  assert.deepEqual(await fs.readFile(f.saved.path), f.original);
});
