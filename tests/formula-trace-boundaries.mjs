// Real native readonly boundary proof; expected edges/budgets are authored independently.
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { readWorkbookMetadata } from '../core/workbook-package.mjs';
import { calculationTraceFixture } from './support/calculation-trace-fixture.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-formula-trace-boundaries-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  call = (input) => client.call({ action: 'workbook', ...input });
async function indexed(label, bytes) {
  const root = path.join(output, label);
  await fs.mkdir(root);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, bytes);
  await call({ operation: 'index', root });
  const { view } = readWorkbookMetadata(await fs.readFile(path.join(root, 'metadata.xlsx')));
  return { root, source, view };
}
async function footprint(root) {
  return {
    source: await record(path.join(root, 'source.xlsx')),
    sqlite: await record(path.join(root, 'cells.sqlite')),
    metadata: await record(path.join(root, 'metadata.xlsx')),
    index: await record(path.join(root, 'index.json')),
  };
}
// Every case runs; failures are collected and asserted at the end.
const failures = [];
async function attempt(name, fn) {
  try {
    await fn();
  } catch (error) {
    failures.push(name + ': ' + (error.stack ?? error));
  }
}
try {
  await attempt(
    'Resumed partial formula counts toward2048 physical inspections; cross-sheet continuation stops after exactly2048 inspections',
    async () => {
      const parts = unzipSync(calculationTraceFixture()),
        n = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
      // First formula has two distinct relevant ranges. A limit1 reply resumes it.
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        `<worksheet xmlns="${n}"><sheetData>${Array.from({ length: 2047 }, (_, i) => `<row r="${i + 1}">${i === 0 ? '<c r="A1"><v>2</v></c>' : ''}<c r="B${i + 1}"><f>${i === 0 ? 'A1+SUM(A1:A2)' : '1+1'}</f><v>2</v></c></row>`).join('')}</sheetData></worksheet>`
      );
      parts['xl/worksheets/sheet2.xml'] = strToU8(
        `<worksheet xmlns="${n}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><f>Inputs!A1</f><v>2</v></c><c r="B1"><f>1+1</f><v>2</v></c></row></sheetData><tableParts count="1"><tablePart r:id="table1"/></tableParts></worksheet>`
      );
      parts['xl/worksheets/sheet3.xml'] = strToU8(`<worksheet xmlns="${n}"><sheetData/></worksheet>`);
      parts['xl/worksheets/sheet4.xml'] = parts['xl/worksheets/sheet3.xml'];
      const setup = await indexed('partial-cross-sheet', zipSync(parts));
      const before = await footprint(setup.root);
      const request = {
        operation: 'trace',
        root: setup.root,
        version: 0,
        view: setup.view,
        sheet: 'Inputs',
        cell: 'A1',
        direction: 'dependents',
      };
      const first = await call({ ...request, limit: 1, cursor: null });
      assert.equal(first.scannedFormulas, 1);
      assert.deepEqual(first.next, { sheetIndex: 0, after: null, offset: 1 });
      const resumed = await call({ ...request, limit: 100, cursor: first.next });
      assert.equal(resumed.scannedFormulas, 2048);
      assert.deepEqual(resumed.next, { sheetIndex: 1, after: 'A1', offset: 0 });
      assert.equal(resumed.edges.length, 2);
      assert.equal(resumed.edges[0].from.range, 'A1:A2');
      assert.deepEqual(resumed.edges[1], {
        from: { sheet: 'Inputs', range: 'A1' },
        to: { sheet: 'Data', range: 'A1' },
        kind: 'reference',
        via: 'formula',
      });
      const next = await call({ ...request, limit: 100, cursor: resumed.next });
      assert.equal(next.scannedFormulas, 1);
      assert.equal(next.next, null);
      assert.deepEqual(next.edges, []);
      assert.deepEqual(await footprint(setup.root), before);
    }
  );
  await attempt(
    'Native independently refuses forged worksheet/name/table resolver metadata without source/SQLite/index writes',
    async () => {
      const setup = await indexed('metadata-binding', calculationTraceFixture());
      const before = await footprint(setup.root);
      const request = {
        operation: 'trace',
        root: setup.root,
        version: 0,
        sheet: 'Report',
        cell: 'B2',
        direction: 'precedents',
        limit: 100,
        cursor: null,
      };
      const valid = await call({ ...request, view: setup.view });
      assert.equal(valid.edges[0].from.sheet, 'Inputs');
      assert.equal(valid.edges[0].from.range, 'A1');
      const mutants = [
        [
          'sheet names swapped',
          (v) => {
            [v.sheets[0].name, v.sheets[1].name] = [v.sheets[1].name, v.sheets[0].name];
          },
        ],
        [
          'sheet paths swapped',
          (v) => {
            [v.sheets[0].path, v.sheets[1].path] = [v.sheets[1].path, v.sheets[0].path];
          },
        ],
        [
          'name formula',
          (v) => {
            v.namedRanges[0].formula = 'Inputs!$A$2';
          },
        ],
        [
          'name scope',
          (v) => {
            v.namedRanges[0].sheet = 'Report';
          },
        ],
        [
          'name omitted',
          (v) => {
            v.namedRanges = [];
          },
        ],
        [
          'table range',
          (v) => {
            v.tables[0].ref = 'A1:B3';
          },
        ],
        [
          'table sheet',
          (v) => {
            v.tables[0].sheet = 'Inputs';
          },
        ],
        [
          'table column',
          (v) => {
            v.tables[0].columns[1].name = 'Other';
          },
        ],
        [
          'table header',
          (v) => {
            v.tables[0].headerRowCount = 0;
          },
        ],
        [
          'table part',
          (v) => {
            v.tables[0].path = 'xl/tables/fake.xml';
          },
        ],
        [
          'extra table',
          (v) => {
            v.tables.push({ ...v.tables[0], name: 'Invented' });
          },
        ],
      ];
      for (const [name, mutate] of mutants) {
        const view = structuredClone(setup.view);
        mutate(view);
        let refusal;
        try {
          const result = await call({ ...request, view });
          throw Error(`Native accepted forged ${name}`);
        } catch (error) {
          if (String(error).includes('Native accepted')) throw error;
          refusal = String(error);
          assert.match(refusal, /metadata.*source package/);
        }
      }
      const after = await footprint(setup.root);
      assert.deepEqual(after, before);
    }
  );
  await attempt(
    'Malformed actual table header identities/count/geometry cannot certify an arbitrary first-column edge',
    async () => {
      const cases = [
        ['duplicate', (xml) => xml.replace('name="Key"', 'name="Amount"')],
        ['case-duplicate', (xml) => xml.replace('name="Key"', 'name="amount"')],
        ['empty', (xml) => xml.replace('name="Key"', 'name=""')],
        ['wrong-count', (xml) => xml.replace('<tableColumns count="2">', '<tableColumns count="1">')],
        ['wrong-width', (xml) => xml.replace('ref="A1:B4"', 'ref="A1:C4"')],
      ];
      for (const [name, mutate] of cases) {
        const parts = unzipSync(calculationTraceFixture());
        parts['xl/tables/table1.xml'] = strToU8(mutate(strFromU8(parts['xl/tables/table1.xml'])));
        const setup = await indexed('table-' + name, zipSync(parts));
        const before = await footprint(setup.root);
        let result, refusal;
        try {
          result = await call({
            operation: 'trace',
            root: setup.root,
            version: 0,
            view: setup.view,
            sheet: 'Report',
            cell: 'B1',
            direction: 'precedents',
            limit: 100,
            cursor: null,
          });
        } catch (error) {
          refusal = String(error);
        }
        assert(refusal, `Native accepted malformed source table ${name}`);
        assert.match(refusal, /auditing table (column|columns)/i);
        assert.deepEqual(await footprint(setup.root), before);
      }
    }
  );
} finally {
  client.close();
  await fs.rm(output, { recursive: true, force: true });
}
assert.deepEqual(failures, []);
