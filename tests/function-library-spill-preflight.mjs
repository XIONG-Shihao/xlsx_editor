/** SEQUENCE spills on the mixed scalar function fixture, and Undo returns the prior state. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { functionLibraryFixture, functionLibraryCases } from './support/function-library-fixture.mjs';
const out = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-preflight-'));
const hash = (b) => createHash('sha256').update(b).digest('hex'),
  client = createWorkbookNativeClient(nativeWorkbookBinary()),
  file = path.join(out, 'source.xlsx'),
  bytes = functionLibraryFixture();
await fs.writeFile(file, bytes);
const owner = new SpreadsheetDocument({
    recoveryDirectory: path.join(out, 'recovery'),
    native: (r) => client.call(r),
    cacheIdentity: null,
  });
try {
  await owner.open(file);
  let s = await owner.request({ operation: 'status' });
  s = await owner.request({
    operation: 'apply',
    state: s.state,
    documentId: s.documentId,
    commands: [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          ...functionLibraryCases.map((c) => ({ cell: c.cell, formula: c.id === 'SUM' ? 'SUM(2,3,4)' : c.formula })),
          { cell: 'C33', formula: 'SUM(A1:A3)' },
        ],
      },
    ],
  });
  const prior = s.state;
  s = await owner.request({
    operation: 'apply',
    state: s.state,
    documentId: s.documentId,
    commands: [{ type: 'setCells', sheet: 'Sheet1', cells: [{ cell: 'F10', formula: 'SEQUENCE(3,2)' }] }],
  });
  const cells = (
    await owner.request({
      operation: 'read',
      state: s.state,
      documentId: s.documentId,
      sheet: 'Sheet1',
      range: 'F10:G12',
    })
  ).cells;
  assert.equal(cells.F11.value, 3);
  assert.deepEqual(cells.F11.spill, { schema: 1, master: 'F10', range: 'F10:G12', role: 'follower' });
  for (const [ref, value] of Object.entries({ F10: 1, G10: 2, F11: 3, G11: 4, F12: 5, G12: 6 }))
    assert.equal(cells[ref].value, value);
  await owner.request({ operation: 'undo' });
  assert.equal((await owner.request({ operation: 'status' })).state, prior);
  assert.equal(hash(await fs.readFile(file)), hash(bytes));
} finally {
  await owner.close();
  client.close();
  await fs.rm(out, { recursive: true, force: true });
}
