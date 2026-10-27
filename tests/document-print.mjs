import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

test('owner-resident print jobs bind accepted state, invalidate after edits and never change dirty/history on reads', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-document-print-')),
    source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, createWorkbook());
  const document = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery') });
  try {
    const opened = await document.open(source),
      filled = await document.request({
        operation: 'apply',
        state: opened.state,
        commands: [
          {
            type: 'setCells',
            sheet: 'Sheet1',
            cells: Array.from({ length: 70 }, (_, i) => ({ cell: 'A' + (i + 1), value: 'row-' + (i + 1) })),
          },
        ],
      });
    const request = { documentId: filled.documentId, state: filled.state, sheet: 'Sheet1' },
      settings = { ...defaultPrintSettings(), area: 'A1:C70' };
    const before = document.status(),
      definition = await document.request({ operation: 'printDefinition', ...request });
    assert.equal(definition.canEdit, true);
    const job = await document.request({ operation: 'printPrepare', ...request, settings });
    assert.equal(job.pageCount, 2);
    assert.equal(job.pages.length, 1);
    const second = await document.request({ operation: 'printPage', ...request, jobId: job.id, pageNumber: 2 });
    assert(second.cells.some((cell) => cell.ref === 'A70' && cell.text === 'row-70'));
    const html = await document.request({ operation: 'printDocument', ...request, jobId: job.id });
    assert.equal(html.pageCount, 2);
    assert(html.html.includes('row-70'));
    assert.deepEqual(document.status(), before);
    await assert.rejects(document.request({ operation: 'printPage', ...request, jobId: job.id, pageNumber: 3 }), {
      code: 'internal-unknown-reference',
    });
    await assert.rejects(document.request({ operation: 'printPrepare', ...request, documentId: 'foreign', settings }), {
      code: 'session-stale',
    });
    await assert.rejects(
      document.request({ operation: 'printDocument', ...request, jobId: '00000000-0000-0000-0000-000000000000' }),
      { code: /^(object-stale|query-stale|session-stale)$/ }
    );
    const edited = await document.request({
      operation: 'apply',
      state: filled.state,
      commands: [{ sheet: 'Sheet1', cell: 'A70', value: 'changed' }],
    });
    await assert.rejects(document.request({ operation: 'printDocument', ...request, jobId: job.id }), {
      code: 'session-unsupported',
    });
    await assert.rejects(
      document.request({ operation: 'printPage', ...request, state: edited.state, jobId: job.id, pageNumber: 1 }),
      { code: /^(object-stale|query-stale|session-stale)$/ }
    );
    assert.deepEqual(await document.request({ operation: 'printRelease', ...request, jobId: job.id }), {
      released: true,
    });
    const undone = await document.request({ operation: 'undo' });
    assert.equal(undone.state, filled.state);
    const next = await document.request({ operation: 'printPrepare', ...request, settings });
    assert.notEqual(next.id, job.id);
    assert.equal(
      (await document.request({ operation: 'printPage', ...request, jobId: next.id, pageNumber: 2 })).cells.find(
        (cell) => cell.ref === 'A70'
      ).text,
      'row-70'
    );
    await document.request({ operation: 'printRelease', ...request, jobId: next.id });
    await assert.rejects(document.request({ operation: 'printDocument', ...request, jobId: next.id }), {
      code: /^(object-stale|query-stale|session-stale)$/,
    });
  } finally {
    await document.close();
    closeNativeWorkbookWorker();
    await fs.rm(root, { recursive: true, force: true });
  }
});

test('failed or cancelled print preparation preserves its previous job and accepted workbook', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-document-print-')),
    source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, createWorkbook());
  const document = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery') });
  try {
    const opened = await document.open(source),
      request = { documentId: opened.documentId, state: opened.state, sheet: 'Sheet1' },
      settings = { ...defaultPrintSettings(), area: 'A1:C3' };
    const job = await document.request({ operation: 'printPrepare', ...request, settings });
    await assert.rejects(
      document.request({ operation: 'printPrepare', ...request, settings: { ...settings, area: 'A1:XFD1048576' } }),
      /10,000|10000|bound|exceed/i
    );
    await assert.rejects(
      document.request(
        { operation: 'printPrepare', ...request, settings },
        {
          checkCancelled: () => {
            throw Error('cancelled');
          },
        }
      ),
      /cancelled/
    );
    assert.deepEqual(document.status(), opened);
    assert.equal((await document.request({ operation: 'printDocument', ...request, jobId: job.id })).pageCount, 1);
  } finally {
    await document.close();
    closeNativeWorkbookWorker();
    await fs.rm(root, { recursive: true, force: true });
  }
});
