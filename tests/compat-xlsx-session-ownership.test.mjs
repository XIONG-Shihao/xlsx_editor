import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createWorkbook, editWorkbook } from '../core/workbook-package.mjs';
import { WorkbookSession, readCalculatedWorkbook } from '../core/workbook-service.mjs';

const digest = (value) => createHash('sha256').update(value).digest('hex');
function freeze(value) {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) freeze(child);
  }
  return value;
}
for (const formula of [false, true])
  test(`unpublished calculation ownership preserves old views and receipts (${formula ? 'formula' : 'literal'})`, async () => {
    const source = editWorkbook(createWorkbook(), [
      { sheet: 'Sheet1', cell: 'A1', value: 2 },
      { sheet: 'Sheet1', cell: 'B1', ...(formula ? { formula: 'A1*3' } : { value: 'KEEP' }) },
    ]);
    const sourceHash = digest(source),
      session = new WorkbookSession(source);
    const original = freeze(await session.read()),
      originalJson = JSON.stringify(original);
    const request = {
      operationId: 'first',
      expectedRevision: original.revision,
      commands: [{ sheet: 'Sheet1', cell: 'A1', expectedValue: 2, value: 3 }],
    };
    const first = await session.apply(request);
    assert.equal(JSON.stringify(original), originalJson, 'Applying must not modify the accepted input view');
    assert.notStrictEqual(first.view, original);
    freeze(first.view);
    const firstJson = JSON.stringify(first.view),
      firstBytesHash = digest(first.bytes);
    const second = await session.apply({
      operationId: 'second',
      expectedRevision: first.view.revision,
      commands: [{ sheet: 'Sheet1', cell: 'A1', expectedValue: 3, value: 4 }],
    });
    assert.equal(JSON.stringify(first.view), firstJson, 'A subsequent batch must not rebind a prior receipt');
    assert.equal(digest(first.bytes), firstBytesHash);
    assert.strictEqual(await session.apply(request), first, 'Retry preserves the original receipt object and revision');
    assert.strictEqual(await session.read(), second.view, 'Retrying an older receipt cannot rewind current state');
    assert.equal(digest(source), sourceHash);
    if (formula) {
      for (const [view, expected] of [
        [first.view, 9],
        [second.view, 12],
      ]) {
        const cell = view.sheets[0].cells.B1;
        assert.equal(cell.value, expected);
        assert.equal(cell.cachedValue, expected);
        assert.equal(cell.calculation.value, expected);
        assert.equal(cell.calculation.revision, view.revision);
      }
    }
    const reopened = await readCalculatedWorkbook(second.bytes);
    assert.equal(reopened.revision, second.view.revision);
    assert.equal(reopened.sheets[0].cells.A1.value, 4);
    assert.equal(reopened.sheets[0].cells.B1[formula ? 'cachedValue' : 'value'], formula ? 12 : 'KEEP');

    freeze(second.view);
    const acceptedJson = JSON.stringify(second.view),
      acceptedHash = digest(session.bytes);
    await assert.rejects(
      session.apply({
        operationId: 'failed',
        expectedRevision: second.view.revision,
        commands: [
          { sheet: 'Sheet1', cell: 'A1', expectedValue: 4, value: 99 },
          { sheet: 'Sheet1', cell: 'A1', expectedValue: -1, value: 100 },
        ],
      }),
      /Conflict/
    );
    assert.strictEqual(await session.read(), second.view);
    assert.equal(JSON.stringify(second.view), acceptedJson);
    assert.equal(digest(session.bytes), acceptedHash);
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(
      session.apply(
        {
          operationId: 'cancelled',
          expectedRevision: second.view.revision,
          commands: [{ sheet: 'Sheet1', cell: 'A1', value: 100 }],
        },
        { signal: controller.signal }
      ),
      { code: 'operation-cancelled' }
    );
    assert.strictEqual(await session.read(), second.view);
    assert.equal(digest(session.bytes), acceptedHash);
  });
