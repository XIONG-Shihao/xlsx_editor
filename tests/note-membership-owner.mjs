/** Original JS owners with real package membership parsing and a controlled
 * native indexing/ledger boundary. This tests envelope/lifecycle preservation;
 * it does not qualify Excel semantics or native preservation. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { validateNoteMembership } from '../core/note-membership.mjs';
import { engineQuestion } from './support/engine-questions.mjs';

const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
function packageBytes({ protectedSheet = false, malformed = false } = {}) {
  const parts = unzipSync(createWorkbook());
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>'
    )
  );
  parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    '<Relationships xmlns="' +
      relationships +
      '"><Relationship Id="notes" Type="' +
      officeRelationships +
      '/comments" Target="../comments1.xml"/></Relationships>'
  );
  parts['xl/comments1.xml'] = strToU8(
    '<comments xmlns="' +
      main +
      '"><authors><author>Controlled author</author></authors><commentList><comment ref="D5" authorId="0"><text><t/></text></comment><comment ref="' +
      (malformed ? 'bad' : 'B3') +
      '" authorId="0"><text><t>Membership text is not exposed</t></text></comment></commentList></comments>'
  );
  if (protectedSheet)
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<sheetProtection sheet="1"/></worksheet>')
    );
  return Buffer.from(zipSync(parts, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') }));
}
async function fixture(t, options = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-note-membership-owner-')),
    source = path.join(directory, 'source.xlsx'),
    bytes = packageBytes(options),
    sourceRevision = digest(bytes),
    calls = [];
  await fs.writeFile(source, bytes);
  const native = async (input) => {
    const answer = engineQuestion(input);
    if (answer) return answer;
    calls.push(structuredClone(input));
    if (input.operation === 'copy-source') {
      await fs.copyFile(input.source, path.join(input.root, 'source.xlsx'));
      return {};
    }
    if (input.operation === 'index') {
      const index = {
        revision: sourceRevision,
        sheets: [
          {
            name: 'Sheet1',
            path: 'xl/worksheets/sheet1.xml',
            rows: 20,
            columns: 8,
            hiddenRows: [],
            visibleRows: [],
            rowHeights: {},
            storedCells: 0,
            formulaCount: 0,
          },
        ],
      };
      await fs.writeFile(path.join(input.root, 'metadata.xlsx'), bytes);
      await fs.writeFile(path.join(input.root, 'index.json'), JSON.stringify(index));
      return index;
    }
    if (input.operation === 'artifact-write') {
      await fs.writeFile(path.join(input.root, input.name), input.data);
      return {};
    }
    if (input.operation === 'compact') return { retainedVersions: input.keep.length };
    if (input.operation === 'calculation-policy' && input.method === 'initialize')
      return {
        policy: {
          schema: 1,
          mode: 'auto',
          recalculateBeforeSave: true,
          pendingFormulaCells: 0,
          canCalculateSheet: true,
          canChangeOptions: true,
        },
      };
    assert.fail('Unexpected controlled native operation: ' + input.operation);
  };
  const owner = new SpreadsheetDocument({
    native,
    recoveryDirectory: path.join(directory, 'recovery'),
    indexDirectory: path.join(directory, 'indices'),
    cacheIdentity: null,
  });
  t.after(async () => {
    await owner.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await owner.open(source, options.prepared ? { fresh: true, prepared: true } : {});
  return { directory, source, sourceRevision, owner, native, calls };
}
const request = (f) => ({
  operation: 'noteMembership',
  documentId: f.owner.status().documentId,
  state: f.owner.status().state,
  sheet: 'Sheet1',
});
const receipt = (f) => {
  const status = f.owner.status();
  return {
    documentId: status.documentId,
    ownerId: status.view.storage.sessionId,
    state: status.state,
    revision: status.view.revision,
  };
};
async function footprint(f) {
  const files = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const file = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(file);
      else files.push({ name: path.relative(f.directory, file), sha256: digest(await fs.readFile(file)) });
    }
  }
  await walk(f.directory);
  return { status: f.owner.status(), files: files.sort((a, b) => a.name.localeCompare(b.name)) };
}
function observeReader(t) {
  const original = IndexedWorkbook.prototype.noteMembership;
  let hold;
  const calls = [];
  t.mock.method(IndexedWorkbook.prototype, 'noteMembership', async function (state, sheet) {
    calls.push({ ownerId: this.id, state, sheet });
    const result = await original.call(this, state, sheet);
    if (hold) {
      const pending = hold;
      hold = undefined;
      pending.started();
      await pending.wait;
    }
    return result;
  });
  return {
    calls,
    hold: () => {
      let started, release;
      const ready = new Promise((r) => (started = r)),
        wait = new Promise((r) => (release = r));
      hold = { started, wait };
      return { ready, release };
    },
  };
}

test('original document owner binds selected source/head and performs only a membership read', async (t) => {
  const f = await fixture(t),
    seen = observeReader(t),
    before = await footprint(f),
    nativeCalls = f.calls.length,
    reply = (await f.owner.request(request(f))).noteMembership;
  validateNoteMembership(reply, {
    sheet: 'Sheet1',
    receipt: receipt(f),
    source: { version: 0, sourceRevision: f.sourceRevision, sheetPath: 'xl/worksheets/sheet1.xml' },
  });
  assert.deepEqual(reply.notes, ['B3', 'D5']);
  assert.deepEqual(reply.threadedPlaceholders, []);
  assert(!JSON.stringify(reply).includes('Controlled author'));
  assert(!JSON.stringify(reply).includes('Membership text'));
  assert.equal(seen.calls.length, 1);
  assert.equal(f.calls.length, nativeCalls, 'membership does not call a native query or writer');
  assert.deepEqual(await footprint(f), before);
});

test('foreign/stale/injected document authority refuses before the membership reader', async (t) => {
  const f = await fixture(t),
    seen = observeReader(t),
    query = request(f),
    before = await footprint(f);
  for (const extra of [
    { documentId: 'other' },
    { state: 'unaccepted-head' },
    { sessionId: 'other' },
    { commands: [] },
    { range: 'A1:H20' },
    { source: {} },
    { receipt: {} },
  ])
    await assert.rejects(() => f.owner.request({ ...query, ...extra }), /another|stale|fields/);
  await assert.rejects(() => f.owner.request({ ...query, sheet: 'Other' }), { code: 'internal-unknown-reference' });
  assert.equal(
    seen.calls.length,
    1,
    'only the validated unknown-sheet request reaches IndexedWorkbook; package reader refuses it'
  );
  assert.deepEqual(await footprint(f), before);
});

test('queued asynchronous document reads retain the original request after caller retargeting', async (t) => {
  const f = await fixture(t),
    seen = observeReader(t),
    original = request(f),
    before = await footprint(f),
    expected = receipt(f),
    timing = seen.hold();
  const pending = f.owner.request(original);
  original.documentId = 'other';
  original.state = 'other';
  original.sheet = 'Other';
  original.commands = [];
  await timing.ready;
  timing.release();
  const reply = (await pending).noteMembership;
  validateNoteMembership(reply, { sheet: 'Sheet1', receipt: expected });
  assert.deepEqual(await footprint(f), before);
});

test('prepared and protected XLSX reads work without activation or edit authority', async (t) => {
  for (const options of [{ prepared: true }, { protectedSheet: true }]) {
    const f = await fixture(t, options),
      before = await footprint(f),
      reply = (await f.owner.request(request(f))).noteMembership;
    assert.equal(reply.status, 'complete');
    assert.deepEqual(reply.notes, ['B3', 'D5']);
    assert.deepEqual(await footprint(f), before);
    if (options.protectedSheet) assert.equal(f.owner.status().view.protectedSheets.length, 1);
  }
});

test('incomplete membership returns null authority through the actual document route', async (t) => {
  const f = await fixture(t, { malformed: true }),
    before = await footprint(f),
    reply = (await f.owner.request(request(f))).noteMembership;
  validateNoteMembership(reply, { sheet: 'Sheet1', receipt: receipt(f) });
  assert.equal(reply.status, 'incomplete');
  assert.equal(reply.notes, null);
  assert.equal(reply.threadedPlaceholders, null);
  assert(reply.reasons.length);
  assert.deepEqual(await footprint(f), before);
});

test('cancelled and closed document owners reject a held completed membership result', async (t) => {
  const f = await fixture(t),
    seen = observeReader(t),
    before = await footprint(f),
    timing = seen.hold();
  let cancelled = false;
  const pending = f.owner.request(request(f), {
    checkCancelled: () => {
      if (cancelled) throw Error('Notes test cancellation');
    },
  });
  await timing.ready;
  cancelled = true;
  timing.release();
  await assert.rejects(() => pending, /cancellation/);
  assert.deepEqual(await footprint(f), before);
  const retired = await fixture(t),
    retiring = seen.hold(),
    late = retired.owner.request(request(retired));
  await retiring.ready;
  const closing = retired.owner.close();
  retiring.release();
  await assert.rejects(() => late, /cancelled|closed|retired/);
  await closing;
});
