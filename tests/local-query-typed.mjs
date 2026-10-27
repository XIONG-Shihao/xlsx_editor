import './support/scratch-dir.mjs';
/** Actual typed query model/registry/main spool. Owner replies below are
 * controlled JS simulations; native and shared replay are separate proofs. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { createBlobDelimitedSource } from '../core/delimited-import.mjs';
import {
  inspectLocalQuery,
  localQueryPages,
  stageLocalQuery,
  typedQueryCellIdentity,
  queryCellFingerprint,
  validateLocalQuery,
} from '../core/local-query.mjs';
import { LocalQueryConnections } from '../adapters/local-query-connections.mjs';
import { LocalQueryHost } from '../adapters/local-query-host.mjs';
import { QUERY_INSPECTION_LIMITS } from '../adapters/local-query-inspection.mjs';
import { convertTextField } from '../core/text-conversion.mjs';
import { cellPosition, columnName, rangeBounds } from '../core/workbook-commands.mjs';
import { pasteDescriptorIdentity } from '../core/paste-protocol.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/typed-query-checks-'));
const check = (name, fn) => {
  test(name, { timeout: 60000 }, async (t) => {
    await fn(t);
  });
};
const source = (text) => createBlobDelimitedSource(new Blob([text]), { name: 'data.csv', chunkBytes: 17 }),
  query = (policies) => ({ columns: [], filters: [], keepHeader: true, ...(policies ? { policies } : {}) }),
  fields = async (data, inspection) => {
    const result = [];
    for await (const page of localQueryPages(data, inspection)) result.push(...page.cells);
    return result;
  };
async function fixture(
  t,
  text = 'id,n,flag,day\n00123,-0,TRUE,2011-07-05\n=SUM(A1),1.00,false,\n',
  policies = ['text', 'number', 'boolean', 'iso-date']
) {
  const dir = await fs.mkdtemp(path.join(output, 'fixture-')),
    filename = path.join(dir, 'source.csv');
  await fs.writeFile(filename, text);
  const store = await LocalQueryConnections.open(path.join(dir, 'connections.json')),
    record = await store.create(
      'doc',
      { name: 'Typed query', target: { sheet: 'S', start: 'B3' }, query: query(policies) },
      filename
    ),
    scope = { documentId: 'doc', ownerId: 'owner', state: 'head' },
    cells = {},
    requests = [],
    uploads = new Map(),
    receipts = new Map();
  let calendar = false,
    lost = false,
    accepted = 0;
  const context = {
    current: () => ({ ...scope }),
    calendar: () => calendar,
    request: async (op, input) => {
      requests.push({ op, input: structuredClone(input) });
      if (op === 'read') {
        const b = rangeBounds(input.range);
        return {
          cells: Object.fromEntries(
            Object.entries(cells).filter(([ref]) => {
              const p = cellPosition(ref);
              return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
            })
          ),
        };
      }
      if (op === 'importBegin') {
        const id = randomUUID();
        uploads.set(id, {
          plan: {
            shape: input.shape,
            count: input.count,
            policy: input.policy,
            ...(input.conversion ? { conversion: input.conversion } : {}),
          },
          cells: [],
          sequence: 0,
        });
        return { id, sequence: 0 };
      }
      if (op === 'importUpload') {
        const u = uploads.get(input.uploadId);
        u.cells.push(...input.cells);
        return { id: input.uploadId, sequence: ++u.sequence, count: u.cells.length };
      }
      if (op === 'importSeal') {
        const u = uploads.get(input.uploadId),
          descriptor = {
            type: 'importCells',
            id: input.uploadId,
            sha256: 'a'.repeat(64),
            bytes: 100,
            ...u.plan,
            ...(u.plan.policy === 'typed-text-v1' ? { schema: 'workbook.xlsx.staged-paste.v2' } : {}),
          };
        u.descriptor = descriptor;
        return { descriptor };
      }
      if (op === 'importReceipt') return receipts.get(input.descriptor.id) ?? { found: false };
      return {};
    },
    apply: async (parent, descriptor) => {
      assert.equal(parent, scope.state);
      const u = uploads.get(descriptor.id),
        start = cellPosition(descriptor.shape.start),
        c = descriptor.conversion;
      for (const [index, raw] of u.cells) {
        const row = Math.floor(index / descriptor.shape.columns),
          col = index % descriptor.shape.columns,
          policy =
            c && row < c.outputRows && col < c.outputColumns && !(c.keepHeader && row === 0) ? c.policies[col] : 'text',
          v = convertTextField(raw, policy, { date1904: c?.date1904 });
        assert.equal(v.status, 'converted');
        const ref = columnName(start.col + col) + (start.row + row + 1);
        cells[ref] = {
          value: v.value,
          dataType: v.kind === 'date' ? 'number' : v.kind,
          rawValue: typeof v.value === 'number' ? String(v.value) : String(v.value),
          ...(v.kind === 'date' ? { format: 14 } : {}),
        };
      }
      accepted++;
      scope.state = 'accepted-' + accepted;
      const receipt = { found: true, head: scope.state, parentState: parent, descriptor };
      receipts.set(descriptor.id, receipt);
      if (lost) throw Error('Simulated durable reply lost');
      return { state: scope.state };
    },
  };
  const host = () => new LocalQueryHost({ store, document: 'doc', directory: path.join(dir, 'private'), context });
  t.after(async () => host().disposeInspections('owner'));
  return {
    dir,
    filename,
    store,
    record,
    scope,
    cells,
    requests,
    uploads,
    receipts,
    context,
    host,
    calendar: (v) => {
      calendar = v;
    },
    lose: (v) => {
      lost = v;
    },
    accepted: () => accepted,
  };
}
check(
  'closed saved policies preserve v1; full conversion keeps source-text filters/header/empty and raw staged pages',
  async () => {
    assert.deepEqual(validateLocalQuery(query()), query());
    for (const bad of [
      { ...query(), columnTypes: ['number'] },
      { ...query(), policies: ['date'] },
      { ...query(), policies: Array(2) },
    ])
      assert.throws(() => validateLocalQuery(bad));
    const data = source('flag,n,date\n00123,1.00,2011-07-05\n123,9007199254740993,2011-07-05\n00123,,\n'),
      definition = {
        columns: [2, 3],
        filters: [{ column: 1, op: 'equals', value: '00123', caseSensitive: true }],
        keepHeader: true,
        policies: ['number', 'iso-date'],
      },
      seen = [],
      i = await inspectLocalQuery(data, {
        sheet: 'S',
        start: 'C8',
        query: definition,
        date1904: false,
        onCell: (c) => seen.push(c),
      });
    assert.equal(i.outputRows, 3);
    assert.equal(i.conversionSummary.rejected, 0);
    assert.deepEqual(
      seen.map((c) => c.result.kind),
      ['text', 'text', 'number', 'date', 'text', 'text']
    );
    assert.equal(seen[2].sourceRow, 2);
    assert.equal(seen[4].sourceRow, 4);
    assert.deepEqual(await fields(data, i), [
      [0, 'n'],
      [1, 'date'],
      [2, '1.00'],
      [3, '2011-07-05'],
      [4, ''],
      [5, ''],
    ]);
    assert.equal(i.policy, 'typed-text-v1');
    assert.equal(i.conversion.date1904, false);
  }
);
check(
  'rejections after row20 are inspected and block staging before any begin; explicit unknown calendar refuses',
  async () => {
    const data = source('n\n' + Array.from({ length: 25 }, (_, i) => (i === 24 ? '9007199254740993' : '1')).join('\n')),
      i = await inspectLocalQuery(data, { sheet: 'S', start: 'A1', query: query(['number']), date1904: false });
    assert.equal(i.preview.length, 20);
    assert.equal(i.conversionSummary.rejected, 1);
    let began = 0;
    await assert.rejects(
      stageLocalQuery(data, i, {
        state: 'head',
        sink: {
          begin: () => {
            began++;
          },
          upload() {},
          seal() {},
          cancel() {},
          prune() {},
        },
      }),
      { code: /^(query-refused|query-unsupported|rules-refused)$/ }
    );
    assert.equal(began, 0);
    await assert.rejects(
      inspectLocalQuery(source('x\n2011-07-05'), { sheet: 'S', start: 'A1', query: query(['iso-date']) }),
      /calendar/
    );
  }
);
check(
  'both calendars and smaller typed results clear old tail as raw empty Text, never duplicate typed padding',
  async () => {
    for (const date1904 of [false, true]) {
      const data = source('n,day\n-0,2011-07-05'),
        i = await inspectLocalQuery(data, {
          sheet: 'S',
          start: 'A1',
          query: query(['number', 'iso-date']),
          date1904,
          previousShape: { sheet: 'S', start: 'A1', rows: 4, columns: 3 },
        }),
        raw = await fields(data, i);
      assert.deepEqual(
        raw.slice(6),
        Array.from({ length: 6 }, (_, n) => [n + 6, ''])
      );
      assert.deepEqual(i.conversion, {
        schema: 1,
        policies: ['number', 'iso-date'],
        date1904,
        keepHeader: true,
        outputRows: 2,
        outputColumns: 2,
      });
      const expected = queryCellFingerprint();
      raw.forEach(([index, text]) => {
        const r = Math.floor(index / 3),
          c = index % 3,
          v = convertTextField(text, r === 1 && c < 2 ? i.conversion.policies[c] : 'text', { date1904 });
        expected.add(
          index,
          v.kind,
          typeof v.value === 'number' ? (Object.is(v.value, -0) ? '-0' : String(v.value)) : v.value
        );
      });
      assert.equal(expected.finish(), i.outputFingerprint);
      expected.close();
    }
  }
);
check(
  'typed managed identities distinguish blank/text/error/boolean/sign, validate raw precision and owned date14',
  async () => {
    const shape = { sheet: 'S', start: 'A1', rows: 1, columns: 2 },
      conversion = {
        schema: 1,
        policies: ['number', 'iso-date'],
        date1904: false,
        keepHeader: false,
        outputRows: 1,
        outputColumns: 2,
      },
      identity = (cell, index = 0) => typedQueryCellIdentity(cell, { shape, conversion, index });
    assert.notDeepEqual(identity(undefined), identity({ value: '', dataType: 'text' }));
    assert.notDeepEqual(identity({ value: 1 }), identity({ value: '1' }));
    assert.notDeepEqual(identity({ value: true }), identity({ value: 1 }));
    assert.notDeepEqual(identity({ value: '#N/A', dataType: 'text' }), identity({ value: '#N/A', dataType: 'error' }));
    assert.deepEqual(identity({ value: -0, rawValue: '0' }), identity({ value: -0, rawValue: '-0' }));
    assert.notDeepEqual(identity({ value: -0, rawValue: '0' }), identity({ value: 0, rawValue: '0' }));
    assert.equal(identity({ value: 9007199254740992, rawValue: '9007199254740993' })[0], 'protected');
    assert.equal(identity({ value: 1.23, rawValue: '1.2300000000000001' })[0], 'protected');
    assert.deepEqual(identity({ value: 1, rawValue: '1.00' }), ['number', '1']);
    assert.equal(identity({ value: 40729, format: 14 }, 1)[0], 'date');
    assert.equal(identity({ value: 40729, format: 0 }, 1)[0], 'protected');
    assert.equal(identity({ value: 2, calculation: { status: 'stored', pending: true } })[0], 'protected');
  }
);
check(
  'main-only issue pages expose every >20-row refusal and notice, with no renderer paths or typed staging',
  async (t) => {
    const f = await fixture(
        t,
        'n\n' + Array.from({ length: 205 }, (_, i) => (i % 2 ? '1.00' : '9007199254740993')).join('\n'),
        ['number']
      ),
      p = (await f.host().preview(f.record.id, 1)).queryPreview,
      c = p.conversionInspection;
    assert.equal(c.rejected, 103);
    assert.equal(c.changed, 102);
    assert.equal(c.totalIssues, 205);
    assert.equal(c.issues.length, 100);
    assert.equal(c.nextOffset, 100);
    assert.equal(c.issues[0].sourceRef, 'A2');
    assert.equal(c.issues[0].destinationRef, 'B4');
    const all = [...c.issues];
    let next = c.nextOffset;
    while (next !== null) {
      const page = (await f.host().preview(f.record.id, 1, { offset: next, expectedConversionRevision: c.revision }))
        .queryPreview.conversionInspection;
      all.push(...page.issues);
      next = page.nextOffset;
    }
    assert.equal(all.length, 205);
    assert.equal(new Set(all.map((i) => i.index)).size, 205);
    assert.equal(all.at(-1).sourceRef, 'A206');
    assert(!JSON.stringify(p).includes(f.dir));
    await assert.rejects(f.host().refresh(f.record.id, 1, { expectedConversionRevision: c.revision }), {
      code: /^(query-refused|query-unsupported|rules-refused)$/,
    });
    assert.equal(f.requests.filter((r) => r.op === 'importBegin').length, 0);
    await f.host().disposeInspections('owner');
    assert.deepEqual(await fs.readdir(path.join(f.dir, 'private')), []);
  }
);
check(
  'review token required and retired by source/settings/calendar/head/runtime changes, including transition to Text',
  async (t) => {
    const f = await fixture(t);
    await assert.rejects(f.host().refresh(f.record.id, 1), /Preview/);
    let c = (await f.host().preview(f.record.id, 1)).queryPreview.conversionInspection;
    f.scope.ownerId = 'replacement';
    await assert.rejects(
      f.host().refresh(f.record.id, 1, { expectedConversionRevision: c.revision }),
      /expired|changed/
    );
    f.scope.ownerId = 'owner';
    f.scope.state = 'new-head';
    await assert.rejects(
      f.host().preview(f.record.id, 1, { expectedConversionRevision: c.revision }),
      /expired|changed/
    );
    f.scope.state = 'head';
    f.calendar(true);
    await assert.rejects(
      f.host().refresh(f.record.id, 1, { expectedConversionRevision: c.revision }),
      /expired|changed/
    );
    f.calendar(false);
    await fs.appendFile(f.filename, 'x,2,TRUE,2011-07-05\n');
    await assert.rejects(
      f.host().preview(f.record.id, 1, { expectedConversionRevision: c.revision }),
      /expired|changed/
    );
    const fresh = (await f.host().preview(f.record.id, 1)).queryPreview;
    await f.store.update('doc', f.record.id, 1, { name: f.record.name, target: f.record.target, query: query() });
    await assert.rejects(
      f.host().refresh(f.record.id, 2, { expectedConversionRevision: fresh.conversionInspection.revision }),
      /expired|changed/
    );
    assert.equal(f.accepted(), 0);
  }
);
check(
  'reviewed raw typed refresh awaits one accepted batch and persists prior profile; cold spelling changes are no conflict',
  async (t) => {
    const f = await fixture(t);
    const p = (await f.host().preview(f.record.id, 1)).queryPreview;
    const result = await f
      .host()
      .refresh(f.record.id, 1, { expectedConversionRevision: p.conversionInspection.revision });
    assert.equal(result.refreshed, true);
    assert.equal(f.accepted(), 1);
    assert.equal(result.connection.last.fingerprintProfile, 'typed-semantic-v1');
    assert.equal(result.connection.last.conversion.date1904, false);
    assert(Object.is(f.cells.C4.value, -0));
    assert.equal(f.cells.E4.format, 14);
    const descriptor = f.requests.find((r) => r.op === 'importSeal').input.uploadId;
    assert.equal(f.uploads.get(descriptor).plan.policy, 'typed-text-v1');
    assert(f.uploads.get(descriptor).cells.every((c) => typeof c[1] === 'string'));
    f.cells.C4.rawValue = '-0';
    f.cells.C5.rawValue = '1.00';
    const next = (await f.host().preview(f.record.id, 2)).queryPreview;
    await f.host().refresh(f.record.id, 2, { expectedConversionRevision: next.conversionInspection.revision });
    assert.equal(f.accepted(), 2);
    f.cells.E4.format = 0;
    const newest = (await f.host().preview(f.record.id, 3)).queryPreview;
    await assert.rejects(
      f.host().refresh(f.record.id, 3, { expectedConversionRevision: newest.conversionInspection.revision }),
      { code: /^(layout-stale|pivot-stale|query-stale)$/ }
    );
    assert.equal(f.accepted(), 2);
  }
);
check('lost durable typed reply retains exact v2 metadata and reconciles only matching owner receipt', async (t) => {
  const f = await fixture(t),
    p = (await f.host().preview(f.record.id, 1)).queryPreview;
  f.lose(true);
  await assert.rejects(
    f.host().refresh(f.record.id, 1, { expectedConversionRevision: p.conversionInspection.revision }),
    /reply lost/
  );
  assert.equal(f.accepted(), 1);
  const pending = (await f.store.granted('doc', f.record.id)).pending;
  assert.equal(pending.descriptor.schema, 'workbook.xlsx.staged-paste.v2');
  const bad = structuredClone(f.receipts.get(pending.descriptor.id));
  bad.descriptor.conversion.date1904 = true;
  assert.notEqual(pasteDescriptorIdentity(bad.descriptor), pasteDescriptorIdentity(pending.descriptor));
  await assert.rejects(f.store.finalize('doc', f.record.id, pending.id, bad), /does not match/);
  const reconciled = await f.host().list();
  assert.equal(reconciled.entries[0].pending, undefined);
  assert.equal(reconciled.entries[0].last.acceptedHead, 'accepted-1');
  assert.equal(reconciled.entries[0].last.fingerprintProfile, 'typed-semantic-v1');
  assert.equal(f.accepted(), 1);
});
check(
  'explicit cancellation/held progress owner retirement sends no begin; capped cache revokes oldest and releases private files',
  async (t) => {
    const f = await fixture(t),
      abort = new AbortController();
    abort.abort();
    await assert.rejects(f.host().preview(f.record.id, 1, { signal: abort.signal }));
    assert.equal(f.accepted(), 0);
    let advanced = false;
    await assert.rejects(
      f.host().preview(f.record.id, 1, {
        onProgress: () => {
          if (!advanced) {
            advanced = true;
            f.scope.ownerId = 'new';
          }
        },
      }),
      { code: 'query-stale' }
    );
    f.scope.ownerId = 'owner';
    const tokens = [];
    for (let n = 0; n < QUERY_INSPECTION_LIMITS.receipts + 1; n++) {
      const record = await f.store.create(
        'doc',
        { name: 'Q' + n, target: f.record.target, query: query(['text', 'number', 'boolean', 'iso-date']) },
        f.filename
      );
      tokens.push({
        id: record.id,
        token: (await f.host().preview(record.id, 1)).queryPreview.conversionInspection.revision,
      });
    }
    await assert.rejects(f.host().preview(tokens[0].id, 1, { expectedConversionRevision: tokens[0].token }), {
      code: /^(object-stale|query-stale|session-stale)$/,
    });
    await f.host().disposeInspections('owner');
    assert.deepEqual(await fs.readdir(path.join(f.dir, 'private')), []);
  }
);

check(
  'late disposal of an old owner cannot revoke replacement reviews; setting updates revoke only original owned review',
  async (t) => {
    const f = await fixture(t),
      old = (await f.host().preview(f.record.id, 1)).queryPreview.conversionInspection;
    f.scope.ownerId = 'new-owner';
    const next = (await f.host().preview(f.record.id, 1)).queryPreview.conversionInspection;
    await f.host().disposeInspections('owner');
    assert.equal(
      (await f.host().preview(f.record.id, 1, { expectedConversionRevision: next.revision })).queryPreview
        .conversionInspection.revision,
      next.revision
    );
    await assert.rejects(f.host().preview(f.record.id, 1, { expectedConversionRevision: old.revision }), {
      code: /^(object-stale|query-stale|session-stale)$/,
    });
    await f.host().update(f.record.id, 1, {
      name: 'Updated',
      target: f.record.target,
      query: query(['text', 'number', 'boolean', 'iso-date']),
    });
    await assert.rejects(f.host().preview(f.record.id, 2, { expectedConversionRevision: next.revision }), {
      code: /^(object-stale|query-stale|session-stale)$/,
    });
    await f.host().disposeInspections('new-owner');
    assert.deepEqual(await fs.readdir(path.join(f.dir, 'private')), []);
  }
);
