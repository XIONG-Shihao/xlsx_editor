import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash, randomUUID } from 'node:crypto';
import { WorkbookSpellingJobs } from '../adapters/workbook-spelling-jobs.mjs';
import {
  createOfflineSpellingProvider,
  OFFLINE_SPELLING_PROVIDER_ID,
} from '@office-editors/xlsx-editor/offline-spelling';

const gate = () => {
  let release, entered;
  return {
    done: new Promise((resolve) => (release = resolve)),
    started: new Promise((resolve) => (entered = resolve)),
    release: () => release(),
    enter: () => entered(),
  };
};
const binding = () => ({
  documentId: randomUUID(),
  ownerId: randomUUID(),
  state: 'accepted-head',
  revision: 'accepted-revision',
});
function fixture(extra = {}) {
  const original = binding(),
    live = { ...original },
    calls = { read: 0, verify: 0, write: 0 };
  let granted = true;
  const access = {
    capture: (request) =>
      request.operation === 'spellingCancel'
        ? Object.fromEntries(Object.keys(original).map((k) => [k, request[k]]))
        : { ...live },
    authorize: async () => {
      calls.read++;
      if (!granted) throw Error('Read grant revoked');
    },
    verify: async (request) => {
      calls.verify++;
      assert.equal(request.ownerId, live.ownerId);
      assert.equal(request.documentId, live.documentId);
      if (request.operation !== 'spellingCancel')
        assert.deepEqual(
          Object.fromEntries(Object.keys(original).map((k) => [k, request[k]])),
          live,
          'Original accepted head changed'
        );
    },
  };
  const jobs = new WorkbookSpellingJobs({ ...access, ...extra }),
    request = (operation, fields = {}) => ({
      operation,
      ...original,
      requestId: randomUUID(),
      ...(operation === 'spellingCheck' ? { locale: 'en-US', words: ['hello', 'helo', 'wrld'] } : {}),
      ...fields,
    });
  return { jobs, request, original, live, calls, revoke: () => (granted = false), access };
}
async function check(name, fn) {
  await test(name, async () => {
    await fn();
  });
}

await check('Real dictionaries bind availability and dialect results to one readonly accepted owner', async () => {
  const f = fixture();
  try {
    const availability = await f.jobs.run(f.request('spellingAvailability'));
    assert.equal(availability.spellingAvailability.availability.providerId, OFFLINE_SPELLING_PROVIDER_ID);
    assert(availability.spellingAvailability.availability.locales.every((row) => row.status === 'available'));
    const us = await f.jobs.run(f.request('spellingCheck', { words: ['hello', 'color', 'colour', 'helo'] })),
      gb = await f.jobs.run(
        f.request('spellingCheck', { locale: 'en-GB', words: ['hello', 'color', 'colour', 'helo'] })
      );
    assert.deepEqual(us.spellingCheck.receipt, f.original);
    assert.deepEqual(gb.spellingCheck.receipt, f.original);
    assert.deepEqual(
      us.spellingCheck.result.results.map((row) => row.correct),
      [true, true, false, false]
    );
    assert.deepEqual(
      gb.spellingCheck.result.results.map((row) => row.correct),
      [true, false, true, false]
    );
    assert(us.spellingCheck.result.results[3].suggestions.includes('hello'));
    assert.notEqual(us.spellingCheck.result.dictionaryId, gb.spellingCheck.result.dictionaryId);
    assert.equal(f.calls.write, 0);
    assert.deepEqual(f.live, f.original);
    const desktop = await fs.mkdtemp(path.join(os.tmpdir(), 'spelling-desktop-layout-'));
    let copied;
    try {
      await fs.mkdir(path.join(desktop, 'adapters'));
      await fs.mkdir(path.join(desktop, 'node_modules'));
      await fs.writeFile(path.join(desktop, 'package.json'), JSON.stringify({ name: 'xlsx-desktop', type: 'module' }));
      for (const name of [
        'workbook-spelling-jobs.mjs',
        'offline-spelling.mjs',
        'offline-spelling-assets.mjs',
        'offline-spelling-worker.mjs',
      ])
        await fs.copyFile(
          fileURLToPath(new URL('../adapters/' + name, import.meta.url)),
          path.join(desktop, 'adapters', name)
        );
      for (const name of ['nspell', 'is-buffer', 'dictionary-en', 'dictionary-en-gb'])
        await fs.symlink(
          fileURLToPath(new URL('../node_modules/' + name, import.meta.url)),
          path.join(desktop, 'node_modules', name),
          'dir'
        );
      const { WorkbookSpellingJobs: CopiedJobs } = await import(
        pathToFileURL(path.join(desktop, 'adapters/workbook-spelling-jobs.mjs')).href
      );
      copied = new CopiedJobs(f.access);
      const reply = await copied.run(f.request('spellingCheck', { words: ['helo'] }));
      assert.equal(reply.spellingCheck.result.status, 'completed');
      assert(reply.spellingCheck.result.results[0].suggestions.includes('hello'));
    } finally {
      await copied?.close();
      await fs.rm(desktop, { recursive: true, force: true });
    }
    return {
      availability,
      us,
      gb,
      authorization: 'read only; no edit grant requested',
      desktopLayout:
        'Actual copied adapters + trusted installed dictionaries in a package named xlsx-desktop with no exports/self alias; no packaged ASAR/app claim.',
    };
  } finally {
    await f.jobs.close();
  }
});
await check('Closed request fields and full word byte limits refuse before provider creation', async () => {
  let created = 0;
  const f = fixture({
    providerFactory: () => {
      created++;
      return createOfflineSpellingProvider();
    },
  });
  try {
    for (const request of [
      f.request('spellingCheck', { assetPath: '/private/a.dic' }),
      f.request('spellingCheck', { words: [] }),
      f.request('spellingCheck', { words: Array(101).fill('hello') }),
      f.request('spellingCheck', { words: ['a'.repeat(257)] }),
      f.request('spellingCheck', { words: Array(100).fill('😀'.repeat(256)) }),
      f.request('spellingCheck', { requestId: 'ABC' }),
      f.request('spellingCheck', { kind: 'indexed', sessionId: randomUUID() }),
      f.request('spellingAvailability', { locale: 'en-US' }),
      f.request('spellingCancel', { words: ['hello'] }),
    ])
      assert.throws(() => f.jobs.run(request), /Invalid|bounded|32 KiB|scoped/);
    assert.equal(created, 0);
    assert.equal(f.calls.read, 0);
    const id = f.request('spellingCheck');
    await f.jobs.run(id);
    assert.throws(() => f.jobs.run(id), /cannot be reused/);
    const unsupported = await f.jobs.run(f.request('spellingCheck', { locale: 'fr-FR' }));
    assert.equal(unsupported.spellingCheck.result.status, 'unavailable');
    assert.deepEqual(unsupported.spellingCheck.result.results, []);
    assert(!Object.hasOwn(unsupported.spellingCheck.result, 'dictionaryId'));
    return { created, unsupported };
  } finally {
    await f.jobs.close();
  }
});
await check('Real completed results are suppressed after head change or read-grant loss', async () => {
  const outcomes = [];
  for (const kind of ['head', 'grant']) {
    const held = gate(),
      f = fixture({
        providerFactory: () => {
          const provider = createOfflineSpellingProvider();
          return {
            ...provider,
            check: async (request) => {
              const result = await provider.check(request);
              held.enter();
              await held.done;
              return result;
            },
          };
        },
      });
    try {
      const pending = f.jobs.run(f.request('spellingCheck'));
      pending.catch(() => {});
      await held.started;
      if (kind === 'head') f.live.state = 'new-head';
      else f.revoke();
      held.release();
      await assert.rejects(pending, kind === 'head' ? /head changed/ : /grant revoked/);
      assert.equal(f.calls.write, 0);
      outcomes.push(kind);
    } finally {
      held.release();
      await f.jobs.close();
    }
  }
  return { outcomes, scope: 'Delayed genuine provider reply; owner/grant transition is controlled.' };
});
await check('Queued cancel preserves lifetime identity and cannot cancel a different receipt', async () => {
  const held = gate(),
    f = fixture(),
    access = {
      ...f.access,
      authorize: async (request, receipt) => {
        await f.access.authorize(request, receipt);
        if (request.operation === 'spellingAvailability') {
          held.enter();
          await held.done;
        }
      },
    };
  try {
    const first = f.jobs.run(f.request('spellingAvailability'), access);
    first.catch(() => {});
    await held.started;
    const request = f.request('spellingCheck'),
      pending = f.jobs.run(request);
    pending.catch(() => {});
    const third = f.jobs.run(f.request('spellingCheck', { words: ['hello'] })),
      fourth = f.jobs.run(f.request('spellingCheck', { words: ['world'] }));
    assert.throws(() => f.jobs.run(f.request('spellingCheck')), /Too many queued/);
    await assert.rejects(
      f.jobs.cancel({ ...request, operation: 'spellingCancel', words: undefined, locale: undefined }),
      /closed/
    );
    const cancel = { operation: 'spellingCancel', ...f.original, requestId: request.requestId };
    await assert.rejects(f.jobs.cancel({ ...cancel, revision: 'foreign-revision' }), /another job owner/);
    const ack = await f.jobs.cancel(cancel);
    assert.equal(ack.spellingCancellation.cancelled, true);
    held.release();
    await first;
    assert.equal((await pending).spellingCheck.result.status, 'cancelled');
    for (const reply of await Promise.all([third, fourth]))
      assert.equal(reply.spellingCheck.result.status, 'completed');
    assert.throws(() => f.jobs.run(request), /cannot be reused/);
    assert.equal((await f.jobs.cancel(cancel)).spellingCancellation.cancelled, false);
    return { ack, noReuse: true, queueBound: 4 };
  } finally {
    held.release();
    await f.jobs.close();
  }
});
await check('Real worker cancellation and owner close await disposal without retargeting following work', async () => {
  let disposed = 0,
    observed;
  const f = fixture({
    providerFactory: () => {
      const provider = createOfflineSpellingProvider();
      return {
        ...provider,
        dispose: async () => {
          disposed++;
          await provider.dispose();
        },
      };
    },
  });
  try {
    await f.jobs.run(f.request('spellingAvailability'));
    const request = f.request('spellingCheck', { words: Array(100).fill('q'.repeat(120)) }),
      pending = f.jobs.run(request);
    pending.catch(() => {});
    await new Promise(setImmediate);
    f.live.state = 'new-head';
    const ack = await f.jobs.cancel({ operation: 'spellingCancel', ...f.original, requestId: request.requestId });
    assert.equal(ack.spellingCancellation.cancelled, true);
    await assert.rejects(pending, /head changed/);
    Object.assign(f.live, f.original);
    const next = await f.jobs.run(f.request('spellingCheck', { words: ['hello'] }));
    assert.equal(next.spellingCheck.result.status, 'completed');
    assert.equal(next.spellingCheck.result.results[0].correct, true);
    const closing = f.jobs.run(f.request('spellingCheck', { words: Array(100).fill('q'.repeat(120)) }));
    closing.catch(() => {});
    await new Promise(setImmediate);
    await f.jobs.close();
    await assert.rejects(closing, /retired/);
    assert(disposed > 0);
    assert.throws(() => f.jobs.run(f.request('spellingCheck')), /retired/);
    observed = { ack, next, disposed, scope: 'Actual provider worker cancellation; controlled owner head retirement.' };
  } finally {
    await f.jobs.close();
  }
  const held = gate(),
    g = fixture({
      providerFactory: () => {
        const provider = createOfflineSpellingProvider();
        return {
          ...provider,
          cancel: async (id) => {
            await provider.cancel(id);
            held.enter();
            await held.done;
          },
        };
      },
    });
  try {
    await g.jobs.run(g.request('spellingAvailability'));
    const req = g.request('spellingCheck', { words: Array(100).fill('q'.repeat(120)) }),
      pending = g.jobs.run(req);
    pending.catch(() => {});
    await new Promise(setImmediate);
    const stopping = g.jobs.cancel({ operation: 'spellingCancel', ...g.original, requestId: req.requestId });
    stopping.catch(() => {});
    await held.started;
    g.revoke();
    held.release();
    await assert.rejects(stopping, /grant revoked/);
    const result = await Promise.allSettled([pending]);
    if (result[0].status === 'fulfilled') assert.equal(result[0].value.spellingCheck.result.status, 'cancelled');
    else assert.match(result[0].reason.message, /grant revoked/);
  } finally {
    held.release();
    await g.jobs.close();
  }
  const available = gate(),
    replyHold = gate(),
    disposal = gate(),
    h = fixture({
      providerFactory: () => {
        const provider = createOfflineSpellingProvider();
        return {
          ...provider,
          availability: async () => {
            const result = await provider.availability();
            available.enter();
            await replyHold.done;
            return result;
          },
          dispose: async () => {
            await provider.dispose();
            disposal.enter();
            await disposal.done;
          },
        };
      },
    });
  try {
    const req = h.request('spellingAvailability'),
      pending = h.jobs.run(req);
    pending.catch(() => {});
    await available.started;
    const stopping = h.jobs.cancel({ operation: 'spellingCancel', ...h.original, requestId: req.requestId });
    stopping.catch(() => {});
    await disposal.started;
    let closed = false;
    const closing = h.jobs.close().then(() => {
      closed = true;
    });
    replyHold.release();
    await new Promise(setImmediate);
    assert.equal(closed, false, 'Close must await cancellation-owned provider disposal');
    disposal.release();
    await closing;
    await assert.rejects(pending, /retired/);
    await assert.rejects(stopping, /retired/);
  } finally {
    replyHold.release();
    disposal.release();
    await h.jobs.close();
  }
  return {
    ...observed,
    lateCancellationReadGrant: 'Actual cancel completed; held acknowledgment refused after controlled read revocation.',
    closeWaitsCancellationDisposal: true,
  };
});
