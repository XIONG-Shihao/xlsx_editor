import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { MessageChannel } from 'node:worker_threads';
import { once } from 'node:events';
import { allocateWorkerStorage, releaseWorkerStorage } from '../adapters/worker-storage.mjs';
import { createNativeRelay, knownReapedNativePid } from '../adapters/native-relay.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';

async function owner(t) {
  const storage = allocateWorkerStorage(),
    index = path.join(storage.root, 'owner-proof'),
    base = path.join(index, 'base');
  await fs.mkdir(base, { recursive: true });
  t.after(() => fs.rm(storage.root, { recursive: true, force: true }));
  return { storage, index };
}
async function reap(storage) {
  const { port1, port2 } = new MessageChannel(),
    relay = createNativeRelay(port1, nativeWorkbookBinary(), { storageRoot: storage.root });
  try {
    const body = Buffer.from(JSON.stringify({ id: 1, input: { action: 'unknown-proof-test' } })),
      prefix = Buffer.alloc(4);
    prefix.writeUInt32BE(body.length);
    const reply = once(port2, 'message');
    port2.postMessage({ bytes: Uint8Array.from(Buffer.concat([prefix, body])) });
    await reply;
    const proof = await relay.close();
    assert(Number.isSafeInteger(proof.reapedNativePid));
    assert(proof.reapedNativePid > 0);
    assert(Object.isFrozen(proof));
    return proof;
  } finally {
    port2.close();
    await relay.close();
  }
}

test('only the genuinely reaped helper proof bound to this private root reclaims a torn activity lease', async (t) => {
  const { storage, index } = await owner(t),
    proof = await reap(storage);
  await fs.writeFile(path.join(index, 'native-active.json'), '');
  assert.equal(knownReapedNativePid(proof, storage.root), proof.reapedNativePid);
  assert.equal(knownReapedNativePid({ ...proof }, storage.root), undefined);
  assert.equal(knownReapedNativePid(proof, storage.root + '-other'), undefined);
  await releaseWorkerStorage(storage, { reapedNativeProof: proof });
  await assert.rejects(fs.lstat(storage.root), { code: 'ENOENT' });
});

test('forged, wrong-root and dead-collector proof claims retain the owned scratch', async (t) => {
  const first = await owner(t),
    second = await owner(t),
    proof = await reap(first.storage);
  await fs.writeFile(path.join(second.index, 'native-active.json'), '');
  await assert.rejects(
    releaseWorkerStorage(second.storage, { reapedNativeProof: { reapedNativePid: proof.reapedNativePid } }),
    { code: 'internal-invalid-field' }
  );
  await assert.rejects(releaseWorkerStorage(second.storage, { reapedNativeProof: proof }), {
    code: 'internal-invalid-field',
  });
  await assert.rejects(releaseWorkerStorage(first.storage, { deadProcess: true, reapedNativeProof: proof }), {
    code: 'internal-invalid-field',
  });
  assert((await fs.lstat(second.storage.root)).isDirectory());
});

test('genuine reaping proof cannot override a different valid native owner or a symlink activity lease', async (t) => {
  const { storage, index } = await owner(t),
    proof = await reap(storage),
    lease = path.join(index, 'native-active.json');
  await fs.writeFile(lease, JSON.stringify({ pid: process.pid }));
  await assert.rejects(releaseWorkerStorage(storage, { reapedNativeProof: proof }), /does not match/);
  assert((await fs.lstat(storage.root)).isDirectory());
  await fs.rm(lease);
  const outside = path.join(storage.root, 'outside.json');
  await fs.writeFile(outside, '');
  await fs.symlink(outside, lease);
  await assert.rejects(releaseWorkerStorage(storage, { reapedNativeProof: proof }), {
    code: /^(internal-invalid-field|session-unsupported)$/,
  });
  assert.equal(await fs.readFile(outside, 'utf8'), '');
});

test('root token and confinement remain authoritative even with a genuine reaped helper', async (t) => {
  const { storage, index } = await owner(t),
    proof = await reap(storage);
  await fs.writeFile(path.join(index, 'native-active.json'), '');
  await assert.rejects(releaseWorkerStorage({ ...storage, token: 'wrong' }, { reapedNativeProof: proof }), {
    code: 'session-refused',
  });
  await assert.rejects(
    releaseWorkerStorage({ ...storage, root: path.dirname(storage.root) }, { reapedNativeProof: proof }),
    /reaping proof|storage/
  );
  assert((await fs.lstat(storage.root)).isDirectory());
});
