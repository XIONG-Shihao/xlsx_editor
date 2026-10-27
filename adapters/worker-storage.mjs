import { FORMAT } from '../core/format.mjs';
/** Host-owned disposable storage; durable recovery is deliberately elsewhere. */
import fs from 'node:fs/promises';
import { mkdirSync, lstatSync, mkdtempSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { knownReapedNativePid } from './native-relay.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const directory = path.join(
  tmpdir(),
  `spreadsheet-workers-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
);
const schema = FORMAT.SCHEMA_WORKER_STORAGE_V1;
const delay = () => new Promise((resolve) => setTimeout(resolve, 10));
function safeDirectory(stat) {
  return stat.isDirectory() && !stat.isSymbolicLink() && (!process.getuid || stat.uid === process.getuid());
}
export function allocateWorkerStorage() {
  mkdirSync(directory, { recursive: true, mode: 0o700 });
  if (!safeDirectory(lstatSync(directory))) throw refusal('session-unsupported', 'Unsafe spreadsheet worker storage');
  const root = mkdtempSync(path.join(directory, 'worker-')),
    token = randomUUID();
  writeFileSync(path.join(root, 'worker-owner.json'), JSON.stringify({ schema, pid: process.pid, token }), {
    mode: 0o600,
    flag: 'wx',
  });
  return { root, token };
}
async function owner(root) {
  if (path.dirname(root) !== directory || !/^worker-[a-zA-Z0-9]+$/.test(path.basename(root)))
    throw refusal('internal-invalid-field', 'Invalid spreadsheet worker storage');
  if (!safeDirectory(await fs.lstat(root))) throw refusal('session-unsupported', 'Unsafe spreadsheet worker storage');
  const file = path.join(root, 'worker-owner.json'),
    stat = await fs.lstat(file);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 4096)
    throw refusal('internal-invalid-field', 'Invalid spreadsheet worker lease');
  const value = JSON.parse(await fs.readFile(file, 'utf8'));
  if (value.schema !== schema || !Number.isSafeInteger(value.pid) || value.pid < 1 || typeof value.token !== 'string')
    throw refusal('internal-invalid-field', 'Invalid spreadsheet worker lease');
  return value;
}
function alive(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    if (error.code === 'ESRCH') return false;
    throw error;
  }
}
async function stopIndex(root, reapedNativePid) {
  // The JS worker must have exited before this cleanup; no new index roots can
  // be created. Its native child may still be processing an already sent frame.
  for (const entry of await fs.readdir(root, { withFileTypes: true }))
    if (entry.isDirectory()) {
      const child = path.join(root, entry.name);
      if (!safeDirectory(await fs.lstat(child))) throw refusal('session-unsupported', 'Unsafe private index directory');
      await fs.writeFile(path.join(child, 'cancelled'), '1', { flag: 'wx' }).catch((error) => {
        if (error.code !== 'EEXIST') throw error;
      });
    }
  const lease = path.join(root, 'native-active.json');
  for (let attempt = 0; attempt <= 500; attempt++) {
    let active;
    try {
      const stat = await fs.lstat(lease);
      if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 4096)
        throw refusal('internal-invalid-field', 'Invalid native activity lease');
      active = JSON.parse(await fs.readFile(lease, 'utf8'));
    } catch (error) {
      if (error.code === 'ENOENT') return;
      if (error instanceof SyntaxError) {
        // A helper can die while truncating/writing this ephemeral activity file.
        // Only its live host's exact relay may supply the known reaped PID. The
        // generic dead-owner collector has no such proof and retains the root.
        if (reapedNativePid !== undefined) return;
        if (attempt < 500) {
          await delay();
          continue;
        }
      }
      throw error;
    }
    if (!Number.isSafeInteger(active.pid) || active.pid < 1)
      throw refusal('internal-invalid-field', 'Invalid native activity owner');
    if (active.pid === reapedNativePid) return;
    if (reapedNativePid !== undefined)
      throw refusal(
        'session-unsupported',
        'Native activity lease does not match the owned reaped helper; private indexes retained'
      );
    if (!alive(active.pid)) return;
    if (attempt === 500)
      throw refusal(
        'operation-cancelled',
        'Native work is still stopping; private indexes retained with cancellation markers'
      );
    await delay();
  }
}
export async function releaseWorkerStorage({ root, token }, { deadProcess = false, reapedNativeProof } = {}) {
  const reapedNativePid = reapedNativeProof === undefined ? undefined : knownReapedNativePid(reapedNativeProof, root);
  if (reapedNativeProof !== undefined && (reapedNativePid === undefined || deadProcess))
    throw refusal('internal-invalid-field', 'Invalid owned native reaping proof');
  let lease;
  try {
    lease = await owner(root);
  } catch (error) {
    if (error.code === 'ENOENT') return;
    throw error;
  }
  if (lease.token !== token || (!deadProcess && lease.pid !== process.pid) || (deadProcess && alive(lease.pid)))
    throw refusal('session-refused', 'Spreadsheet worker storage is still owned');
  for (const entry of await fs.readdir(root, { withFileTypes: true }))
    if (entry.isDirectory()) {
      if (!/^owner-[a-zA-Z0-9]+$/.test(entry.name))
        throw refusal('session-malformed', 'Unexpected private index directory');
      const index = path.join(root, entry.name);
      if (!safeDirectory(await fs.lstat(index))) throw refusal('session-unsupported', 'Unsafe private index directory');
      await stopIndex(index, reapedNativePid);
    }
  await fs.rm(root, { recursive: true, force: true });
}
export async function collectDeadWorkerStorage() {
  let entries;
  try {
    entries = await fs.readdir(directory, { withFileTypes: true });
  } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
  if (!safeDirectory(await fs.lstat(directory)))
    throw refusal('session-unsupported', 'Unsafe spreadsheet worker storage');
  const retained = [];
  for (const entry of entries)
    if (entry.isDirectory() && /^worker-[a-zA-Z0-9]+$/.test(entry.name)) {
      const root = path.join(directory, entry.name);
      try {
        const lease = await owner(root);
        if (!alive(lease.pid)) await releaseWorkerStorage({ root, token: lease.token }, { deadProcess: true });
      } catch (error) {
        if (error.code !== 'ENOENT') retained.push({ root, error: String(error.message ?? error) });
      }
    }
  return retained;
}
