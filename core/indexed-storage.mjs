import { FORMAT, matchesFormat } from './format.mjs';
/** Disposable native indexes are separate from durable recovery snapshots.
 * A dead process's lease permits cleanup after a crash; live owners are kept. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { refusal } from './refusal-codes.mjs';
// Linux /tmp is shared across accounts, unlike macOS's per-user temp directory.
const defaultDirectory = path.join(
  tmpdir(),
  `workbook-indexes-v1${process.platform === 'linux' ? `-${process.getuid()}` : ''}`
);
let prepared;
async function prepare(directory) {
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink() || (process.getuid && stat.uid !== process.getuid()))
    throw refusal('session-unsupported', 'Unsafe workbook cache directory');
  for (const item of await fs.readdir(directory, { withFileTypes: true })) {
    if (!item.isDirectory() || !/^owner-[a-zA-Z0-9]+$/.test(item.name)) continue;
    const root = path.join(directory, item.name);
    try {
      const leasePath = path.join(root, 'owner.json'),
        metadata = await fs.lstat(leasePath);
      if (!metadata.isFile() || metadata.isSymbolicLink() || metadata.size > 4096) continue;
      const lease = JSON.parse(await fs.readFile(leasePath, 'utf8'));
      if (!matchesFormat(lease.schema, 'SCHEMA_INDEX_OWNER_V1') || !Number.isSafeInteger(lease.pid) || lease.pid <= 0)
        continue;
      try {
        process.kill(lease.pid, 0);
      } catch (error) {
        if (error.code === 'ESRCH') await fs.rm(root, { recursive: true, force: true });
      }
    } catch (error) {
      if (error.code !== 'ENOENT' && !(error instanceof SyntaxError)) throw error;
    }
  }
}
export async function createIndexedStorage(sessionId, maxSessionBytes = 4 * 1024 ** 3, storageDirectory) {
  const directory = storageDirectory === undefined ? defaultDirectory : path.resolve(storageDirectory);
  if (!Number.isSafeInteger(maxSessionBytes) || maxSessionBytes < 128 * 1024 || maxSessionBytes > 4 * 1024 ** 3)
    throw refusal('internal-invalid-field', 'Invalid workbook session disk budget');
  if (storageDirectory === undefined)
    await (prepared ??= prepare(directory).catch((error) => {
      prepared = undefined;
      throw error;
    }));
  else await prepare(directory);
  const root = await fs.mkdtemp(path.join(directory, 'owner-'));
  try {
    await fs.writeFile(
      path.join(root, 'owner.json'),
      JSON.stringify({ schema: FORMAT.SCHEMA_INDEX_OWNER_V1, pid: process.pid, sessionId, maxSessionBytes }),
      { mode: 0o600, flag: 'wx' }
    );
    return root;
  } catch (error) {
    await fs.rm(root, { recursive: true, force: true });
    throw error;
  }
}
