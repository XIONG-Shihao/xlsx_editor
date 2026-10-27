// Byte identity of a file, for checking that a test left its input unchanged.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';

export const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');

export async function record(file) {
  const bytes = await fs.readFile(file);
  return { path: path.resolve(file), sha256: digest(bytes), bytes: bytes.length };
}
