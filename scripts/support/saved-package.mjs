/** Independent saved-file identity and formula-cache expectations. */
import assert from 'node:assert/strict';
import { createReadStream } from 'node:fs';
import { createHash } from 'node:crypto';

export async function hashFile(file) {
  const digest = createHash('sha256');
  for await (const chunk of createReadStream(file)) digest.update(chunk);
  return digest.digest('hex');
}
export function formulaCacheExpectation(value) {
  assert(
    typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value)),
    'Configured formula caches must be finite numeric, text or boolean scalars'
  );
  if (typeof value !== 'number') return { value };
  const bytes = new Uint8Array(8);
  new DataView(bytes.buffer).setFloat64(0, value, false);
  return { value, binary64Bits: [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('') };
}
