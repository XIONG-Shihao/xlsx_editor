/** Live dictionary identity for controlled UI replies; no dependency on disposable historical evidence. */
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import { createOfflineSpellingProvider } from '../../adapters/offline-spelling.mjs';

export async function reviewProviderAvailability() {
  if (process.env.XLSX_REVIEW_PROVIDER_PROOF) {
    const proof = JSON.parse(await fs.readFile(process.env.XLSX_REVIEW_PROVIDER_PROOF, 'utf8'));
    return proof.cases[0].availability;
  }
  const provider = createOfflineSpellingProvider();
  try {
    const availability = await provider.availability();
    assert(
      availability.locales.some((locale) => locale.status === 'available' && locale.verified),
      'A verified local spelling dictionary is required'
    );
    return availability;
  } finally {
    await provider.dispose();
  }
}
