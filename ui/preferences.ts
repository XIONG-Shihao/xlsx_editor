/** Current component preferences; retired profile keys are not imported. */
export function readPreference(storage: { getItem(key: string): string | null }, key: string): string | null {
  return storage.getItem(key);
}
