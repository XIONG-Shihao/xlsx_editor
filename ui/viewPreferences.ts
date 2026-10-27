import { readPreference } from './preferences';
/** Private UI preferences. These values are never workbook commands or caches. */
export interface ViewPreferences {
  zoom: number;
  gridlines: boolean;
  headings: boolean;
  formulaBar: boolean;
  showFormulas: boolean;
}
export type ViewPreferencesPatch = Partial<ViewPreferences>;
export const VIEW_PREFERENCES_KEY = 'workbook.view-preferences.v1';
export const DEFAULT_VIEW_PREFERENCES: Readonly<ViewPreferences> = Object.freeze({
  zoom: 100,
  gridlines: true,
  headings: true,
  formulaBar: true,
  showFormulas: false,
});
export interface ViewPreferenceStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}
export interface ViewPreferenceStoreOptions {
  key?: string;
  maxEntries?: number;
  maxBytes?: number;
}
interface Entry {
  documentId: string;
  sheet: string;
  preferences: ViewPreferencesPatch;
}
interface Pending {
  documentId: string;
  sheet: string;
  patch: ViewPreferencesPatch;
  replace?: boolean;
}
const fields = new Set(['zoom', 'gridlines', 'headings', 'formulaBar', 'showFormulas']);
const defaults = (): ViewPreferences => ({ ...DEFAULT_VIEW_PREFERENCES });
const bytes = (value: string): number => new TextEncoder().encode(value).length;
function record(value: unknown): value is Record<string, unknown> {
  return (
    !!value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null)
  );
}
function closed(value: unknown, keys: Set<string>): asserts value is Record<string, unknown> {
  if (!record(value) || Object.keys(value).some((key) => !keys.has(key)))
    throw Error('Invalid view preference members');
}
export function validateViewPreferencesPatch(value: unknown): ViewPreferencesPatch {
  closed(value, fields);
  const result: ViewPreferencesPatch = {};
  for (const key of Object.keys(value)) {
    if (key === 'zoom') {
      const zoom = value.zoom;
      if (typeof zoom !== 'number' || !Number.isInteger(zoom) || zoom < 10 || zoom > 400)
        throw Error('Zoom must be an integer from 10 to 400');
      result.zoom = zoom;
    } else {
      if (typeof value[key] !== 'boolean') throw Error('View display preferences must be boolean');
      result[key as Exclude<keyof ViewPreferences, 'zoom'>] = value[key] as boolean;
    }
  }
  return result;
}
function full(value: unknown): ViewPreferences {
  const patch = validateViewPreferencesPatch(value);
  if (Object.keys(patch).length !== fields.size) throw Error('Incomplete view preferences');
  return { ...DEFAULT_VIEW_PREFERENCES, ...patch };
}
function identity(value: unknown, maximum: number): asserts value is string {
  if (typeof value !== 'string' || !value.trim() || value.length > maximum || /[\u0000-\u001f\u007f]/.test(value))
    throw Error('Invalid stable view identity');
}
function limit(value: number | undefined, fallback: number, maximum: number): number {
  const selected = value ?? fallback;
  if (!Number.isInteger(selected) || selected < 1 || selected > maximum) throw Error('Invalid view preference quota');
  return selected;
}

export class ViewPreferenceStore {
  readonly key: string;
  readonly maxEntries: number;
  readonly maxBytes: number;
  #entries: Entry[] = [];
  #ephemeral = new Map<string, ViewPreferencesPatch>();
  #pending = new Map<string, Pending>();
  #removed = new Set<string>();
  #removedSheets = new Set<string>();
  #storage: ViewPreferenceStorage | null;
  #warning: string | null = null;
  get warning(): string | null {
    return this.#warning;
  }
  constructor(storage?: ViewPreferenceStorage | null, options: ViewPreferenceStoreOptions = {}) {
    this.key = options.key ?? VIEW_PREFERENCES_KEY;
    identity(this.key, 256);
    this.maxEntries = limit(options.maxEntries, 64, 64);
    this.maxBytes = limit(options.maxBytes, 64 * 1024, 64 * 1024);
    if (storage === undefined) {
      try {
        this.#storage = globalThis.localStorage ?? null;
      } catch {
        this.#storage = null;
      }
    } else this.#storage = storage;
    if (!this.#storage) {
      this.#warning = 'View preferences are local to this session.';
      return;
    }
    this.#refresh();
  }
  #read(): Entry[] {
    const raw = readPreference(this.#storage!, this.key);
    if (raw === null) return [];
    if (raw.length > this.maxBytes || bytes(raw) > this.maxBytes)
      throw Error('View preferences exceed their storage limit');
    const parsed: unknown = JSON.parse(raw);
    closed(parsed, new Set(['schema', 'entries']));
    if (
      ![1, 2].includes(parsed.schema as number) ||
      !Array.isArray(parsed.entries) ||
      parsed.entries.length > this.maxEntries
    )
      throw Error('Invalid view preference schema');
    const entries: Entry[] = [],
      seen = new Set<string>();
    for (const item of parsed.entries) {
      closed(item, new Set(['documentId', 'sheet', 'preferences']));
      identity(item.documentId, 512);
      identity(item.sheet, 128);
      const id = JSON.stringify([item.documentId, item.sheet]);
      if (seen.has(id)) throw Error('Duplicate view preference identity');
      seen.add(id);
      entries.push({
        documentId: item.documentId,
        sheet: item.sheet,
        preferences: parsed.schema === 1 ? full(item.preferences) : validateViewPreferencesPatch(item.preferences),
      });
    }
    return entries;
  }
  #refresh(): void {
    if (!this.#storage) return;
    try {
      let entries = this.#read().filter(
        (entry) =>
          !this.#removed.has(entry.documentId) &&
          !this.#removedSheets.has(JSON.stringify([entry.documentId, entry.sheet]))
      );
      for (const pending of this.#pending.values()) {
        const prior = entries.find((entry) => entry.documentId === pending.documentId && entry.sheet === pending.sheet);
        entries = entries.filter((entry) => entry !== prior);
        entries.push({
          documentId: pending.documentId,
          sheet: pending.sheet,
          preferences: { ...(pending.replace ? {} : (prior?.preferences ?? {})), ...pending.patch },
        });
      }
      while (entries.length > this.maxEntries) entries.shift();
      this.#entries = entries;
      if (!this.#pending.size && !this.#removed.size) this.#warning = null;
    } catch {
      this.#warning = 'Saved view preferences could not be loaded; the local view is used.';
    }
  }
  get(documentId: string | null | undefined, sheet: string, source: ViewPreferencesPatch = {}): ViewPreferences {
    identity(sheet, 128);
    const base = { ...DEFAULT_VIEW_PREFERENCES, ...validateViewPreferencesPatch(source) };
    if (documentId === null || documentId === undefined || documentId === '')
      return { ...base, ...this.#ephemeral.get(sheet) };
    identity(documentId, 512);
    this.#refresh();
    return {
      ...base,
      ...this.#entries.find((entry) => entry.documentId === documentId && entry.sheet === sheet)?.preferences,
    };
  }
  patch(
    documentId: string | null | undefined,
    sheet: string,
    value: ViewPreferencesPatch,
    source: ViewPreferencesPatch = {}
  ): ViewPreferences {
    const patch = validateViewPreferencesPatch(value);
    identity(sheet, 128);
    const next = { ...this.get(documentId, sheet, source), ...patch };
    if (documentId === null || documentId === undefined || documentId === '') {
      const overrides = { ...this.#ephemeral.get(sheet), ...patch };
      this.#ephemeral.delete(sheet);
      this.#ephemeral.set(sheet, overrides);
      while (this.#ephemeral.size > this.maxEntries) this.#ephemeral.delete(this.#ephemeral.keys().next().value!);
      this.#warning = 'Stable document identity is unavailable; view preferences remain local.';
      return { ...next };
    }
    identity(documentId, 512);
    this.#removed.delete(documentId);
    this.#removedSheets.delete(JSON.stringify([documentId, sheet]));
    const key = JSON.stringify([documentId, sheet]),
      prior = this.#pending.get(key);
    this.#pending.set(key, { ...prior, documentId, sheet, patch: { ...prior?.patch, ...patch } });
    const entries = this.#entries.filter((entry) => entry.documentId !== documentId || entry.sheet !== sheet);
    entries.push({
      documentId,
      sheet,
      preferences: {
        ...this.#entries.find((entry) => entry.documentId === documentId && entry.sheet === sheet)?.preferences,
        ...patch,
      },
    });
    while (entries.length > this.maxEntries) entries.shift();
    let serialized = this.#serialize(entries);
    while (bytes(serialized) > this.maxBytes && entries.length > 1) {
      entries.shift();
      serialized = this.#serialize(entries);
    }
    // A deliberately tiny custom quota can reject persistence without losing the
    // current instance's local choice or writing an oversized profile record.
    this.#entries = entries;
    const retained = new Set(entries.map((entry) => JSON.stringify([entry.documentId, entry.sheet])));
    for (const id of this.#pending.keys()) if (!retained.has(id)) this.#pending.delete(id);
    if (bytes(serialized) > this.maxBytes) {
      this.#warning = 'View preferences could not be saved within their storage limit.';
      return { ...next };
    }
    this.#save(serialized);
    return { ...next };
  }
  reset(documentId: string | null | undefined, sheet: string): void {
    identity(sheet, 128);
    if (!documentId) {
      this.#ephemeral.delete(sheet);
      return;
    }
    identity(documentId, 512);
    this.#refresh();
    const key = JSON.stringify([documentId, sheet]);
    this.#pending.delete(key);
    this.#removedSheets.add(key);
    while (this.#removedSheets.size > this.maxEntries)
      this.#removedSheets.delete(this.#removedSheets.values().next().value!);
    this.#entries = this.#entries.filter((entry) => entry.documentId !== documentId || entry.sheet !== sheet);
    this.#save(this.#serialize(this.#entries));
  }
  removeDocument(documentId: string): boolean {
    identity(documentId, 512);
    this.#refresh();
    const entries = this.#entries.filter((entry) => entry.documentId !== documentId);
    if (entries.length === this.#entries.length) return false;
    this.#entries = entries;
    this.#removed.add(documentId);
    while (this.#removed.size > this.maxEntries) this.#removed.delete(this.#removed.values().next().value!);
    for (const [id, pending] of this.#pending) if (pending.documentId === documentId) this.#pending.delete(id);
    this.#save(this.#serialize(entries));
    return true;
  }
  /** First Save changes a private new-document ID into the stable file binding.
   * Source sheet preferences win collisions; unrelated target sheets survive.
   * Queue the complete transaction before one write, so a denied/quota write
   * cannot restore old IDs on the next shared-profile refresh. */
  moveDocument(from: string, to: string): boolean {
    identity(from, 512);
    identity(to, 512);
    if (from === to) return false;
    this.#refresh();
    const moved = this.#entries.filter((entry) => entry.documentId === from);
    if (!moved.length) return false;
    const sheets = new Set(moved.map((entry) => entry.sheet));
    const entries = this.#entries.filter(
      (entry) => entry.documentId !== from && (entry.documentId !== to || !sheets.has(entry.sheet))
    );
    this.#removed.delete(to);
    this.#removed.add(from);
    while (this.#removed.size > this.maxEntries) this.#removed.delete(this.#removed.values().next().value!);
    for (const [id, pending] of this.#pending) if (pending.documentId === from) this.#pending.delete(id);
    for (const entry of moved) {
      const preferences = { ...entry.preferences },
        key = JSON.stringify([to, entry.sheet]);
      entries.push({ documentId: to, sheet: entry.sheet, preferences });
      this.#pending.set(key, { documentId: to, sheet: entry.sheet, patch: preferences, replace: true });
    }
    this.#entries = entries;
    this.#save(this.#serialize(entries));
    return true;
  }
  #serialize(entries: Entry[]): string {
    return JSON.stringify({ schema: 2, entries });
  }
  #save(serialized: string): void {
    if (bytes(serialized) > this.maxBytes) {
      this.#warning = 'View preferences could not be saved within their storage limit.';
      return;
    }
    if (!this.#storage) {
      this.#warning = 'View preferences are local to this session.';
      return;
    }
    try {
      this.#storage.setItem(this.key, serialized);
      this.#pending.clear();
      this.#removed.clear();
      this.#removedSheets.clear();
      this.#warning = null;
    } catch {
      this.#warning = 'View preferences could not be saved; the current view is kept for this session.';
    }
  }
}
export function createViewPreferenceStore(
  storage?: ViewPreferenceStorage | null,
  options: ViewPreferenceStoreOptions = {}
): ViewPreferenceStore {
  return new ViewPreferenceStore(storage, options);
}
