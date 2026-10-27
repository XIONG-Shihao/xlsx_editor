import { readPreference } from './preferences';
import {
  DEFAULT_VIEW_PREFERENCES,
  validateViewPreferencesPatch,
  type ViewPreferences,
  type ViewPreferenceStorage,
} from './viewPreferences';
export const VIEW_PRESETS_KEY = 'workbook.view-presets.v1';
export const BUILTIN_VIEW_PRESETS = Object.freeze([
  { id: 'default', name: 'Default', preferences: Object.freeze({ ...DEFAULT_VIEW_PREFERENCES }) },
  { id: 'data-entry', name: 'Data entry', preferences: Object.freeze({ ...DEFAULT_VIEW_PREFERENCES, zoom: 125 }) },
  {
    id: 'formula-audit',
    name: 'Formula audit',
    preferences: Object.freeze({ ...DEFAULT_VIEW_PREFERENCES, showFormulas: true }),
  },
]);
export interface NamedViewPreset {
  documentId: string;
  sheet: string;
  name: string;
  preferences: ViewPreferences;
}
const plain = (value: unknown): value is Record<string, unknown> =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
const key = (entry: Pick<NamedViewPreset, 'documentId' | 'sheet' | 'name'>) =>
  JSON.stringify([entry.documentId, entry.sheet, entry.name.toLowerCase()]);
function text(value: unknown, max: number): asserts value is string {
  if (typeof value !== 'string' || !value.trim() || value.length > max || /[\u0000-\u001f\u007f]/.test(value))
    throw Error('Invalid local view identity or name.');
}
function scope(documentId: string, sheet: string): void {
  text(documentId, 512);
  text(sheet, 128);
}
function preferences(value: unknown): ViewPreferences {
  const patch = validateViewPreferencesPatch(value);
  if (Object.keys(patch).length !== 5) throw Error('Save all five current view settings.');
  return { ...DEFAULT_VIEW_PREFERENCES, ...patch };
}
function name(value: string): string {
  text(value, 64);
  const result = value.trim();
  if (BUILTIN_VIEW_PRESETS.some((p) => p.name.toLowerCase() === result.toLowerCase()))
    throw Error('Choose a name different from the built-in presets.');
  return result;
}
const bytes = (value: string) => new TextEncoder().encode(value).length;
/** Private per-file/sheet snapshots. No workbook/custom-view/filter model. */
export class ViewPresetStore {
  #entries: NamedViewPreset[] = [];
  #pending = new Map<string, NamedViewPreset | null>();
  #removed = new Set<string>();
  #storedKeys = new Set<string>();
  #storage: ViewPreferenceStorage | null;
  warning: string | null = null;
  constructor(storage: ViewPreferenceStorage | null) {
    this.#storage = storage;
    this.#refresh();
  }
  #refresh(): void {
    if (!this.#storage) {
      this.warning = 'Local views are available for this session only.';
      return;
    }
    try {
      const raw = readPreference(this.#storage, VIEW_PRESETS_KEY);
      let entries: NamedViewPreset[] = [];
      if (raw !== null) {
        if (bytes(raw) > 65536) throw Error('Local view storage limit.');
        const data: unknown = JSON.parse(raw);
        if (
          !plain(data) ||
          Object.keys(data).sort().join(',') !== 'entries,schema' ||
          data.schema !== 1 ||
          !Array.isArray(data.entries) ||
          data.entries.length > 32
        )
          throw Error('Invalid local views.');
        const seen = new Set<string>();
        for (const item of data.entries) {
          if (!plain(item) || Object.keys(item).sort().join(',') !== 'documentId,name,preferences,sheet')
            throw Error('Invalid local view record.');
          text(item.documentId, 512);
          text(item.sheet, 128);
          text(item.name, 64);
          const entry = {
              documentId: item.documentId,
              sheet: item.sheet,
              name: name(item.name),
              preferences: preferences(item.preferences),
            },
            id = key(entry);
          if (seen.has(id)) throw Error('Duplicate local view.');
          seen.add(id);
          entries.push(entry);
        }
      }
      this.#storedKeys = new Set(entries.map(key));
      entries = entries.filter((entry) => !this.#removed.has(entry.documentId));
      for (const [id, entry] of this.#pending) {
        entries = entries.filter((item) => key(item) !== id);
        if (entry) entries.push(structuredClone(entry));
      }
      while (entries.length > 32) entries.shift();
      this.#entries = entries;
      if (!this.#pending.size && !this.#removed.size) this.warning = null;
    } catch {
      this.warning = 'Saved local views could not be loaded; current session views are kept.';
    }
  }
  #save(): void {
    // Only persisted deleted keys need tombstones while a write is denied.
    for (const [id, item] of this.#pending) if (item === null && !this.#storedKeys.has(id)) this.#pending.delete(id);
    const serialized = JSON.stringify({ schema: 1, entries: this.#entries });
    if (bytes(serialized) > 65536) {
      this.warning = 'Local views exceed the 64 KiB storage limit; current views are kept for this session.';
      return;
    }
    if (!this.#storage) {
      this.warning = 'Local views are available for this session only.';
      return;
    }
    try {
      this.#storage.setItem(VIEW_PRESETS_KEY, serialized);
      this.#pending.clear();
      this.#removed.clear();
      this.warning = null;
    } catch {
      this.warning = 'Local views could not be saved; current views are kept for this session.';
    }
  }
  list(documentId: string, sheet: string): NamedViewPreset[] {
    scope(documentId, sheet);
    this.#refresh();
    return structuredClone(this.#entries.filter((entry) => entry.documentId === documentId && entry.sheet === sheet));
  }
  create(documentId: string, sheet: string, label: string, value: ViewPreferences): void {
    scope(documentId, sheet);
    const entry = { documentId, sheet, name: name(label), preferences: preferences(value) };
    this.#refresh();
    if (this.#entries.some((item) => key(item) === key(entry)))
      throw Error('A local view with that name already exists. Use Update saved view.');
    if (this.#entries.length >= 32)
      throw Error('The application profile already has 32 saved local views. Delete a saved view first.');
    this.#entries.push(entry);
    this.#pending.set(key(entry), entry);
    this.#removed.delete(documentId);
    this.#save();
  }
  update(documentId: string, sheet: string, label: string, value: ViewPreferences): void {
    const entries = this.list(documentId, sheet),
      prior = entries.find((entry) => entry.name === label);
    if (!prior) throw Error('Choose an existing saved local view.');
    const entry = { ...prior, preferences: preferences(value) };
    this.#entries = this.#entries.map((item) => (key(item) === key(entry) ? entry : item));
    this.#pending.set(key(entry), entry);
    this.#save();
  }
  delete(documentId: string, sheet: string, label: string): void {
    const prior = this.list(documentId, sheet).find((entry) => entry.name === label);
    if (!prior) throw Error('Choose an existing saved local view.');
    this.#entries = this.#entries.filter((item) => key(item) !== key(prior));
    this.#pending.set(key(prior), null);
    this.#save();
  }
  moveDocument(from: string, to: string): void {
    text(from, 512);
    text(to, 512);
    if (from === to) return;
    this.#refresh();
    const moved = this.#entries.filter((item) => item.documentId === from);
    if (!moved.length) return;
    this.#removed.delete(to);
    this.#removed.add(from);
    while (this.#removed.size > 32) this.#removed.delete(this.#removed.values().next().value!);
    for (const item of moved) {
      this.#pending.set(key(item), null);
      const entry = { ...item, documentId: to };
      this.#pending.set(key(entry), entry);
    }
    this.#entries = this.#entries.filter((item) => item.documentId !== from);
    for (const item of moved) {
      const entry = { ...item, documentId: to };
      this.#entries = this.#entries.filter((prior) => key(prior) !== key(entry));
      this.#entries.push(entry);
    }
    this.#refresh();
    while (this.#entries.length > 32) this.#entries.shift();
    this.#save();
  }
}
