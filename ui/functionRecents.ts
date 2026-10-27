/** Private profile preference: catalog IDs only, recorded after owner acceptance. */
import { FUNCTION_CATALOG, type FunctionId } from './functionCatalog';

export interface FunctionRecentStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}
export const FUNCTION_RECENTS_KEY = 'xlsx.function-recents.v1';
export const FUNCTION_RECENTS_LIMIT = 12;
const allowed = new Set<string>(FUNCTION_CATALOG.map((item) => item.id));
function decode(raw: string | null): FunctionId[] {
  if (raw === null || raw.length > 4096) return [];
  try {
    const parsed: unknown = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return [];
    const record = parsed as { version?: unknown; ids?: unknown };
    if (record.version !== 1 || !Array.isArray(record.ids) || record.ids.length > 100) return [];
    return [...new Set(record.ids.filter((id): id is FunctionId => typeof id === 'string' && allowed.has(id)))].slice(
      0,
      FUNCTION_RECENTS_LIMIT
    );
  } catch {
    return [];
  }
}
export class FunctionRecentStore {
  #storage: FunctionRecentStorage | null;
  #lastSeen: string | null | undefined;
  #ids: FunctionId[] = [];
  #dirty = false;
  constructor(storage?: FunctionRecentStorage | null) {
    if (storage === undefined) {
      try {
        this.#storage = globalThis.localStorage ?? null;
      } catch {
        this.#storage = null;
      }
    } else this.#storage = storage;
  }
  read(): readonly FunctionId[] {
    try {
      const raw = this.#storage?.getItem(FUNCTION_RECENTS_KEY);
      if (raw !== undefined && raw !== this.#lastSeen) {
        this.#lastSeen = raw;
        const stored = decode(raw);
        this.#ids = this.#dirty ? [...new Set([...this.#ids, ...stored])].slice(0, FUNCTION_RECENTS_LIMIT) : stored;
      }
    } catch {
      /* Unavailable profile storage keeps the bounded in-memory preference. */
    }
    return Object.freeze([...this.#ids]);
  }
  /** Call only for an actual accepted insertion, including an accepted late reply. */
  recordAccepted(id: FunctionId): readonly FunctionId[] {
    if (typeof id !== 'string' || !allowed.has(id)) throw Error('Unknown recent function');
    const previous = this.read();
    this.#ids = [id, ...previous.filter((item) => item !== id)].slice(0, FUNCTION_RECENTS_LIMIT);
    this.#dirty = true;
    const raw = JSON.stringify({ version: 1, ids: this.#ids });
    try {
      this.#storage?.setItem(FUNCTION_RECENTS_KEY, raw);
      if (this.#storage) {
        this.#lastSeen = raw;
        this.#dirty = false;
      }
    } catch {
      /* The accepted workbook result does not depend on preference persistence. */
    }
    return Object.freeze([...this.#ids]);
  }
}
export function createFunctionRecentStore(storage?: FunctionRecentStorage | null): FunctionRecentStore {
  return new FunctionRecentStore(storage);
}
