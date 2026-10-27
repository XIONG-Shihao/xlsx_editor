import { readPreference } from './preferences';
import { inkPen, type InkPenId } from '../core/ink-stroke.mjs';
export const INK_PREFERENCE_KEY = 'workbook.ink-preset.v1';
type PrivateStorage = Pick<Storage, 'getItem' | 'setItem'>;
/** A single bounded last-used choice, separate from workbook state/history. */
export function createInkPreferenceStore(storage: PrivateStorage | null) {
  let chosen: InkPenId = 'black',
    sessionOnly = false,
    warning: string | null = null;
  return {
    get warning() {
      return warning;
    },
    read(): InkPenId {
      if (sessionOnly) return chosen;
      if (!storage) {
        warning = 'Pen choice is remembered for this session only.';
        return chosen;
      }
      try {
        const raw = readPreference(storage, INK_PREFERENCE_KEY);
        if (raw === null) {
          chosen = 'black';
          warning = null;
          return chosen;
        }
        if (raw.length > 256) throw Error('Oversized pen preference');
        const value = JSON.parse(raw);
        if (
          !value ||
          Array.isArray(value) ||
          Object.keys(value).sort().join(',') !== 'pen,schema' ||
          value.schema !== 1
        )
          throw Error('Invalid pen preference');
        chosen = inkPen(value.pen).id;
        warning = null;
      } catch {
        warning = 'Saved pen choice could not be read; using the current choice for this session.';
      }
      return chosen;
    },
    write(pen: InkPenId): InkPenId {
      chosen = inkPen(pen).id;
      if (!storage) {
        sessionOnly = true;
        warning = 'Pen choice is remembered for this session only.';
        return chosen;
      }
      try {
        storage.setItem(INK_PREFERENCE_KEY, JSON.stringify({ schema: 1, pen: chosen }));
        sessionOnly = false;
        warning = null;
      } catch {
        sessionOnly = true;
        warning = 'Pen choice could not be saved; it is remembered for this session only.';
      }
      return chosen;
    },
  };
}
