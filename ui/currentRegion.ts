import type { WorkbookCell } from '../core/contracts';
import type { SpecialSelectionReader } from './specialSelection';

export const CURRENT_REGION_LIMITS = Object.freeze({
  positions: 10_000,
  discoveryPositions: 60_000,
  readPositions: 1000,
  bytes: 16 * 1024 * 1024,
});
type Point = { row: number; column: number };
type Bounds = { first: Point; last: Point };
const address = (value: string): Point => {
  const m = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(value);
  if (!m) throw Error('Current Region needs a canonical cell.');
  const column = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(m[2]) - 1;
  if (column > 16383 || row > 1048575) throw Error('Current Region cell is outside the worksheet.');
  return { row, column };
};
const ref = (row: number, column: number) => {
  let name = '';
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name + (row + 1);
};
const range = (b: Bounds) => {
  const a = ref(b.first.row, b.first.column),
    z = ref(b.last.row, b.last.column);
  return a === z ? a : a + ':' + z;
};
const size = (b: Bounds) => (b.last.row - b.first.row + 1) * (b.last.column - b.first.column + 1);
const inside = (p: Point, b: Bounds) =>
  p.row >= b.first.row && p.row <= b.last.row && p.column >= b.first.column && p.column <= b.last.column;
const plain = (x: unknown): x is Record<string, unknown> =>
  !!x &&
  typeof x === 'object' &&
  !Array.isArray(x) &&
  (Object.getPrototypeOf(x) === Object.prototype || Object.getPrototypeOf(x) === null);
const key = (r: number, c: number) => r * 16384 + c;
interface Fact {
  occupied: boolean;
  grouped: boolean;
}
export interface CurrentRegionInput {
  sheet: string;
  anchor: string;
  merges: readonly string[];
  read: SpecialSelectionReader;
  check(): void;
  validateCell(value: WorkbookCell): WorkbookCell;
  yieldWork?: () => Promise<void>;
}
export interface CurrentRegionResult {
  range: string;
  positions: number;
  anchor: string;
  references: string[];
  readRanges: string[];
}

/** Ordinary populated blocks, bounded by empty source rows/columns. Reads use
 * authored cell presence, not cached display text or loaded viewport cells. */
export async function collectCurrentRegion(input: CurrentRegionInput): Promise<CurrentRegionResult> {
  const anchor = address(input.anchor),
    cache = new Map<number, Fact>(),
    readRanges: string[] = [],
    steps = [1, 1, 1, 1];
  let readBytes = 0,
    bounds: Bounds = { first: { ...anchor }, last: { ...anchor } };
  const merges = input.merges.map((value) => {
    const [a, z = a] = value.split(':');
    return { first: address(a), last: address(z) };
  });
  const check = () => {
    input.check();
    if (size(bounds) > CURRENT_REGION_LIMITS.positions)
      throw Error('Current Region supports up to 10,000 positions. Select a smaller data block.');
  };
  async function ensure(b: Bounds) {
    check();
    const pending: Bounds[] = [];
    for (let r = b.first.row; r <= b.last.row; r++) {
      for (let c = b.first.column; c <= b.last.column; ) {
        if (cache.has(key(r, c))) {
          c++;
          continue;
        }
        const start = c;
        while (c <= b.last.column && !cache.has(key(r, c)) && c - start < CURRENT_REGION_LIMITS.readPositions) c++;
        pending.push({ first: { row: r, column: start }, last: { row: r, column: c - 1 } });
      }
    }
    const added = pending.reduce((n, b) => n + size(b), 0);
    if (cache.size + added > CURRENT_REGION_LIMITS.discoveryPositions)
      throw Error('Current Region discovery exceeds its 60,000-position limit.');
    // Batch adjacent equal-width rows into ordinary bounded owner reads.
    const windows: Bounds[] = [];
    for (const p of pending) {
      const last = windows.at(-1);
      if (
        last &&
        last.last.row + 1 === p.first.row &&
        last.first.column === p.first.column &&
        last.last.column === p.last.column &&
        size({ ...last, last: p.last }) <= CURRENT_REGION_LIMITS.readPositions
      )
        last.last.row = p.last.row;
      else windows.push(structuredClone(p));
    }
    for (const w of windows) {
      check();
      const name = range(w),
        reply = await input.read(input.sheet, name);
      check();
      if (
        !plain(reply) ||
        Object.keys(reply).sort().join(',') !== 'cells,complete,range' ||
        reply.range !== name ||
        reply.complete !== true ||
        !plain(reply.cells)
      )
        throw Error('Current Region needs a complete original-owner source read.');
      readBytes += new TextEncoder().encode(JSON.stringify(reply)).length;
      if (readBytes > CURRENT_REGION_LIMITS.bytes) throw Error('Current Region source replies exceed 16 MiB.');
      const facts = new Map<number, Fact>();
      for (const [reference, value] of Object.entries(reply.cells)) {
        const p = address(reference);
        if (!inside(p, w)) throw Error('Current Region reader returned a foreign cell.');
        const cell = input.validateCell(value as WorkbookCell);
        facts.set(key(p.row, p.column), {
          occupied: cell.formula !== undefined || cell.formulaType !== undefined || cell.value !== null,
          grouped: !!cell.formulaType || !!cell.spill || !!cell.formulaOutput || !!cell.scenarioResult,
        });
      }
      for (let r = w.first.row; r <= w.last.row; r++)
        for (let c = w.first.column; c <= w.last.column; c++)
          cache.set(key(r, c), facts.get(key(r, c)) ?? { occupied: false, grouped: false });
      readRanges.push(name);
      if (input.yieldWork) await input.yieldWork();
      check();
    }
  }
  await ensure(bounds);
  if (!cache.get(key(anchor.row, anchor.column))?.occupied)
    throw Error('Select a populated cell inside an ordinary data block before using Current Region.');
  for (;;) {
    check();
    const old = structuredClone(bounds),
      sides: Bounds[] = [
        {
          first: { row: Math.max(0, old.first.row - steps[0]), column: old.first.column },
          last: { row: old.first.row - 1, column: old.last.column },
        },
        {
          first: { row: old.last.row + 1, column: old.first.column },
          last: { row: Math.min(1048575, old.last.row + steps[1]), column: old.last.column },
        },
        {
          first: { row: old.first.row, column: Math.max(0, old.first.column - steps[2]) },
          last: { row: old.last.row, column: old.first.column - 1 },
        },
        {
          first: { row: old.first.row, column: old.last.column + 1 },
          last: { row: old.last.row, column: Math.min(16383, old.last.column + steps[3]) },
        },
      ];
    for (let side = 0; side < 4; side++) {
      const b = sides[side];
      if (b.first.row > b.last.row || b.first.column > b.last.column) continue;
      await ensure(b);
      let count = 0;
      if (side < 2) {
        for (
          let r = side === 0 ? b.last.row : b.first.row;
          side === 0 ? r >= b.first.row : r <= b.last.row;
          r += side === 0 ? -1 : 1
        ) {
          let occupied = false;
          for (let c = b.first.column; c <= b.last.column; c++) occupied ||= cache.get(key(r, c))!.occupied;
          if (!occupied) break;
          count++;
        }
        if (side === 0) bounds.first.row -= count;
        else bounds.last.row += count;
      } else {
        for (
          let c = side === 2 ? b.last.column : b.first.column;
          side === 2 ? c >= b.first.column : c <= b.last.column;
          c += side === 2 ? -1 : 1
        ) {
          let occupied = false;
          for (let r = b.first.row; r <= b.last.row; r++) occupied ||= cache.get(key(r, c))!.occupied;
          if (!occupied) break;
          count++;
        }
        if (side === 2) bounds.first.column -= count;
        else bounds.last.column += count;
      }
      steps[side] = count === steps[side] ? Math.min(16, steps[side] * 2) : 1;
      check();
    }
    await ensure(bounds);
    if (
      merges.some(
        (m) =>
          m.first.row <= bounds.last.row &&
          m.last.row >= bounds.first.row &&
          m.first.column <= bounds.last.column &&
          m.last.column >= bounds.first.column
      )
    )
      throw Error('Current Region for merged data is deferred. Use a regular explicit selection.');
    for (let r = bounds.first.row; r <= bounds.last.row; r++)
      for (let c = bounds.first.column; c <= bounds.last.column; c++)
        if (cache.get(key(r, c))?.grouped) throw Error('Current Region for array/spill/scenario groups is deferred.');
    if (range(old) === range(bounds)) break;
  }
  // Diagonal-only neighboring blocks have ambiguous extent semantics in this
  // first ordinary-block profile; refuse instead of silently joining them.
  for (const r of [bounds.first.row - 1, bounds.last.row + 1])
    for (const c of [bounds.first.column - 1, bounds.last.column + 1])
      if (r >= 0 && r <= 1048575 && c >= 0 && c <= 16383) {
        await ensure({ first: { row: r, column: c }, last: { row: r, column: c } });
        if (cache.get(key(r, c))?.occupied)
          throw Error('Diagonal-only neighboring data is outside the ordinary Current Region profile.');
      }
  const references: string[] = [];
  for (let r = bounds.first.row; r <= bounds.last.row; r++)
    for (let c = bounds.first.column; c <= bounds.last.column; c++) references.push(ref(r, c));
  check();
  return { range: range(bounds), positions: size(bounds), anchor: input.anchor, references, readRanges };
}
