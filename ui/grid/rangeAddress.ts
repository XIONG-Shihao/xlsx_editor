import { addressPosition, colName } from '../sheetModel';

export function rectangle(from: string, to: string): string {
  const a = addressPosition(from),
    b = addressPosition(to);
  const first = `${colName(Math.min(a.col, b.col))}${Math.min(a.row, b.row) + 1}`,
    last = `${colName(Math.max(a.col, b.col))}${Math.max(a.row, b.row) + 1}`;
  return first === last ? first : `${first}:${last}`;
}
export function selectedBounds(
  range: string,
  limit: number
): { first: { row: number; col: number }; last: { row: number; col: number }; rows: number; columns: number } {
  if (range.split(':').length > 2) throw Error('Select one rectangular range.');
  const [from, to = from] = range.split(':'),
    first = addressPosition(from),
    last = addressPosition(to);
  if (first.row < 0 || first.col < 0 || last.row > 1048575 || last.col > 16383)
    throw Error('The selection is outside the worksheet limits.');
  const rows = last.row - first.row + 1,
    columns = last.col - first.col + 1;
  if (rows < 1 || columns < 1 || rows * columns > limit)
    throw Error(`Selection exceeds the ${limit.toLocaleString()}-cell limit.`);
  return { first, last, rows, columns };
}
export function rangesOverlap(a: string, b: string): boolean {
  const first = selectedBounds(a, 10000),
    second = selectedBounds(b, 10000);
  return (
    first.first.row <= second.last.row &&
    second.first.row <= first.last.row &&
    first.first.col <= second.last.col &&
    second.first.col <= first.last.col
  );
}

export function inRange(range: string, cell: string): boolean {
  try {
    const [from, to = from] = range.split(':'),
      a = addressPosition(from),
      b = addressPosition(to),
      p = addressPosition(cell);
    return p.row >= a.row && p.row <= b.row && p.col >= a.col && p.col <= b.col;
  } catch {
    return false;
  }
}
