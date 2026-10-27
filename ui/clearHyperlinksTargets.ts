import type { WorkbookClearHyperlinksCommand } from '../core/contracts';
import { clearHyperlinksCommand } from '../core/clear-hyperlinks.mjs';
import { addressPosition, colName } from './sheetModel';
import { captureHomeBinding, type HomeBinding } from './homeCommands';

/** Exact captured selection geometry only. No cells, hyperlinks, styles or URLs
 * are read in the renderer; the original native owner owns removal admission. */
export function clearHyperlinksTargets(input: HomeBinding): WorkbookClearHyperlinksCommand {
  const binding = captureHomeBinding(input);
  if (binding.mode !== 'normal' || !binding.ownerIdentity)
    throw Error('Clear Hyperlinks needs an accepted worksheet owner in Normal view.');
  if (!binding.specialSelection) {
    const [start, end = start] = binding.selection.split(':'),
      a = addressPosition(start),
      b = addressPosition(end),
      first = colName(a.col) + (a.row + 1),
      last = colName(b.col) + (b.row + 1);
    return clearHyperlinksCommand({
      type: 'clearHyperlinks',
      sheet: binding.sheet,
      ranges: [first === last ? first : first + ':' + last],
    });
  }
  const points = binding.specialSelection
      .map((cell) => addressPosition(cell))
      .sort((a, b) => a.row - b.row || a.col - b.col),
    runs: Array<{ row: number; first: number; last: number }> = [];
  for (const at of points) {
    const last = runs.at(-1);
    if (last && last.row === at.row && last.last + 1 === at.col) last.last = at.col;
    else runs.push({ row: at.row, first: at.col, last: at.col });
  }
  // Join identical column runs on adjacent rows, without filling an L-shaped
  // hole or including any unselected cell in a sparse selection.
  const areas: Array<{ firstRow: number; lastRow: number; first: number; last: number }> = [],
    previous = new Map<string, number>();
  for (const run of runs) {
    const key = run.first + ':' + run.last,
      index = previous.get(key),
      prior = index === undefined ? undefined : areas[index];
    if (prior && prior.lastRow + 1 === run.row) prior.lastRow = run.row;
    else {
      previous.set(key, areas.length);
      areas.push({ firstRow: run.row, lastRow: run.row, first: run.first, last: run.last });
    }
  }
  const ranges = areas.map((area) => {
    const first = colName(area.first) + (area.firstRow + 1),
      last = colName(area.last) + (area.lastRow + 1);
    return first === last ? first : first + ':' + last;
  });
  return clearHyperlinksCommand({ type: 'clearHyperlinks', sheet: binding.sheet, ranges });
}
