import { inkStrokeHit, INK_EMU_PER_PIXEL } from '../core/ink-stroke.mjs';
import type { WorkbookObjectPreview } from '../core/contracts';
import type { ObjectAxes } from './worksheetObjectGeometry';
/** Pick the last painted qualified owned path, with six screen pixels of tolerance. */
export function pickOwnedInk(
  objects: readonly WorkbookObjectPreview[],
  point: readonly [number, number],
  axes: ObjectAxes,
  scale: number
): WorkbookObjectPreview | null {
  if (!Number.isFinite(scale) || scale < 0.1 || scale > 4) throw Error('Invalid stroke selection zoom.');
  for (const object of [...objects].reverse()) {
    if (
      object.kind !== 'ink' ||
      !object.managed ||
      !object.canEdit ||
      !object.canDelete ||
      object.protectedReason ||
      !object.stroke ||
      !object.anchor
    )
      continue;
    const box = axes.box(object.anchor);
    if (
      inkStrokeHit(
        object.stroke,
        (point[0] - box.left) * INK_EMU_PER_PIXEL,
        (point[1] - box.top) * INK_EMU_PER_PIXEL,
        Math.round((6 / scale) * INK_EMU_PER_PIXEL)
      )
    )
      return object;
  }
  return null;
}
