import {
  inkStrokeFromLogicalPoints,
  INK_EMU_PER_PIXEL,
  type InkPenId,
  type InkStrokeModel,
} from '../core/ink-stroke.mjs';
import type { WorkbookObjectAnchor } from '../core/contracts';
import type { ObjectAxes, ObjectBox } from './worksheetObjectGeometry';
export interface InkViewport {
  scale: number;
  rowHeaderWidth: number;
  columnHeaderHeight: number;
  scrollLeft: number;
  scrollTop: number;
  frozenRows?: number;
  frozenColumns?: number;
}
/** Input is local to the existing grid viewport. Sticky headings do not become points. */
export function inkWorksheetPoint(view: InkViewport, x: number, y: number): [number, number] | null {
  if (
    !Number.isFinite(view.scale) ||
    view.scale < 0.1 ||
    view.scale > 4 ||
    [view.rowHeaderWidth, view.columnHeaderHeight, view.scrollLeft, view.scrollTop].some(
      (n) => !Number.isFinite(n) || n < 0
    ) ||
    !Number.isFinite(x) ||
    !Number.isFinite(y)
  )
    throw Error('Invalid ink viewport.');
  if (view.frozenRows || view.frozenColumns)
    throw Error('Drawing across frozen panes is outside the first stroke profile.');
  const left = x - view.rowHeaderWidth,
    top = y - view.columnHeaderHeight;
  if (left < 0 || top < 0) return null;
  return [(left + view.scrollLeft) / view.scale, (top + view.scrollTop) / view.scale];
}
export function anchoredInkStroke(
  points: readonly (readonly [number, number])[],
  pen: InkPenId,
  axes: Pick<ObjectAxes, 'anchor' | 'box'>
): { stroke: InkStrokeModel; anchor: WorkbookObjectAnchor; box: ObjectBox } {
  const { stroke, originEmu } = inkStrokeFromLogicalPoints(points, pen),
    box = {
      left: originEmu.x / INK_EMU_PER_PIXEL,
      top: originEmu.y / INK_EMU_PER_PIXEL,
      width: stroke.widthEmu / INK_EMU_PER_PIXEL,
      height: stroke.heightEmu / INK_EMU_PER_PIXEL,
    },
    anchor = axes.anchor(box),
    accepted = axes.box(anchor);
  if (
    anchor.widthEmu !== stroke.widthEmu ||
    anchor.heightEmu !== stroke.heightEmu ||
    Math.abs(accepted.left - box.left) > 1 / INK_EMU_PER_PIXEL ||
    Math.abs(accepted.top - box.top) > 1 / INK_EMU_PER_PIXEL
  )
    throw Error('Worksheet axes changed the captured stroke placement.');
  return { stroke, anchor, box };
}
