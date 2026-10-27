/** Worksheet drawing coordinates use the same source axes as grid and print. */
import {
  createWorksheetDimensions,
  worksheetColumnOffsets,
  PIXELS_PER_POINT,
  type GeometryProfile,
  type WorksheetGeometrySource,
} from '../core/worksheet-geometry.mjs';
import type { WorkbookObjectAnchor, WorkbookObjectPreview, WorkbookObjectCommand } from '../core/contracts';
export const OBJECT_EMU_PER_PIXEL = 9525;
export interface ObjectBox {
  left: number;
  top: number;
  width: number;
  height: number;
}
export interface ObjectAxes {
  columns: Float64Array;
  rowOffset: (row: number) => number;
  box: (anchor: WorkbookObjectAnchor) => ObjectBox;
  anchor: (box: ObjectBox) => WorkbookObjectAnchor;
  extent: (box: ObjectBox) => { rows: number; columns: number };
}
const maxRows = 1048576,
  maxColumns = 16384;
const finite = (value: number) => {
  if (!Number.isFinite(value)) throw Error('Object geometry is unavailable.');
  return value;
};
function upperBound(offset: (n: number) => number, count: number, value: number): number {
  let first = 0,
    last = count + 1;
  while (first < last) {
    const middle = Math.floor((first + last) / 2);
    if (offset(middle) <= value) first = middle + 1;
    else last = middle;
  }
  return first;
}
function cell(ref: string): { row: number; col: number } {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(ref);
  if (!match) throw Error('Object anchor is unavailable.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(match[2]) - 1;
  if (col >= maxColumns || row >= maxRows) throw Error('Object anchor is outside the worksheet.');
  return { row, col };
}
function name(col: number, row: number): string {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters + (row + 1);
}
function integer(value: number, min: number, max: number): number {
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw Error('Object placement exceeds its supported drawing bounds.');
  return value;
}
export function createObjectAxes(sheet: WorksheetGeometrySource, profile?: GeometryProfile): ObjectAxes {
  const dimensions = createWorksheetDimensions(sheet, profile),
    columns = worksheetColumnOffsets(sheet, maxColumns, profile),
    rowOffset = (row: number) => dimensions.rowDistancePt(0, row) * PIXELS_PER_POINT;
  const position = (left: number, top: number) => {
    if (finite(left) < 0 || finite(top) < 0 || left >= columns[maxColumns] || top >= rowOffset(maxRows))
      throw Error('Object position is outside the visible worksheet axes.');
    return { col: upperBound((n) => columns[n], maxColumns, left) - 1, row: upperBound(rowOffset, maxRows, top) - 1 };
  };
  return {
    columns,
    rowOffset,
    box: (a) => {
      const p = cell(a.cell);
      if (a.kind !== 'oneCell') throw Error('Only one-cell object anchors are supported.');
      return {
        left: columns[p.col] + integer(a.dxEmu, 0, 10000000) / OBJECT_EMU_PER_PIXEL,
        top: rowOffset(p.row) + integer(a.dyEmu, 0, 10000000) / OBJECT_EMU_PER_PIXEL,
        width: integer(a.widthEmu, 1, 50000000) / OBJECT_EMU_PER_PIXEL,
        height: integer(a.heightEmu, 1, 50000000) / OBJECT_EMU_PER_PIXEL,
      };
    },
    anchor: (b) => {
      const p = position(b.left, b.top);
      return {
        kind: 'oneCell',
        cell: name(p.col, p.row),
        dxEmu: integer(Math.round((b.left - columns[p.col]) * OBJECT_EMU_PER_PIXEL), 0, 10000000),
        dyEmu: integer(Math.round((b.top - rowOffset(p.row)) * OBJECT_EMU_PER_PIXEL), 0, 10000000),
        widthEmu: integer(Math.round(finite(b.width) * OBJECT_EMU_PER_PIXEL), 1, 50000000),
        heightEmu: integer(Math.round(finite(b.height) * OBJECT_EMU_PER_PIXEL), 1, 50000000),
      };
    },
    extent: (b) => ({
      rows: Math.min(
        maxRows,
        Math.max(30, upperBound(rowOffset, maxRows, Math.min(rowOffset(maxRows), b.top + b.height)) + 1)
      ),
      columns: Math.min(
        maxColumns,
        Math.max(12, upperBound((n) => columns[n], maxColumns, Math.min(columns[maxColumns], b.left + b.width)) + 1)
      ),
    }),
  };
}
export function objectScreenBox(
  box: ObjectBox,
  view: { scale: number; rowHeaderWidth: number; columnHeaderHeight: number; scrollLeft: number; scrollTop: number }
): ObjectBox {
  return {
    left: view.rowHeaderWidth + box.left * view.scale - view.scrollLeft,
    top: view.columnHeaderHeight + box.top * view.scale - view.scrollTop,
    width: box.width * view.scale,
    height: box.height * view.scale,
  };
}
export function objectGestureBox(
  original: ObjectBox,
  kind: 'move' | 'resize',
  dx: number,
  dy: number,
  object: WorkbookObjectPreview
): ObjectBox {
  if (kind === 'resize' && object.kind === 'ink') throw Error('Stroke resizing is outside the owned drawing profile.');
  if (kind === 'move')
    return { ...original, left: Math.max(0, original.left + finite(dx)), top: Math.max(0, original.top + finite(dy)) };
  let width = Math.max(24, Math.min(50000000 / OBJECT_EMU_PER_PIXEL, original.width + finite(dx))),
    height = Math.max(24, Math.min(50000000 / OBJECT_EMU_PER_PIXEL, original.height + finite(dy)));
  if (object.kind === 'image') {
    if (!object.image?.widthPx || !object.image.heightPx) throw Error('Wait for the PNG preview before resizing.');
    const ratio = object.image.heightPx / object.image.widthPx,
      scale =
        Math.abs(dx / original.width) >= Math.abs(dy / original.height)
          ? width / original.width
          : height / original.height;
    width = Math.max(24, 24 / ratio, original.width * scale);
    height = width * ratio;
    if (width > 50000000 / OBJECT_EMU_PER_PIXEL || height > 50000000 / OBJECT_EMU_PER_PIXEL)
      throw Error('PNG size exceeds supported drawing bounds.');
  }
  return { ...original, width, height };
}
export function objectPlacementCommand(
  sheet: string,
  object: WorkbookObjectPreview,
  anchor: WorkbookObjectAnchor
): WorkbookObjectCommand {
  if (!object.managed || !object.canEdit || object.protectedReason)
    throw Error(object.protectedReason ?? 'This drawing is preserved and protected.');
  if (object.kind === 'ink' && object.stroke) {
    if (anchor.widthEmu !== object.stroke.widthEmu || anchor.heightEmu !== object.stroke.heightEmu)
      throw Error('Stroke movement preserves its geometry.');
    return {
      operation: 'updateInkStroke',
      sheet,
      objectId: object.id,
      expectedObjectRevision: object.revision,
      anchor,
    };
  }
  if (object.kind === 'chart')
    return {
      operation: 'updateChart',
      sheet,
      objectId: object.id,
      expectedObjectRevision: object.revision,
      patch: {},
      anchor,
    };
  if (object.kind === 'shape' && object.shape)
    return {
      operation: 'updateShape',
      sheet,
      objectId: object.id,
      expectedObjectRevision: object.revision,
      patch: {},
      anchor,
    };
  if (object.kind === 'smartart' && object.smartart)
    return {
      operation: 'updateSmartArt',
      sheet,
      objectId: object.id,
      expectedObjectRevision: object.revision,
      patch: {},
      anchor,
    };
  if (object.kind === 'model3d' && object.asset3d)
    return {
      operation: 'updateModel3D',
      sheet,
      objectId: object.id,
      expectedObjectRevision: object.revision,
      patch: {},
      anchor,
    };
  if (object.kind === 'image')
    return { operation: 'updateImage', sheet, objectId: object.id, expectedObjectRevision: object.revision, anchor };
  throw Error('This object kind cannot be moved or resized.');
}
