import type { InkStrokeModel } from './ink-stroke.mjs';
import type { WorkbookObjectAnchor } from './contracts';
export const INK_DRAW_NS: string;
export const INK_MAIN_NS: string;
export function inkStrokeAnchorXml(
  id: string,
  number: number,
  stroke: InkStrokeModel,
  anchor: WorkbookObjectAnchor,
  altText?: string
): string;
export function readInkStrokeAnchor(source: string): {
  id: string;
  number: number;
  stroke: InkStrokeModel;
  anchor: WorkbookObjectAnchor;
  altText: string;
};
