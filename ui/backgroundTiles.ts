import { FORMAT, matchesFormat } from '../core/format.mjs';
import type { WorkbookSheet } from '../core/contracts';
import { worksheetBackgroundSample } from '../core/page-style-samples.mjs';
export interface BackgroundTile {
  id: string;
  image: string;
  size: number;
}
/** Native import verifies sample bytes and relationship ownership. A changed
 * sheet or unknown descriptor must not display a catalog image in its place. */
export function backgroundTile(sheet: WorkbookSheet | undefined, scale: number): BackgroundTile | undefined {
  const value = sheet?.background;
  if (
    !value ||
    !matchesFormat(value.schema, 'SCHEMA_SHEET_BACKGROUND_V1') ||
    value.sheet !== sheet?.path ||
    !value.canEdit ||
    value.protectedReason !== null ||
    !value.revision ||
    !value.sample ||
    !Number.isFinite(scale) ||
    scale <= 0 ||
    scale > 4
  )
    return undefined;
  try {
    const sample = worksheetBackgroundSample(value.sample);
    return { id: sample.id, image: `url("data:${sample.mime};base64,${sample.base64}")`, size: sample.width * scale };
  } catch {
    return undefined;
  }
}
export function backgroundCellStyle(tile: BackgroundTile | undefined, x: number, y: number, hasFill: boolean) {
  if (!tile || hasFill) return {};
  return {
    backgroundImage: tile.image,
    backgroundRepeat: 'repeat',
    backgroundSize: `${tile.size}px ${tile.size}px`,
    backgroundPosition: `${-x}px ${-y}px`,
  };
}
