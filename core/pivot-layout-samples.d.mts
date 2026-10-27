import type { WorkbookPivotLayoutSample } from './contracts';
export interface PivotLayoutSample {
  readonly id: WorkbookPivotLayoutSample;
  readonly label: string;
  readonly compact: 0 | 1;
  readonly outline: 0 | 1;
  readonly style: string;
  readonly stripes: 0 | 1;
}
export const PIVOT_LAYOUT_SAMPLES: readonly PivotLayoutSample[];
export function pivotLayoutSample(id: WorkbookPivotLayoutSample): PivotLayoutSample;
export function pivotPresentation(root: unknown): { sample: WorkbookPivotLayoutSample; grandTotals: boolean };
export function pivotOutputData(
  model: unknown,
  options?: { sample?: WorkbookPivotLayoutSample; grandTotals?: boolean }
): Array<Array<string | number>>;
