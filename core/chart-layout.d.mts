import type { WorkbookObjectPreview } from './contracts';
export function chartPreviewGeometry(object: WorkbookObjectPreview): {
  width: number;
  height: number;
  left: number;
  plotWidth: number;
  lines: number[][];
  bars: Array<{ x: number; y: number; width: number; height: number; series: number; title: string }>;
  segments: Array<{ series: number; points: number[][] }>;
  labels: Array<{ x: number; y: number; anchor: 'middle' | 'end'; text: string }>;
};
