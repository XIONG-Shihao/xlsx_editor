import type { WorkbookChartModel } from './contracts';
export const CHART_SOURCE_LIMIT: number;
export interface ChartData {
  categories: string[];
  series: Array<{ name: string; values: Array<number | null> }>;
  hidden: boolean[];
}
export function normalizeChartModel(value: unknown): WorkbookChartModel;
export function readChartData(
  readRange: (sheet: string, range: string) => Promise<{ cells: Record<string, unknown> } | Record<string, unknown>>,
  sheet: string,
  model: WorkbookChartModel,
  view: { sheets: Array<{ name: string; hiddenRows?: number[] }> }
): Promise<ChartData>;
export function createChartRecommendations(input: {
  sheet: string;
  range: string;
  view: { sheets: Array<{ name: string; hiddenRows?: number[] }> };
  readRange: (sheet: string, range: string) => Promise<{ cells: Record<string, unknown> } | Record<string, unknown>>;
}): Promise<{
  sheet: string;
  range: string;
  rows: number;
  series: number;
  proposals: Array<{ id: string; label: string; reason: string; chart: WorkbookChartModel; data: ChartData }>;
}>;
