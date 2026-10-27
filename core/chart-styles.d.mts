import type { WorkbookChartStyle } from './contracts';
export const CHART_STYLES: readonly Readonly<{
  id: WorkbookChartStyle;
  label: string;
  colors: readonly string[];
  outline: boolean;
}>[];
export function chartStyle(
  id: unknown
): Readonly<{ id: WorkbookChartStyle; label: string; colors: readonly string[]; outline: boolean }>;
export function chartSeriesAppearance(
  id: WorkbookChartStyle | undefined,
  index: number
): { color: string; fill: string; stroke: string | undefined; width: number | undefined; dash: string | undefined };
export function chartSeriesPropertiesXml(id: WorkbookChartStyle | undefined, index: number): string;
