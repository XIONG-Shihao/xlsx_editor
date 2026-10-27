export interface PageThemeSample {
  readonly id: string;
  readonly label: string;
  readonly majorFont: string;
  readonly minorFont: string;
  readonly colors: Readonly<Record<string, string>>;
}
export interface WorksheetBackgroundSample {
  readonly id: string;
  readonly label: string;
  readonly width: number;
  readonly height: number;
  readonly mime: 'image/png';
  readonly base64: string;
  readonly sha256: string;
}
export const pageThemeSamples: ReadonlyArray<PageThemeSample>;
export function pageThemeSample(id: string): PageThemeSample;
export function pageThemeXml(id: string): string;
export const worksheetBackgroundSamples: ReadonlyArray<WorksheetBackgroundSample>;
export function worksheetBackgroundSample(id: string): WorksheetBackgroundSample;
