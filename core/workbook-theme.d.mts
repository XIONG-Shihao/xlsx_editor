export interface WorkbookThemeProjection {
  readonly part: string;
  readonly name: string;
  readonly colors: Readonly<Record<string, string>>;
  readonly fonts: Readonly<Partial<Record<'major' | 'minor', string>>>;
}
export const spreadsheetThemeColors: readonly string[];
export function readWorkbookTheme(
  parts: Record<string, Uint8Array>,
  relationships: Iterable<{ type: string; path: string | null }>
): { theme: WorkbookThemeProjection | null; issues: string[] };
export function resolveThemeColor(
  node: Element | null | undefined,
  theme: WorkbookThemeProjection | null,
  issues: Set<string>
): string | undefined;
export function resolveThemeFont(
  explicit: string | null | undefined,
  scheme: string | null | undefined,
  theme: WorkbookThemeProjection | null,
  issues: Set<string>
): string | undefined;
