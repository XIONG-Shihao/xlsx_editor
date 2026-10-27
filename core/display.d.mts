export interface WorkbookDisplayedText {
  text: string;
  formatStatus: 'supported' | 'unsupported' | 'not-applicable';
  color?: string;
  reason?: 'unsupported-custom-format' | 'precision-preserved';
}
export function displayCellDetails(cell: any, date1904?: boolean): WorkbookDisplayedText;
export function displayCell(cell: any, date1904?: boolean): string;
export function formatDisplayDetails(
  value: unknown,
  format: number | string | undefined,
  date1904?: boolean
): WorkbookDisplayedText;
export function formatDisplayValue(value: unknown, format: number | string | undefined, date1904?: boolean): string;
