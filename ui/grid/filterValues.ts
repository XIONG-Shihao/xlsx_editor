import type { WorkbookFilterValuePage } from '../../core/contracts';

export type FilterValue = WorkbookFilterValuePage['values'][number];
export type FilterValuePage = WorkbookFilterValuePage;
export function filterValueLabel(value: FilterValue['value'], words: (text: string) => string): string {
  if (value === null) return words('(blank)');
  if (typeof value === 'string') return JSON.stringify(value) + ' ' + words('(text)');
  return (
    (Object.is(value, -0) ? '-0' : String(value)) + ' ' + words(typeof value === 'boolean' ? '(boolean)' : '(number)')
  );
}
