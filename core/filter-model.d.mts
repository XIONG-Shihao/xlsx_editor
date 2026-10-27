import type { WorkbookFilterDefinition } from './contracts.js';
export const FILTER_URI: string;
/** Internal source-derived provenance; never part of WorkbookFilterDefinition. */
export interface OwnedFilterRecord {
  model: {
    schema: 1;
    range: string;
    column: string;
    values: unknown[];
    baseline: { format: Record<string, string>; rows: Array<[number, string | null]> };
  };
  text: string;
  revision: string;
}
export function inspectFilter(filter: Element | null): {
  definition?: WorkbookFilterDefinition;
  record: OwnedFilterRecord | null;
};
export function filterDefinition(record: OwnedFilterRecord): WorkbookFilterDefinition;
export function prepareFilterState(
  command: unknown,
  record: OwnedFilterRecord | null,
  baseline: unknown
): OwnedFilterRecord;
