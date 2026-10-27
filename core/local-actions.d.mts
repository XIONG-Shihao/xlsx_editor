export type LocalActionFormat = Partial<{
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  horizontalAlign: 'left' | 'center' | 'right';
  verticalAlign: 'top' | 'center' | 'bottom';
  wrapText: boolean;
  border: 'none' | 'all';
  numberFormat:
    | 'general'
    | 'number'
    | 'percent'
    | 'date'
    | 'text'
    | 'currency'
    | 'accounting'
    | 'thousands'
    | 'decimal0'
    | 'decimal1'
    | 'decimal2';
}>;
export type LocalActionDescriptor =
  | { schema: 1; command: 'freeze'; parameters: { sheet: string; rows: number; columns: number } }
  | { schema: 1; command: 'formatRange'; parameters: { sheet: string; range: string; format: LocalActionFormat } }
  | { schema: 1; command: 'recalculate'; parameters: Record<string, never> }
  | {
      schema: 1;
      command: 'sortRows';
      parameters: { sheet: string; range: string; keys: Array<{ column: string; direction: 'asc' | 'desc' }> };
    };
export interface LocalActionDraft {
  name: string;
  descriptor: LocalActionDescriptor;
}
export interface LocalActionEntry extends LocalActionDraft {
  id: string;
  revision: number;
  author: 'local-profile';
  createdAt: number;
  updatedAt: number;
}
export interface LocalActionList {
  schema: 1;
  scope: 'private-profile';
  revision: number;
  entries: LocalActionEntry[];
  warning?: string;
}
export interface LocalActionCatalogEntry {
  id: string;
  label: string;
  command: LocalActionDescriptor['command'];
  descriptorSchema: 1;
  supportedProfiles: readonly ['xlsx-writable'];
  requiresWrite: true;
  readScope: string;
  writeScope: string;
  budgets: Readonly<Record<string, number>>;
  prerequisites: readonly string[];
  capabilityReason: string;
}
export const LOCAL_ACTION_LIMITS: Readonly<{
  entries: 128;
  registryBytes: 1048576;
  descriptorBytes: 65536;
  nameCharacters: 128;
  nameBytes: 512;
  freezeRows: 100;
  freezeColumns: 30;
  formatPositions: 10000;
  sortPositions: 1000000;
  sortKeys: 16;
}>;
export const LOCAL_ACTION_CATALOG: readonly LocalActionCatalogEntry[];
export function localActionCatalog(command: LocalActionDescriptor['command']): LocalActionCatalogEntry;
export function validateLocalActionDescriptor(value: unknown): LocalActionDescriptor;
export function validateLocalActionDraft(value: unknown): LocalActionDraft;
export function localActionId(value: unknown): string;
export function localActionRevision(value: unknown): number;
export function validateLocalActionEntry(value: unknown): LocalActionEntry;
export function validateLocalActionList(value: unknown): LocalActionList;
