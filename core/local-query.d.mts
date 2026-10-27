import { FORMAT } from './format.mjs';
import type { DelimitedSource, DelimitedInspection, DelimitedProgress } from './delimited-import.mjs';
import type { PasteShape, PasteDescriptor, PastePage, TypedPasteConversion } from './paste-protocol.mjs';
import type { TextConversionPolicy, TextConversionResult } from './text-conversion.mjs';
export interface LocalQuery {
  columns: number[];
  filters: {
    column: number;
    op: 'equals' | 'notEquals' | 'contains' | 'isEmpty' | 'notEmpty';
    value: string;
    caseSensitive: boolean;
  }[];
  keepHeader: boolean;
  policies?: TextConversionPolicy[];
}
export interface LocalQueryDraft {
  name: string;
  target: { sheet: string; start: string };
  query: LocalQuery;
}
export interface LocalQueryOutput {
  shape: PasteShape;
  fingerprint: string;
  sourceSha256: string;
  sourceBytes: number;
  outputRows: number;
  outputColumns: number;
  acceptedHead: string;
  updatedAt: number;
  fingerprintProfile?: 'typed-semantic-v1';
  conversion?: TypedPasteConversion;
}
export interface LocalQueryConnection extends LocalQueryDraft {
  id: string;
  revision: number;
  source: { id: string; name: string };
  last?: LocalQueryOutput;
  pending?: boolean;
}
export interface LocalQueryList {
  entries: LocalQueryConnection[];
  warning?: string;
}
export interface LocalQueryConversionIssue {
  index: number;
  sourceRow: number;
  sourceColumn: number;
  outputRow: number;
  outputColumn: number;
  sourceRef: string;
  destinationRef: string;
  raw: string;
  policy: TextConversionPolicy;
  status: 'converted' | 'rejected';
  kind?: 'text' | 'number' | 'boolean' | 'date';
  display?: string;
  reason?: string;
  notices: string[];
}
export interface LocalQueryConversionInspection {
  schema: 1;
  revision: string;
  date1904: boolean | null;
  policies: TextConversionPolicy[];
  rejected: number;
  changed: number;
  totalIssues: number;
  offset: number;
  nextOffset: number | null;
  issues: LocalQueryConversionIssue[];
  expiresAt: number;
}
export interface LocalQueryReplacementRow {
  index: number;
  sourceRef: string | null;
  destinationRef: string;
  sourceText: string;
  outputKind: 'text' | 'number' | 'boolean' | 'date' | 'rejected';
  outputDisplay: string;
  current: {
    kind: 'blank' | 'text' | 'number' | 'boolean' | 'error' | 'formula' | 'protected';
    display: string;
    rawValue: string | null;
    formula: string | null;
    format: string | number | null;
    reason?: string;
    pending?: boolean;
    storedDisplay?: string;
  };
}
export interface LocalQueryReplacementInspection {
  schema: 1;
  revision: string;
  occupied: number;
  blocked: number;
  totalNewPositions: number;
  offset: number;
  nextOffset: number | null;
  rows: LocalQueryReplacementRow[];
}
export interface LocalQueryPreview {
  connectionId: string;
  revision: number;
  source: { id: string; name: string; bytes: number; sha256: string };
  outputRows: number;
  outputColumns: number;
  rows: string[][];
  truncated: boolean;
  target: { sheet: string; start: string };
  conversionInspection?: LocalQueryConversionInspection;
  replacementInspection?: LocalQueryReplacementInspection;
}
export interface LocalQueryPreviewOptions {
  offset?: number;
  replacementOffset?: number;
  expectedConversionRevision?: string;
}
export interface LocalQueryRefreshOptions {
  expectedConversionRevision?: string;
  expectedReplacementRevision?: string;
}
export interface LocalQueryInspection {
  schema: typeof FORMAT.SCHEMA_LOCAL_QUERY_V1;
  query: LocalQuery;
  raw: DelimitedInspection;
  source: DelimitedInspection['source'];
  outputRows: number;
  outputColumns: number;
  dataRows: number;
  shape: PasteShape;
  count: number;
  policy: 'text' | 'typed-text-v1';
  conversion?: TypedPasteConversion;
  conversionSummary: { rejected: number; changed: number; issues: number; fingerprint: string };
  preview: string[][];
  previewTruncated: boolean;
  outputFingerprint?: string;
}
export interface LocalQueryInspectedCell {
  index: number;
  raw: string;
  sourceRow: number;
  sourceColumn: number;
  outputRow: number;
  outputColumn: number;
  target: { sheet: string; start: string };
  policy: TextConversionPolicy;
  result: TextConversionResult;
}
export const QUERY_COLUMNS: number,
  QUERY_ROW_BYTES: number,
  QUERY_PREVIEW_BYTES: number,
  QUERY_TYPED_FINGERPRINT: 'typed-semantic-v1';
export function validateLocalQuery(value: unknown): LocalQuery;
export function queryHasConversions(query: LocalQuery): boolean;
export function queryReplacementIdentity(cell: unknown): [string, string];
export function queryReplacementCurrent(cell: unknown): {
  current: LocalQueryReplacementRow['current'];
  occupied: boolean;
  blocked: boolean;
};
export function queryCellIdentity(cell: unknown): [string, unknown];
export function typedQueryCellIdentity(
  cell: unknown,
  options: { conversion: TypedPasteConversion; shape: PasteShape; index: number }
): [string, unknown];
export function queryConversionDisplay(result: TextConversionResult): string | undefined;
export function localQueryConversionIssue(cell: LocalQueryInspectedCell): LocalQueryConversionIssue;
export function queryCellFingerprint(): {
  add(index: number, kind: string, value: unknown): void;
  finish(): string;
  close(): void;
};
export function inspectLocalQuery(
  source: DelimitedSource,
  options: {
    sheet: string;
    start: string;
    query: LocalQuery;
    previousShape?: PasteShape;
    date1904?: boolean;
    signal?: AbortSignal;
    onProgress?: (event: DelimitedProgress) => void;
    onCell?: (cell: LocalQueryInspectedCell) => void | Promise<void>;
    onLayout?: (layout: { outputColumns: number; sourceRows: number; columns: number[] }) => void | Promise<void>;
  }
): Promise<LocalQueryInspection>;
export function localQueryFields(
  source: DelimitedSource,
  inspection: LocalQueryInspection,
  options?: {
    signal?: AbortSignal;
    onProgress?: (event: DelimitedProgress) => void;
    onComplete?: (proof: {
      source: DelimitedInspection['source'];
      outputRows: number;
      outputColumns: number;
      shape: PasteShape;
      fingerprint: string;
    }) => void | Promise<void>;
  }
): AsyncGenerator<LocalQueryInspectedCell>;
export function localQueryPages(
  source: DelimitedSource,
  inspection: LocalQueryInspection,
  options?: { signal?: AbortSignal; onProgress?: (event: DelimitedProgress) => void }
): AsyncGenerator<PastePage>;
export function stageLocalQuery(
  source: DelimitedSource,
  inspection: LocalQueryInspection,
  options: {
    state: string;
    signal?: AbortSignal;
    checkCurrent?: () => void | Promise<void>;
    sink: any;
    onProgress?: (event: DelimitedProgress) => void;
  }
): Promise<{ descriptor: PasteDescriptor; outputFingerprint: string }>;

/** Source-wide preview covers every captured local connection. Samples are bounded. */
export interface LocalQueryRefreshAllScope {
  documentId: string;
  ownerId?: string | null;
  state: string;
}
export interface LocalQueryRefreshAllEntry {
  connectionId: string;
  revision: number;
  name: string;
  target: { sheet: string; start: string };
  status: 'ready' | 'refused';
  reason?: string;
  source?: LocalQueryPreview['source'];
  outputRows?: number;
  outputColumns?: number;
  range?: string;
  sample?: string[][];
  sampleTruncated?: boolean;
  conversions?: { rejected: number; changed: number; policies: TextConversionPolicy[]; date1904: boolean | null };
  replacements?: { occupied: number; blocked: number };
}
export interface LocalQueryRefreshAllPreview {
  schema: 1;
  id: string;
  scope: LocalQueryRefreshAllScope;
  entries: LocalQueryRefreshAllEntry[];
  expiresAt: number;
  undoPolicy: 'One Undo step for each accepted connection';
  warning?: string;
}
export interface LocalQueryRefreshAllOutcome {
  connectionId: string;
  name: string;
  target: { sheet: string; start: string };
  status: 'refreshed' | 'refused' | 'cancelled' | 'not-run' | 'uncertain';
  undoUnits: 0 | 1 | null;
  parentState?: string;
  state?: string;
  outputRows?: number;
  outputColumns?: number;
  reason?: string;
  warning?: string;
}
export interface LocalQueryRefreshAllResult {
  schema: 1;
  previewId: string;
  scope: LocalQueryRefreshAllScope;
  lastAcknowledgedState: string;
  cancelled: boolean;
  stopped: boolean;
  outcomes: LocalQueryRefreshAllOutcome[];
  accepted: number;
  uncertain: number;
  undoUnits: number | null;
}
