import { FORMAT } from './format.mjs';
/** Public XLSX/CSV component contracts; no application identity or routing types. */
/** Indexed CSV wire data. Kept independent of the parser and platform owners. */
export interface CsvDocumentView {
  kind: 'csv';
  sheet: 'CSV';
  // Raw LAN/worker views use a string; desktop decoding can supply a bounded
  // local source. Its methods are never sent through IPC or serialized to disk.
  text: string | { readonly length: number; slice(start?: number, end?: number): string };
  rowOffsets: Float64Array;
  rowWidths: Uint32Array;
  rows: number;
  columns: number;
  nonemptyCells: number;
  storedCells: number;
  delimiter: ',' | ';' | '\t';
  encoding: 'utf-8' | 'utf-16le' | 'utf-16be';
  bom: boolean;
  newline: string;
  terminalNewline: boolean;
  changes: Record<string, string>;
  widthOverrides: Record<string, number>;
  revision: string;
  sourceSnapshot: string;
}
export interface WorkbookCell {
  spill?: { schema: 1; master: string; range: string; role: 'master' | 'follower'; intendedRange?: string };
  /** Known formula output without an authored follower formula; grants no ownership. */
  formulaOutput?: { schema: 1; master: string; range: string };
  value: string | number | boolean | null;
  formula?: string;
  formulaType?: string;
  scenarioResult?: boolean;
  sharedFormulaId?: string;
  formulaReference?: string;
  dataType?: 'number' | 'text' | 'boolean' | 'blank' | 'error';
  rawValue?: string | null;
  cachedValue?: string | number | boolean | null;
  error?: string;
  calculation?: {
    status: 'calculated' | 'stored' | 'unresolved' | 'error' | 'circular';
    value?: string | number | boolean;
    error?: string;
    revision?: string;
    pending?: boolean;
  };
  cachedResult?: {
    authority: 'stored';
    status: 'calculated' | 'error' | 'unresolved' | 'circular';
    value?: string | number | boolean;
    error?: string;
    revision?: string;
    rawValue?: string;
  };
  styleId?: number;
  format?: string | number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  border?: 'none' | 'all';
}
export interface WorkbookFilterValuePage {
  values: Array<{ key: string; value: string | number | boolean | null; count: number }>;
  total: number;
  nextOffset: number | null;
  protected: { errors: number; unresolved: number; special: number; unsafeNumbers: number };
}
/** Object queries/commands carry semantic IDs and bounded inline PNG bytes;
 * package and filesystem paths remain private to the indexed owner. */
export interface WorkbookObjectAnchor {
  kind: 'oneCell';
  cell: string;
  dxEmu: number;
  dyEmu: number;
  widthEmu: number;
  heightEmu: number;
}
export type WorkbookObjectPlacement = Partial<WorkbookObjectAnchor> & { cell?: string };
export type WorkbookChartStyle = 'colorful' | 'monochrome' | 'outline';
export interface WorkbookChartModel {
  style?: WorkbookChartStyle;
  kind: 'column' | 'bar' | 'line';
  range: string;
  title?: string;
  legend?: 'bottom' | 'none';
  blanks?: 'gap';
  includeHidden?: false;
}
export interface WorkbookPngAsset {
  mime: 'image/png';
  base64: string;
  sha256?: string;
  bytes?: number;
  widthPx?: number;
  heightPx?: number;
}
export type WorkbookShapeSample =
  | 'text-plain'
  | 'text-note'
  | 'shape-rectangle'
  | 'shape-ellipse'
  | 'shape-arrow'
  | 'icon-check'
  | 'icon-star'
  | 'icon-plus'
  | 'wordart-blue'
  | 'wordart-outline'
  | 'wordart-shadow';
export interface WorkbookShapeModel {
  sample: WorkbookShapeSample;
  text: string;
  color?: string;
}
export type WorkbookInkStrokeModel = import('./ink-stroke.mjs').InkStrokeModel;
export type WorkbookObjectCommand =
  | {
      operation: 'createInkStroke';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      stroke: WorkbookInkStrokeModel;
      anchor: WorkbookObjectAnchor;
      altText?: string;
    }
  | {
      operation: 'updateInkStroke';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | { operation: 'deleteInkStroke'; sheet: string; objectId: string; expectedObjectRevision: string }
  | {
      operation: 'createModel3D';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      model3d: WorkbookModel3DModel;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | {
      operation: 'updateModel3D';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      patch?: Partial<WorkbookModel3DModel>;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | { operation: 'deleteModel3D'; sheet: string; objectId: string; expectedObjectRevision: string }
  | {
      operation: 'createSmartArt';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      smartart: WorkbookSmartArtModel;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | {
      operation: 'updateSmartArt';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      patch?: Partial<WorkbookSmartArtModel>;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | { operation: 'deleteSmartArt'; sheet: string; objectId: string; expectedObjectRevision: string }
  | {
      operation: 'createShape';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      shape: WorkbookShapeModel;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | {
      operation: 'updateShape';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      patch: Partial<WorkbookShapeModel>;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | { operation: 'deleteShape'; sheet: string; objectId: string; expectedObjectRevision: string }
  | {
      operation: 'createChart';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      chart: WorkbookChartModel;
      anchor?: WorkbookObjectPlacement;
    }
  | {
      operation: 'updateChart';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      patch: Partial<WorkbookChartModel>;
      anchor?: WorkbookObjectPlacement;
    }
  | {
      operation: 'insertImage';
      sheet: string;
      objectId: string;
      expectedObjectRevision: null;
      asset: WorkbookPngAsset;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | {
      operation: 'updateImage';
      sheet: string;
      objectId: string;
      expectedObjectRevision: string;
      anchor?: WorkbookObjectPlacement;
      altText?: string;
    }
  | { operation: 'deleteObject'; sheet: string; objectId: string; expectedObjectRevision: string };
export type WorkbookModel3DSample = 'cube' | 'pyramid' | 'octahedron';
export interface WorkbookModel3DModel {
  sample: WorkbookModel3DSample;
  color: string;
  rotation: { x: number; y: number; z: number };
}
export interface WorkbookModel3DMesh {
  positions: Array<[number, number, number]>;
  normals: Array<[number, number, number]>;
  triangles: Array<[number, number, number]>;
}
export interface WorkbookModel3DAsset {
  model: WorkbookModel3DModel;
  mesh: WorkbookModel3DMesh;
}
export type WorkbookSmartArtLayout = 'process' | 'hierarchy' | 'cycle';
export interface WorkbookSmartArtModel {
  layout: WorkbookSmartArtLayout;
  nodes: Array<{ id: string; text: string }>;
  connections: Array<{ from: string; to: string }>;
}
export interface WorkbookObjectSummary {
  id: string;
  name: string;
  kind: 'chart' | 'image' | 'shape' | 'smartart' | 'model3d' | 'ink' | 'opaque';
  revision: string;
  anchor: WorkbookObjectAnchor | null;
  managed: boolean;
  protectedReason: string | null;
  canEdit: boolean;
  canDelete: boolean;
  model3d?: WorkbookModel3DModel;
  smartart?: WorkbookSmartArtModel;
  stroke?: WorkbookInkStrokeModel;
  shape?: WorkbookShapeModel;
  chart?: WorkbookChartModel;
  altText?: string;
}
export interface WorkbookObjectPage {
  objects: WorkbookObjectSummary[];
  total: number;
  nextCursor: string | null;
}
export interface WorkbookObjectPreview extends WorkbookObjectSummary {
  asset3d?: WorkbookModel3DAsset;
  data?: { categories: string[]; series: Array<{ name: string; values: Array<number | null> }>; hidden: boolean[] };
  image?: WorkbookPngAsset & { sha256: string; bytes: number; widthPx: number; heightPx: number };
}
/** First print slice: one rectangle, <=10k source positions, no device paths. */
export interface WorkbookPrintTextSections {
  left: string;
  center: string;
  right: string;
}
export interface WorkbookPrintHeaderFooter {
  header: WorkbookPrintTextSections;
  footer: WorkbookPrintTextSections;
  scaleWithDoc: boolean;
}
export interface WorkbookPrintSettings {
  area: string | null;
  paper: 'A4' | 'Letter';
  orientation: 'portrait' | 'landscape';
  marginsPt: { top: number; right: number; bottom: number; left: number; header: number; footer: number };
  scaling: { mode: 'scale'; percent: number } | { mode: 'fit'; width: number; height: number };
  repeatRows: { first: number; last: number } | null;
  repeatColumns: { first: string; last: string } | null;
  gridlines: boolean;
  headerFooter?: WorkbookPrintHeaderFooter;
}
export interface WorkbookPrintDefinition {
  sheet: string;
  revision: string | null;
  canEdit: boolean;
  settings: WorkbookPrintSettings | null;
  protectedReasons: string[];
}
export interface WorkbookNormalFont {
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
}
/** Local rendering dimensions; this profile never authorizes cell/cache values. */
export interface WorkbookGeometryProfile {
  schema: 1;
  font: WorkbookNormalFont;
  maximumDigitWidth: number;
  source: 'local-font' | 'substituted-font' | 'reference-fallback';
  availability: 'confirmed' | 'unverified' | 'substituted';
  renderedFontFamily?: string;
}
export type WorkbookPrintStyle = Pick<
  WorkbookCell,
  | 'bold'
  | 'italic'
  | 'underline'
  | 'fontFamily'
  | 'fontSize'
  | 'fontColor'
  | 'fillColor'
  | 'horizontalAlign'
  | 'verticalAlign'
  | 'wrapText'
  | 'border'
>;
export interface WorkbookPrintSnapshot {
  schema: 1;
  sessionId: string;
  state: string;
  revision: string;
  sheet: string;
  date1904: boolean;
  settings: WorkbookPrintSettings;
  definitionRevision: string | null;
  fingerprint: string;
  geometryProfile?: WorkbookGeometryProfile;
  rows: Array<{ index: number; heightPt: number; hidden: boolean }>;
  columns: Array<{ index: number; widthPt: number; hidden: boolean }>;
  merges: string[];
  cells: Array<{ ref: string; text: string; style: WorkbookPrintStyle }>;
  objects: Array<WorkbookObjectPreview & { boxPt: [number, number, number, number] }>;
}
export interface WorkbookPrintPage {
  number: number;
  rowIndices: number[];
  columnIndices: number[];
  cells: Array<{
    ref: string;
    text: string;
    style: WorkbookPrintStyle;
    bbox: [number, number, number, number];
    merge?: string;
  }>;
  objects: Array<{
    id: string;
    kind: 'chart' | 'image';
    bbox: [number, number, number, number];
    clipBox: [number, number, number, number];
    clipped: boolean;
    object: WorkbookObjectPreview;
  }>;
}
export interface WorkbookPrintBoundary {
  number: number;
  rows: [number, number];
  columns: [number, number];
  bodyRows: [number, number] | null;
  bodyColumns: [number, number] | null;
}
export interface WorkbookPrintLayout {
  units: 'pt';
  pageSizePt: [number, number];
  scale: number;
  sourceFingerprint: string;
  revision: string;
  sheet: string;
  settings: WorkbookPrintSettings;
  geometryProfile?: WorkbookGeometryProfile;
  pages: WorkbookPrintPage[];
  boundaries?: WorkbookPrintBoundary[];
}
export type WorkbookPrintCommand = {
  operation: 'setPrintSettings';
  sheet: string;
  expectedPrintRevision: string | null;
  settings: WorkbookPrintSettings;
};
export type WorkbookConditionalRule = {
  range: string;
  kind: 'greaterThan' | 'lessThan' | 'equal' | 'between' | 'textContains' | 'duplicateValues';
  values?: number[];
  text?: string;
  fillColor: string;
  stopIfTrue: boolean;
};
export type WorkbookValidationRule = {
  range: string;
  kind: 'list' | 'wholeNumber' | 'decimal' | 'textLength';
  values?: string[];
  minimum?: number;
  maximum?: number;
  allowBlank: boolean;
  showDropdown: boolean;
  promptTitle: string;
  prompt: string;
  errorTitle: string;
  error: string;
};
export type WorkbookRuleDefinition<T> = {
  sheet: string;
  revision: string | null;
  canEdit: boolean;
  protectedReasons: string[];
  protectedRanges?: string[];
  evaluationReasons?: string[];
  rules: T[];
};
export type WorkbookConditionalDefinition = WorkbookRuleDefinition<WorkbookConditionalRule> & {
  duplicates?: Record<string, Record<string, number>>;
};
export type WorkbookValidationDefinition = WorkbookRuleDefinition<WorkbookValidationRule>;
export type WorkbookRulesCommand =
  | {
      operation: 'setConditionalRules';
      sheet: string;
      expectedRulesRevision: string | null;
      rules: WorkbookConditionalRule[];
    }
  | {
      operation: 'setValidationRules';
      sheet: string;
      expectedRulesRevision: string | null;
      rules: WorkbookValidationRule[];
    };
export interface WorkbookFilterDefinition {
  range: string;
  column?: string;
  values?: Array<string | number | boolean | null>;
  owned: boolean;
  revision: string | null;
  canReapply: boolean;
  canClear: boolean;
  reason?: string;
}
export type WorkbookBackgroundCommand =
  | {
      type: 'setSheetBackground';
      sheet: string;
      sample: 'graph-paper' | 'soft-dots';
      expectedBackgroundRevision: string | null;
    }
  | { type: 'removeSheetBackground'; sheet: string; expectedBackgroundRevision: string | null };
export interface WorkbookSheet {
  background?: {
    schema: typeof FORMAT.SCHEMA_SHEET_BACKGROUND_V1;
    sheet: string;
    revision: string | null;
    sample: string | null;
    canEdit: boolean;
    protectedReason: string | null;
  };
  /** Known primary source view defaults; local UI overrides never serialize here. */
  sourceView?: { zoom?: number; gridlines?: boolean; headings?: boolean };
  /** Accepted worksheet defaults; absent attributes retain their source absence. */
  sheetFormat?: {
    defaultRowHeight?: number;
    defaultColumnWidth?: number;
    baseColumnWidth?: number;
    rowsHiddenByDefault: boolean;
  };
  normalFont?: WorkbookNormalFont;
  geometryIssues?: string[];
  storedCells?: number;
  duplicateValues?: Record<string, string[]>;
  storedFormulaCount?: number;
  calculationRevision?: string;
  date1904?: boolean;
  name: string;
  state: string;
  path: string;
  cells: Record<string, WorkbookCell>;
  rows: number;
  columns: number;
  merges: string[];
  conditionalRules?: Array<{
    range: string;
    rule: WorkbookConditionalRule['kind'];
    threshold?: number;
    values?: number[];
    text?: string;
    fillColor: string;
    stopIfTrue?: boolean;
    priority?: number;
  }>;
  validationRules?: WorkbookValidationRule[];
  tableRanges?: Array<{ name: string; ref: string }>;
  widths: Array<{ min: number; max: number; width: number; hidden: boolean }>;
  hiddenRows?: number[];
  /** Explicit visible-row overrides for hidden-by-default worksheets. */
  visibleRows?: number[];
  rowHeights?: Array<{ row: number; height: number }>;
  rowOrder?: Array<{ from: number; to: number }>;
  capabilities?: {
    referenceColumns?: boolean;
    deleteRows?: boolean;
    insertRows: boolean;
    insertRowsReason: string | null;
  };
  filter?: WorkbookFilterDefinition;
  freeze?: { rows: number; columns: number };
}
export interface WorkbookStorage {
  kind: 'rust-indexed-v1' | 'csv-indexed-v1';
  sessionId: string;
  state: string;
}
export type WorkbookPageThemeCommand = {
  type: 'applyPageTheme';
  sample: 'neutral' | 'forest' | 'plum';
  expectedThemeRevision: string | null;
};
export interface WorkbookView {
  pageTheme?: { revision: string | null; sample: string | null; canApply: boolean; protectedReason: string | null };
  storage?: WorkbookStorage;
  /** Exact bytes for this view, supplied by desktop transport for recovery/copy. */
  sourceData?: string;
  sourceSnapshot?: string;
  revision: string;
  /** Passive workbook formats retain original bytes and saved formula caches. */
  format?: 'xlsx' | 'csv' | 'tsv' | 'xlsm' | 'xls' | 'xlsb';
  originalFormat?: 'xlsm' | 'xls' | 'xlsb';
  passiveNotice?: string;
  pivots?: WorkbookPivotDefinition[];
  readOnly: boolean;
  limitations: string[];
  /** The package carries a macro project; it is kept untouched. */
  macroProject?: boolean;
  structuralReason?: string;
  calculation: string;
  calculationPolicy?: WorkbookCalculationPolicy;
  calculationEngine?: string;
  namedRanges?: Array<{ name: string; formula: string; sheet?: string; scopeInvalid?: true }>;
  tables?: Array<{
    name: string;
    sheet: string;
    ref: string;
    columns: Array<{ name: string }>;
    headerRowCount?: number;
    totalsRowCount?: number;
  }>;
  features?: Record<string, unknown>;
  capabilities?: Record<string, unknown>;
  sheets: WorkbookSheet[];
}
export interface WorkbookDuplicatePreview {
  schema: 1;
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
  sheet: string;
  range: string;
  hasHeaders: boolean;
  columns: string[];
  fingerprint: string;
  totalRows: number;
  keptRows: number;
  removedRows: number;
  comparison: 'exact-typed-values-v1';
  samples: Array<{ row: number; keptRow: number }>;
  offset: number;
  nextOffset: number | null;
}
export interface WorkbookRemoveDuplicatesCommand {
  operation: 'removeDuplicates';
  type?: undefined;
  sheet: string;
  range: string;
  hasHeaders: boolean;
  columns: string[];
  expectedFingerprint: string;
}
export interface WorkbookSortKey {
  column: string;
  direction: 'asc' | 'desc';
}
export type WorkbookBooleanCellCommand = {
  type: 'booleanCell';
  action: 'insert' | 'toggle' | 'remove';
  sheet: string;
  cell: string;
  expectedValue: boolean | null;
};
export type WorkbookHyperlinkDestination =
  | { kind: 'url'; target: string }
  | { kind: 'internal'; sheet: string; cell: string };
export interface WorkbookHyperlinkReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export interface WorkbookHyperlinkInner {
  schema: typeof FORMAT.SCHEMA_HYPERLINK_DEFINITION_V1;
  sheet: string;
  cell: string;
  hyperlinkRevision: string | null;
  link: WorkbookHyperlinkDestination | null;
  tooltip: string;
  canEdit: boolean;
  protectedReason: string | null;
}
export type WorkbookHyperlinkDefinition = WorkbookHyperlinkInner & WorkbookHyperlinkReceipt;
export type WorkbookHyperlinkCommand =
  | {
      type: 'setHyperlink';
      sheet: string;
      cell: string;
      expectedHyperlinkRevision: string | null;
      link: WorkbookHyperlinkDestination;
      tooltip?: string;
    }
  | { type: 'removeHyperlink'; sheet: string; cell: string; expectedHyperlinkRevision: string | null };

export interface WorkbookEdit {
  sheet: string;
  /** Indexed-owner recovery can retain a compact range format command. */
  type?:
    | 'setSheetBackground'
    | 'removeSheetBackground'
    | 'applyPageTheme'
    | 'formatRange'
    | 'clearRange'
    | 'fillRange'
    | 'fillSeries'
    | 'fillAcrossSheets'
    | 'clearHyperlinks'
    | 'setHyperlink'
    | 'removeHyperlink'
    | 'booleanCell'
    | 'pasteValues'
    | 'pasteFormulas';
  /** Native metadata geometry; never a cell/style projection. */
  ranges?: string[];
  source?: string;
  sourceSheet?: string;
  mode?: WorkbookFillAcrossSheetsMode;
  format?: Pick<
    WorkbookEdit,
    | 'bold'
    | 'italic'
    | 'underline'
    | 'fontFamily'
    | 'fontSize'
    | 'fontColor'
    | 'fillColor'
    | 'horizontalAlign'
    | 'verticalAlign'
    | 'wrapText'
    | 'border'
    | 'styleId'
    | 'numberFormat'
  >;
  cell?: string;
  value?: string | number | boolean | null;
  /** Only canonical typed spreadsheet errors; ordinary strings stay text. */
  valueType?: 'error';
  formula?: string;
  operation?:
    | 'insertRow'
    | 'deleteRow'
    | 'insertColumn'
    | 'deleteColumn'
    | 'sortRows'
    | 'filterRows'
    | 'reapplyFilter'
    | 'clearFilter'
    | 'removeFilter'
    | 'resizeColumn'
    | 'resizeRow'
    | 'freeze'
    | 'unhideAll'
    | 'addSheet'
    | 'renameSheet'
    | 'duplicateSheet'
    | 'mergeCells'
    | 'unmergeCells'
    | 'addConditionalRule'
    | 'createTable'
    | 'defineName'
    | 'removeName'
    | 'updateName'
    | 'createChart'
    | 'updateChart'
    | 'insertImage'
    | 'updateImage'
    | 'deleteObject'
    | 'setPrintSettings'
    | 'setConditionalRules'
    | 'setValidationRules';
  expectedRulesRevision?: string | null;
  rules?: WorkbookConditionalRule[] | WorkbookValidationRule[];
  expectedPrintRevision?: string | null;
  settings?: WorkbookPrintSettings;
  geometryProfile?: WorkbookGeometryProfile;
  objectId?: string;
  expectedObjectRevision?: string | null;
  chart?: WorkbookChartModel;
  patch?: Partial<WorkbookChartModel>;
  anchor?: WorkbookObjectPlacement;
  asset?: WorkbookPngAsset;
  altText?: string;
  name?: string;
  nameScope?: 'workbook' | 'worksheet';
  expectedNameFormula?: string | null;
  tableName?: string;
  rule?: 'greaterThan' | 'lessThan' | 'duplicateValues';
  threshold?: number;
  newName?: string;
  range?: string;
  column?: string;
  direction?: 'asc' | 'desc';
  /** Ordered 1–16 distinct in-range sort keys; excludes column/direction. */
  keys?: WorkbookSortKey[];
  values?: Array<string | number | boolean | null>;
  expectedFilterRevision?: string;
  width?: number;
  height?: number;
  rows?: number;
  columns?: number;
  count?: number;
  expectedValue?: string | number | boolean | null;
  expectedFormula?: string | null;
  expectedRawValue?: string | null;
  index?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  fillColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'center' | 'bottom';
  wrapText?: boolean;
  border?: 'none' | 'all';
  styleId?: number;
  numberFormat?:
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
}

/** Compact replay/history may retain a strict staged descriptor. Its location
 * is shape.sheet/start, never a fabricated top-level sheet/cell edit. */
export type WorkbookFillAcrossSheetsMode = 'all' | 'contents' | 'formats';
/** Copies the same address rectangle from one visible worksheet to another. */
export interface WorkbookFillAcrossSheetsCommand extends WorkbookEdit {
  type: 'fillAcrossSheets';
  sheet: string;
  sourceSheet: string;
  range: string;
  mode: WorkbookFillAcrossSheetsMode;
}
export interface WorkbookClearHyperlinksCommand extends WorkbookEdit {
  type: 'clearHyperlinks';
  sheet: string;
  ranges: string[];
}
export type WorkbookHistoryEdit =
  | WorkbookClearHyperlinksCommand
  | WorkbookFillAcrossSheetsCommand
  | WorkbookRemoveDuplicatesCommand
  | WorkbookCalculationCommand
  | WorkbookRecalculateCommand
  | WorkbookEdit
  | import('./paste-protocol.mjs').PasteDescriptor;
/** Direct syntax/owned-spill relationships, never an exhaustive Excel graph. */
export interface WorkbookTracePosition {
  sheet: string;
  range: string;
}
export interface WorkbookTraceEdge {
  from: WorkbookTracePosition;
  to: WorkbookTracePosition;
  kind: 'reference' | 'spill-origin' | 'spill-range';
  via: 'formula' | 'name' | 'table';
}
export interface WorkbookTraceMarker {
  sheet: string;
  cell?: string;
  kind: 'dynamic' | 'external' | 'unsupported' | 'special' | 'limit';
  reason: string;
}
export interface WorkbookFormulaTrace {
  schema: 1;
  documentId?: string;
  ownerId: string;
  state: string;
  revision: string;
  sheet: string;
  cell: string;
  direction: 'precedents' | 'dependents';
  edges: WorkbookTraceEdge[];
  markers: WorkbookTraceMarker[];
  nextCursor: string | null;
  coverage: {
    mode: 'direct-static-references';
    complete: boolean;
    scannedFormulas: number;
    pages: number;
    limited: boolean;
  };
}
export interface WorkbookTraceOptions {
  direction: 'precedents' | 'dependents';
  after?: string;
  limit?: number;
}
export type WorkbookTableCommand =
  | { operation: 'renameTable'; sheet: string; tableName: string; expectedTableRef: string; newName: string }
  | { operation: 'resizeTable'; sheet: string; tableName: string; expectedTableRef: string; range: string }
  | {
      operation: 'renameTableColumn';
      sheet: string;
      tableName: string;
      expectedTableRef: string;
      columnName: string;
      newName: string;
    }
  | {
      operation: 'insertTableColumn';
      sheet: string;
      tableName: string;
      expectedTableRef: string;
      index: number;
      newName: string;
    }
  | { operation: 'deleteTableColumn'; sheet: string; tableName: string; expectedTableRef: string; columnName: string };
export type WorkbookPivotLayoutSample = 'tabular-light' | 'compact-blue' | 'outline-dark';
export type WorkbookPivotCommand =
  | {
      operation: 'createPivot';
      sheet: string;
      name: string;
      range: string;
      rowField: number;
      valueField: number;
      aggregate: 'sum' | 'count';
      pivotSheet: string;
      expectedPivotRevision: null;
    }
  | { operation: 'refreshPivot'; sheet: string; name: string; expectedPivotRevision: string }
  | {
      operation: 'updatePivot';
      sheet: string;
      name: string;
      sourceSheet: string;
      range: string;
      rowField: number;
      valueField: number;
      aggregate: 'sum' | 'count';
      expectedPivotRevision: string;
    };
export type WorkbookPivotLayoutCommand = {
  operation: 'setPivotLayout';
  sheet: string;
  name: string;
  sample: WorkbookPivotLayoutSample;
  grandTotals: boolean;
  expectedPivotRevision: string;
};
export interface WorkbookPivotDefinition {
  name: string;
  sheet: string | null;
  sourceSheet?: string;
  sourceRange?: string;
  rowField?: number;
  valueField?: number;
  aggregate?: 'sum' | 'count';
  fieldNames?: string[];
  sample?: WorkbookPivotLayoutSample;
  grandTotals?: boolean;
  revision: string | null;
  outputRange?: string;
  canEdit: boolean;
  reason?: string;
}
/** Explicit full recalculation of the captured workbook, independent of its persisted scheduling mode. */
export interface WorkbookRecalculateCommand {
  type: 'recalculate';
}
export interface WorkbookCalculationPolicy {
  schema: 1;
  mode: 'auto' | 'manual';
  recalculateBeforeSave: boolean;
  pendingFormulaCells: number;
  pendingCacheObjects?: number;
  protectedSource?: boolean;
  preserveSource?: boolean;
  canRecalculate?: boolean;
  canCalculateSheet: boolean;
  canChangeOptions: boolean;
  reason?: string | null;
}
/** A private draft result never supplies accepted-cell calculation authority. */
export interface WorkbookFormulaPreviewReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export type WorkbookFormulaPreviewResult =
  | { status: 'value'; type: 'number'; value: number }
  | { status: 'value'; type: 'text' | 'error'; value: string }
  | { status: 'value'; type: 'boolean'; value: boolean }
  | { status: 'invalid' | 'unavailable' | 'cancelled'; reason: string };
export interface WorkbookFormulaPreview {
  schema: 1;
  requestId: string;
  receipt: WorkbookFormulaPreviewReceipt;
  sheet: string;
  cell: string;
  formula: string;
  authority: 'draft-preview';
  result: WorkbookFormulaPreviewResult;
}
export interface WorkbookFormulaPreviewCancellation {
  schema: 1;
  requestId: string;
  cancelled: boolean;
}
/** Complete source/effective geometry; padded display dimensions are never a basis. */
export interface WorkbookUsedRangeBounds {
  minRow: number;
  minColumn: number;
  maxRow: number;
  maxColumn: number;
}
export interface WorkbookUsedRangeReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export interface WorkbookUsedRangeNativeIdentity {
  version: number;
  sourceRevision: string;
  sheetPath: string;
}
export interface WorkbookUsedRangeProvenance {
  factsSchema: typeof FORMAT.SCHEMA_USED_RANGE_FACTS_V1 | typeof FORMAT.SCHEMA_USED_RANGE_FACTS_V2;
  declaration: {
    status: 'absent' | 'valid' | 'invalid' | 'duplicate';
    ref: string | null;
    bounds: WorkbookUsedRangeBounds | null;
  };
  sourceCells: WorkbookUsedRangeBounds | null;
  sourceMerges: WorkbookUsedRangeBounds | null;
  effectiveCells: WorkbookUsedRangeBounds | null;
  sourceAxisFormatting: WorkbookUsedRangeBounds | null;
  acceptedHistory: WorkbookUsedRangeBounds | null;
  sourceDefaultFormatting: boolean;
}
export interface WorkbookUsedRangeReason {
  code: string;
  message: string;
}
export type WorkbookUsedRangeInner = {
  schema: 1;
  policy: 'xlsx-used-range-v1';
  sheet: string;
  native: WorkbookUsedRangeNativeIdentity;
  provenance: WorkbookUsedRangeProvenance;
  reasons: WorkbookUsedRangeReason[];
} & (
  | { status: 'complete'; scopeMinMax: WorkbookUsedRangeBounds; range: string; lastCell: string; empty: boolean }
  | { status: 'incomplete'; scopeMinMax: null; range: null; lastCell: null; empty: null }
);
export type WorkbookUsedRangeResult = WorkbookUsedRangeInner & { receipt: WorkbookUsedRangeReceipt };
export interface WorkbookUsedRangeControls {
  readonly receipt: WorkbookUsedRangeReceipt;
  read(sheet: string): Promise<WorkbookUsedRangeResult>;
}
/** Source membership is independent of cell records, display geometry, and edit capability. */
export interface WorkbookNoteMembershipSource {
  version: number;
  sourceRevision: string;
  sheetPath: string;
}
export interface WorkbookNoteMembershipReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export interface WorkbookNoteMembershipReason {
  code: string;
  message: string;
}
export type WorkbookNoteMembershipInner = {
  schema: 1;
  policy: 'xlsx-legacy-note-membership-v1';
  sheet: string;
  source: WorkbookNoteMembershipSource;
  reasons: WorkbookNoteMembershipReason[];
} & (
  | { status: 'complete'; notes: string[]; threadedPlaceholders: string[] }
  | { status: 'incomplete'; notes: null; threadedPlaceholders: null }
);
export type WorkbookNoteMembershipResult = WorkbookNoteMembershipInner & { receipt: WorkbookNoteMembershipReceipt };
export interface WorkbookNoteMembershipControls {
  readonly receipt: WorkbookNoteMembershipReceipt;
  read(sheet: string): Promise<WorkbookNoteMembershipResult>;
}
export type WorkbookCalculationCommand =
  | { type: 'setCalculationOptions'; mode: 'auto' | 'manual'; recalculateBeforeSave?: boolean }
  | { type: 'calculateSheet'; sheet: string };
export interface WorkbookAccessibilityControls {
  receipt: import('./notes.mjs').NoteReceipt;
  read: () => Promise<import('./accessibility.mjs').AccessibilityReport & import('./notes.mjs').NoteReceipt>;
}
export interface WorkbookNotesControls {
  receipt: import('./notes.mjs').NoteReceipt;
  read: (sheet: string) => Promise<WorkbookNotes>;
}
export type WorkbookNoteCommand = import('./notes.mjs').NoteCommand;
export type WorkbookNotes = import('./notes.mjs').NoteInventory & import('./notes.mjs').NoteReceipt;
export type WorkbookCommand =
  | WorkbookNoteCommand
  | WorkbookBackgroundCommand
  | WorkbookPageThemeCommand
  | WorkbookBooleanCellCommand
  | WorkbookHyperlinkCommand
  | WorkbookCalculationCommand
  | WorkbookHistoryEdit
  | WorkbookRecalculateCommand
  | WorkbookObjectCommand
  | WorkbookPrintCommand
  | WorkbookRulesCommand
  | WorkbookTableCommand
  | WorkbookPivotCommand
  | WorkbookPivotLayoutCommand
  | { type: 'clearRange' | 'fillSeries'; sheet: string; range: string }
  | { type: 'pasteValues'; sheet: string; sourceSheet?: string; source: string; range: string }
  | { type: 'setCells'; sheet: string; cells: Array<Omit<WorkbookEdit, 'sheet'>> }
  | { type: 'fillRange' | 'pasteFormulas'; sheet: string; source: string; range: string }
  | {
      type: 'formatRange';
      sheet: string;
      range: string;
      format: Pick<
        WorkbookEdit,
        | 'bold'
        | 'italic'
        | 'underline'
        | 'fontFamily'
        | 'fontSize'
        | 'fontColor'
        | 'fillColor'
        | 'horizontalAlign'
        | 'verticalAlign'
        | 'wrapText'
        | 'border'
        | 'styleId'
        | 'numberFormat'
      >;
    };
