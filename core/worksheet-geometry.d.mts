import type { WorkbookNormalFont, WorkbookGeometryProfile } from './contracts';
export type GeometryFont = WorkbookNormalFont;
export type GeometryProfile = WorkbookGeometryProfile;
export interface WorksheetGeometrySource {
  sheetFormat?: {
    defaultRowHeight?: number;
    defaultColumnWidth?: number;
    baseColumnWidth?: number;
    rowsHiddenByDefault: boolean;
  };
  normalFont?: GeometryFont;
  rowHeights?: Array<{ row: number; height: number }>;
  hiddenRows?: number[];
  visibleRows?: number[];
  widths?: Array<{ min: number; max: number; width: number; hidden: boolean }>;
}
export interface WorksheetDimensions {
  profile: GeometryProfile;
  font: GeometryFont;
  defaultRowHeight: number;
  defaultColumnWidth: number;
  defaultColumnPixels: number;
  rowHidden(row: number): boolean;
  rowHeightPt(row: number): number;
  columnPixels(col: number): number;
  rowDistancePt(first: number, last: number): number;
}
export const WORKSHEET_MAX_ROWS: number, WORKSHEET_MAX_COLUMNS: number, PIXELS_PER_POINT: number;
export const DEFAULT_NORMAL_FONT: Readonly<GeometryFont>;
export function normalGeometryFont(sheet?: WorksheetGeometrySource): Readonly<GeometryFont>;
export function validateGeometryProfile(value: unknown, expectedFont?: GeometryFont): Readonly<GeometryProfile>;
export function referenceGeometryProfile(font?: GeometryFont): Readonly<GeometryProfile>;
export function xmlColumnWidthPixels(width: number, maximumDigitWidth: number): number;
export function encodedColumnWidth(characters: number, maximumDigitWidth: number): number;
export function worksheetDefaultRowHeight(sheet?: WorksheetGeometrySource): number;
export function worksheetDefaultColumnWidth(
  sheet: WorksheetGeometrySource | undefined,
  profile?: GeometryProfile
): number;
export function createWorksheetDimensions(
  sheet?: WorksheetGeometrySource,
  profile?: GeometryProfile
): WorksheetDimensions;
export function worksheetRowOffsets(
  sheet: WorksheetGeometrySource | undefined,
  count: number,
  profile?: GeometryProfile
): Float64Array;
export function worksheetColumnOffsets(
  sheet: WorksheetGeometrySource | undefined,
  count: number,
  profile?: GeometryProfile
): Float64Array;

export function rowHeightPoints(sheet: WorksheetGeometrySource | undefined, row: number): number;
export function columnWidthPixels(
  sheet: WorksheetGeometrySource | undefined,
  col: number,
  profile?: GeometryProfile
): number;
export function rowOffsetPixels(sheet: WorksheetGeometrySource | undefined, row: number): number;
export function columnOffsetPixels(
  sheet: WorksheetGeometrySource | undefined,
  col: number,
  profile?: GeometryProfile
): number;
