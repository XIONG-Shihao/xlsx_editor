import type { WorkbookSmartArtModel, WorkbookSmartArtLayout, WorkbookObjectAnchor } from './contracts';
export const SMARTART_NS: string;
export const SMARTART_CACHE_NS: string;
export const SMARTART_SAMPLES: readonly Readonly<{
  id: WorkbookSmartArtLayout;
  label: string;
  algorithm: string;
  labels: readonly string[];
}>[];
export const SMARTART_LIMITS: Readonly<{ nodes: number; text: number; partBytes: number; graphBytes: number }>;
export function normalizeSmartArt(value: unknown): WorkbookSmartArtModel;
export function smartartSampleModel(layout: WorkbookSmartArtLayout): WorkbookSmartArtModel;
export function smartartGeometry(value: WorkbookSmartArtModel): {
  width: number;
  height: number;
  boxes: Array<{ id: string; text: string; x: number; y: number; w: number; h: number }>;
  edges: Array<{ from: string; to: string; x1: number; y1: number; x2: number; y2: number }>;
};
export function smartartPartNames(id: string): Record<string, string>;
export function smartartEdgeIds(id: string): Record<string, string>;
export function smartartAnchorXml(id: string, number: number, anchor: WorkbookObjectAnchor, alt?: string): string;
export function smartartPackage(
  id: string,
  value: WorkbookSmartArtModel,
  anchor: WorkbookObjectAnchor
): {
  model: WorkbookSmartArtModel;
  names: Record<string, string>;
  parts: Record<string, string>;
  types: Record<string, string>;
};
export function smartartModelFromPackage(
  anchorNode: unknown,
  parts: Record<string, Uint8Array | string>,
  drawingPart: string,
  anchor: WorkbookObjectAnchor,
  id: string
): { model: WorkbookSmartArtModel; names: Record<string, string>; parts: Record<string, string> };
export function isSmartArtCommand(command: unknown): boolean;
export function captureSmartArtCommands<T>(commands: T[]): T[];
