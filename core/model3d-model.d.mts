import type { WorkbookModel3DSample, WorkbookModel3DModel, WorkbookModel3DAsset } from './contracts';
export const MODEL3D_LIMITS: Readonly<{ assetBytes: number; vertices: number; triangles: number; coordinate: number }>;
export const MODEL3D_SAMPLES: readonly Readonly<{
  id: WorkbookModel3DSample;
  label: string;
  positions: readonly (readonly number[])[];
  faces: readonly (readonly number[])[];
}>[];
export function normalizeModel3D(value: unknown): WorkbookModel3DModel;
export function model3DGlb(value: unknown): Uint8Array;
export function parseModel3DGlb(bytes: Uint8Array): WorkbookModel3DAsset;
export function model3DSampleAsset(sample: WorkbookModel3DSample): WorkbookModel3DAsset;
export function projectModel3D(
  asset: WorkbookModel3DAsset,
  width?: number,
  height?: number
): {
  width: number;
  height: number;
  faces: Array<{ index: number; points: Array<[number, number]>; depth: number; color: string }>;
};
