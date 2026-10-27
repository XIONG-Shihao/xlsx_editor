import type { WorkbookObjectAnchor, WorkbookModel3DModel } from './contracts';
export const MODEL3D_NS: string;
export const MODEL3D_REL: string;
export const MODEL3D_MIME: string;
export function model3DEdgeId(id: string): string;
export function model3DAnchorXml(
  id: string,
  number: number,
  placement: WorkbookObjectAnchor,
  value: WorkbookModel3DModel,
  alt?: string
): string;
export function model3DFromAnchor(node: unknown, placement: WorkbookObjectAnchor, id: string): WorkbookModel3DModel;
export function isModel3DCommand(command: unknown): boolean;
export function captureModel3DCommands<T>(commands: T[]): T[];
