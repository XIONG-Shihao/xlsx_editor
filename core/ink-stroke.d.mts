export type InkPenId = 'black' | 'red' | 'yellow-highlighter';
export interface InkStrokeModel {
  schema: 1;
  pen: InkPenId;
  widthEmu: number;
  heightEmu: number;
  points: [number, number][];
}
export interface InkPen {
  readonly id: InkPenId;
  readonly label: string;
  readonly color: string;
  readonly widthEmu: number;
  readonly alpha: number;
  readonly cap: 'round' | 'flat';
}
export const INK_EMU_PER_PIXEL: 9525;
export const INK_LIMITS: Readonly<{
  points: number;
  inputPoints: number;
  extentEmu: number;
  offsetEmu: number;
  alt: number;
  xmlBytes: number;
}>;
export const INK_PENS: readonly InkPen[];
export function inkPen(id: unknown): InkPen;
export function normalizeInkStroke(value: unknown): InkStrokeModel;
export function inkStrokeFromLogicalPoints(
  input: unknown,
  penId?: InkPenId
): { stroke: InkStrokeModel; originEmu: { x: number; y: number } };
export function inkStrokePath(value: InkStrokeModel): string;
export function inkStrokeHit(value: InkStrokeModel, x: number, y: number, toleranceEmu?: number): boolean;

export function isInkStrokeCommand(value: unknown): boolean;
export function captureInkStrokeCommands<T>(commands: T[]): T[];
