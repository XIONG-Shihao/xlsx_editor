import type {
  WorkbookShapeModel,
  WorkbookShapeSample,
  WorkbookObjectPlacement,
  WorkbookObjectCommand,
} from './contracts';
export interface ShapeSample {
  readonly id: WorkbookShapeSample;
  readonly family: 'text' | 'shape' | 'icon' | 'wordart';
  readonly label: string;
  readonly geometry: 'rect' | 'ellipse' | 'rightArrow' | 'custom';
  readonly fill: string | null;
  readonly line: string | null;
  readonly color: string;
  readonly fontSize: number;
  readonly align: 'l' | 'ctr';
  readonly textBox: boolean;
  readonly outline?: string | null;
  readonly shadow?: boolean;
  readonly asset?: string;
  readonly points?: readonly (readonly [number, number])[];
}
export const SHAPE_SAMPLES: readonly ShapeSample[];
export const ILLUSTRATION_SAMPLES: readonly ShapeSample[];
export const SHAPE_LIMITS: Readonly<{ text: number; paragraphs: number; alt: number }>;
export const SHAPE_DRAW_NS: string;
export const SHAPE_MAIN_NS: string;
export function shapeSample(id: string): ShapeSample;
export function normalizeShape(value: unknown): WorkbookShapeModel;
export function shapeAnchorXml(
  id: string,
  number: number,
  shape: WorkbookShapeModel,
  placement: Required<WorkbookObjectPlacement>,
  altText?: string
): string;
export function shapeModelFromAnchor(node: unknown, placement: Required<WorkbookObjectPlacement>): WorkbookShapeModel;
export function isShapeCommand(command: unknown): boolean;
export function captureShapeCommands(commands: WorkbookObjectCommand[]): WorkbookObjectCommand[];
