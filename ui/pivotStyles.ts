/** Finite viewport sample projection from a qualified owned pivot definition. */
import type { WorkbookPivotDefinition, WorkbookPivotLayoutSample } from '../core/contracts';
import { selectedPivot } from './pivotSelection';
export interface PivotSamplePalette {
  header: string;
  headerText: string;
  band: string;
  text: string;
  striped: boolean;
  compact: boolean;
}
const palettes: Readonly<Record<WorkbookPivotLayoutSample, PivotSamplePalette>> = Object.freeze({
  'tabular-light': Object.freeze({
    header: '#DCE6F1',
    headerText: '#17365D',
    band: '#FFFFFF',
    text: '#1F2937',
    striped: false,
    compact: false,
  }),
  'compact-blue': Object.freeze({
    header: '#4472C4',
    headerText: '#FFFFFF',
    band: '#E9EFF7',
    text: '#1F2937',
    striped: true,
    compact: true,
  }),
  'outline-dark': Object.freeze({
    header: '#333333',
    headerText: '#FFFFFF',
    band: '#F2F2F2',
    text: '#222222',
    striped: true,
    compact: false,
  }),
});
export function pivotSamplePalette(sample: string | undefined): PivotSamplePalette | undefined {
  return sample && Object.hasOwn(palettes, sample) ? palettes[sample as WorkbookPivotLayoutSample] : undefined;
}
export function pivotCellStyle(
  pivots: readonly WorkbookPivotDefinition[],
  sheet: string,
  ref: string
):
  | {
      sample: WorkbookPivotLayoutSample;
      role: 'header' | 'body' | 'total';
      background: string;
      color: string;
      bold: boolean;
      paddingLeft?: number;
    }
  | undefined {
  const p = selectedPivot(pivots, sheet, ref),
    palette = pivotSamplePalette(p?.sample),
    output = /^A1:B([1-9]\d*)$/.exec(p?.outputRange ?? ''),
    cell = /^([AB])([1-9]\d*)$/.exec(ref);
  if (
    !p?.canEdit ||
    !p.revision ||
    !p.sample ||
    !palette ||
    typeof p.grandTotals !== 'boolean' ||
    !output ||
    !cell ||
    Number(output[1]) < 2 ||
    Number(output[1]) > 514
  )
    return undefined;
  const row = Number(cell[2]),
    role = row === 1 ? 'header' : p.grandTotals && row === Number(output[1]) ? 'total' : 'body',
    bold = role !== 'body';
  return {
    sample: p.sample,
    role,
    background: bold ? palette.header : palette.striped && row % 2 === 0 ? palette.band : '#FFFFFF',
    color: bold ? palette.headerText : palette.text,
    bold,
    ...(palette.compact && cell[1] === 'A' && role === 'body' ? { paddingLeft: 12 } : {}),
  };
}
