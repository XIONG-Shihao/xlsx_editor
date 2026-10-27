export type TextConversionPolicy = 'text' | 'number' | 'boolean' | 'iso-date';
export interface TextConversionOptions {
  date1904?: boolean;
  padded?: boolean;
}
export type TextConversionResult =
  | Readonly<{
      status: 'converted';
      kind: 'text';
      value: string;
      notices: readonly string[];
      empty: boolean;
      padded: boolean;
    }>
  | Readonly<{
      status: 'converted';
      kind: 'number';
      value: number;
      notices: readonly string[];
      empty: boolean;
      padded: boolean;
    }>
  | Readonly<{
      status: 'converted';
      kind: 'boolean';
      value: boolean;
      notices: readonly string[];
      empty: boolean;
      padded: boolean;
    }>
  | Readonly<{
      status: 'converted';
      kind: 'date';
      value: number;
      numberFormat: 'date';
      notices: readonly string[];
      empty: boolean;
      padded: boolean;
    }>
  | Readonly<{ status: 'rejected'; reason: string; notices: readonly string[]; empty: boolean; padded: boolean }>;
export const TEXT_CONVERSION_POLICIES: readonly TextConversionPolicy[];
export const TEXT_CONVERSION_LIMITS: Readonly<{ textLength: 32767 }>;
/** Pure literal conversion. Empty input/padding stays empty text; ISO dates
 * require the explicit captured workbook calendar. No formula interpretation. */
export function convertTextField(
  raw: string,
  policy?: TextConversionPolicy,
  options?: TextConversionOptions
): TextConversionResult;
