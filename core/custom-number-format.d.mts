/** Editor resource/runtime bounds, not Excel format limits. */
export const CUSTOM_NUMBER_FORMAT_LIMITS: Readonly<{ codeUnits: 4096; outputUnits: 65536; fractionPlaces: 100 }>;
export interface CustomNumberDisplay {
  text: string;
  color?: string;
  kind: 'number' | 'literal' | 'date' | 'text';
  supported: true;
}
/** null explicitly denotes an unsupported or malformed display profile. */
export function formatCustomNumber(
  value: number | string,
  code: string,
  date1904?: boolean
): CustomNumberDisplay | null;
