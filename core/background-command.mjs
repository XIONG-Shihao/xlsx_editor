import { worksheetBackgroundSample } from './page-style-samples.mjs';
import { refusal } from './refusal-codes.mjs';
export function backgroundCommand(value) {
  const remove = value?.type === 'removeSheetBackground';
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !==
      (remove ? 'expectedBackgroundRevision,sheet,type' : 'expectedBackgroundRevision,sample,sheet,type') ||
    !['setSheetBackground', 'removeSheetBackground'].includes(value.type)
  )
    throw refusal('internal-invalid-field', 'Invalid worksheet background command.');
  if (
    typeof value.sheet !== 'string' ||
    !value.sheet.trim() ||
    value.sheet.length > 31 ||
    /[\x00-\x1f\[\]:*?\\/]/u.test(value.sheet)
  )
    throw refusal('internal-invalid-field', 'Invalid background worksheet.');
  if (
    value.expectedBackgroundRevision !== null &&
    (typeof value.expectedBackgroundRevision !== 'string' || !/^[a-f0-9]{64}$/.test(value.expectedBackgroundRevision))
  )
    throw refusal('internal-invalid-field', 'Invalid captured background revision.');
  if (!remove) worksheetBackgroundSample(value.sample);
  return {
    type: value.type,
    sheet: value.sheet,
    expectedBackgroundRevision: value.expectedBackgroundRevision,
    ...(!remove ? { sample: value.sample } : {}),
  };
}
export function normalizeBackgroundCommands(commands) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-missing-field', 'Missing background command batch.');
  const result = commands.map((command) =>
    ['setSheetBackground', 'removeSheetBackground'].includes(command?.type)
      ? backgroundCommand(command)
      : structuredClone(command)
  );
  if (new TextEncoder().encode(JSON.stringify(result)).byteLength > 2 * 1024 * 1024)
    throw refusal('object-limit', 'Background batch exceeds 2 MiB.');
  return result;
}
