import { FORMAT } from './format.mjs';
import { pageThemeSample } from './page-style-samples.mjs';
import { refusal } from './refusal-codes.mjs';
export const ownedPageThemePart = FORMAT.XL_THEME_PAGE_THEME_XML;
export function pageThemeCommand(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== 'expectedThemeRevision,sample,type' ||
    value.type !== 'applyPageTheme'
  )
    throw refusal('internal-invalid-field', 'Invalid page theme command.');
  if (
    value.expectedThemeRevision !== null &&
    (typeof value.expectedThemeRevision !== 'string' || !/^[a-f0-9]{64}$/.test(value.expectedThemeRevision))
  )
    throw refusal('internal-invalid-field', 'Invalid captured page theme revision.');
  pageThemeSample(value.sample);
  return { type: 'applyPageTheme', sample: value.sample, expectedThemeRevision: value.expectedThemeRevision };
}
export function normalizePageThemeCommands(commands) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-missing-field', 'Missing page theme command batch.');
  const result = commands.map((command) =>
    command?.type === 'applyPageTheme' ? pageThemeCommand(command) : structuredClone(command)
  );
  if (new TextEncoder().encode(JSON.stringify(result)).byteLength > 2 * 1024 * 1024)
    throw refusal('print-limit', 'Page theme batch exceeds 2 MiB.');
  return result;
}
