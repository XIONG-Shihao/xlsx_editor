/** Read-only Latin-font / literal-color theme projection. No authoring grant. */
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8 } from 'fflate';
import { zipPartSize } from './zip-parts.mjs';
import { refusal } from './refusal-codes.mjs';

const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';
const MAX_THEME_BYTES = 64 * 1024;
// Spreadsheet indices differ from the XML sequence (dk1, lt1, dk2, lt2).
// Independently checked against the installed SheetJS producer/parser.
export const spreadsheetThemeColors = Object.freeze([
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]);
const children = (node) => Array.from(node?.childNodes ?? []).filter((n) => n.nodeType === 1);
const direct = (node, name) => children(node).filter((n) => n.namespaceURI === DRAWING && n.localName === name);
function single(node, name) {
  const nodes = direct(node, name);
  if (nodes.length !== 1) throw refusal('style-refused', 'Theme has a missing or ambiguous ' + name);
  return nodes[0];
}
const rgb = (value) => (typeof value === 'string' && /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : undefined);
const family = (value) =>
  typeof value === 'string' && value.trim() && value.length <= 128 && !/[\x00-\x1f\x7f]/u.test(value)
    ? value
    : undefined;

/** The relationship inventory comes from the namespace-aware workbook reader.
 * Bound the ZIP member before inflation; never assume theme1.xml is the theme. */
export function readWorkbookTheme(parts, relationships) {
  const issues = new Set(),
    edges = Array.from(relationships).filter((edge) => edge.type === THEME_REL);
  if (!edges.length) return { theme: null, issues: [] };
  try {
    if (edges.length !== 1 || !edges[0].path)
      throw refusal('style-refused', 'Theme relationship is external or ambiguous');
    const part = edges[0].path;
    if (!/^[A-Za-z0-9_\-/]+\.xml$/.test(part) || part.includes('..') || part.includes('//'))
      throw refusal('style-unsupported', 'Theme part path is unsupported');
    const size = zipPartSize(parts, part);
    if (size === undefined) throw refusal('internal-missing-field', 'Theme part is missing');
    if (size > MAX_THEME_BYTES) throw refusal('style-limit', 'Theme XML exceeds 64 KiB');
    const source = strFromU8(parts[part]);
    if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw refusal('style-malformed', 'Theme XML entities are unsupported');
    const document = new DOMParser({
      errorHandler: {
        warning() {},
        error(message) {
          throw refusal('package-malformed', message);
        },
        fatalError(message) {
          throw refusal('package-malformed', message);
        },
      },
    }).parseFromString(source, 'application/xml');
    const root = document.documentElement;
    if (root.namespaceURI !== DRAWING || root.localName !== 'theme')
      throw refusal('style-unsupported', 'Theme namespace/root is unsupported');
    const elements = single(root, 'themeElements'),
      scheme = single(elements, 'clrScheme'),
      fonts = single(elements, 'fontScheme');
    const colors = {};
    for (const key of spreadsheetThemeColors) {
      const slot = single(scheme, key),
        values = children(slot),
        value = values[0];
      const color =
        values.length === 1 && value.namespaceURI === DRAWING && children(value).length === 0
          ? value.localName === 'srgbClr'
            ? rgb(value.getAttribute('val'))
            : value.localName === 'sysClr'
              ? rgb(value.getAttribute('lastClr'))
              : undefined
          : undefined;
      if (color) colors[key] = color;
      else issues.add('A theme color uses an unsupported value or transform; its cell color is unresolved');
    }
    const fontSchemes = {};
    for (const [key, tag] of [
      ['major', 'majorFont'],
      ['minor', 'minorFont'],
    ]) {
      const font = family(single(single(fonts, tag), 'latin').getAttribute('typeface'));
      if (font) fontSchemes[key] = font;
      else issues.add('A theme Latin font is unresolved; the explicit cell font is used');
    }
    return {
      theme: { part, name: root.getAttribute('name') || 'Imported theme', colors, fonts: fontSchemes },
      issues: [...issues],
    };
  } catch (error) {
    return { theme: null, issues: ['Workbook theme is unresolved: ' + String(error.message).slice(0, 200)] };
  }
}

/** Unsupported tints remain explicit until qualified; do not display the base
 * color as if it were the requested transformed color. Source XML is untouched. */
export function resolveThemeColor(node, theme, issues) {
  if (!node) return undefined;
  if (node.hasAttribute('tint') && node.getAttribute('tint') !== '' && Number(node.getAttribute('tint')) !== 0) {
    issues.add('Cell color tint is not qualified; its display color is unresolved');
    return undefined;
  }
  if (node.hasAttribute('theme')) {
    const raw = node.getAttribute('theme'),
      index = /^(?:0|[1-9]\d*)$/.test(raw) ? Number(raw) : -1;
    const result =
      index >= 0 && index < spreadsheetThemeColors.length ? theme?.colors[spreadsheetThemeColors[index]] : undefined;
    if (!result) issues.add('Cell theme color is unresolved');
    return result;
  }
  const value = node.getAttribute('rgb');
  return /^[0-9a-f]{8}$/i.test(value) ? value.slice(-6).toUpperCase() : rgb(value);
}

export function resolveThemeFont(explicit, scheme, theme, issues) {
  if (!scheme || scheme === 'none') return explicit ?? undefined;
  if (['major', 'minor'].includes(scheme) && theme?.fonts[scheme]) return theme.fonts[scheme];
  issues.add('Cell theme font is unresolved; its explicit font is used');
  return explicit ?? undefined;
}
