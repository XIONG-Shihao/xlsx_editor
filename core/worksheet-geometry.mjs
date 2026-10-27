/** Physical worksheet dimensions at 96 dpi. No cells, history or owner state. */
import { refusal } from './refusal-codes.mjs';

export const WORKSHEET_MAX_ROWS = 1_048_576,
  WORKSHEET_MAX_COLUMNS = 16_384;
export const PIXELS_PER_POINT = 96 / 72;
export const DEFAULT_NORMAL_FONT = Object.freeze({ fontFamily: 'Calibri', fontSize: 11, bold: false, italic: false });
const object = (value, label) => {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
};
const exact = (value, required, optional = []) => {
  for (const key of required)
    if (!Object.hasOwn(value, key)) throw refusal('internal-missing-field', 'Missing geometry ' + key);
  for (const key of Object.keys(value))
    if (!required.includes(key) && !optional.includes(key))
      throw refusal('internal-unknown-reference', 'Unknown geometry ' + key);
};
const integer = (value, min, max, label) => {
  if (!Number.isInteger(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
};
const finite = (value, min, max, label) => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
};
const family = (value) => {
  if (typeof value !== 'string' || !value.trim() || value.length > 256 || /[\x00-\x1f\x7f]/u.test(value))
    throw refusal('internal-invalid-field', 'Invalid Normal font family');
  return value;
};
export function normalGeometryFont(sheet) {
  const font = sheet?.normalFont ?? DEFAULT_NORMAL_FONT;
  object(font, 'Normal font');
  exact(font, ['fontFamily', 'fontSize', 'bold', 'italic']);
  if (typeof font.bold !== 'boolean' || typeof font.italic !== 'boolean')
    throw refusal('internal-invalid-field', 'Invalid Normal font style');
  return Object.freeze({
    fontFamily: family(font.fontFamily),
    fontSize: finite(font.fontSize, Number.MIN_VALUE, 409.5, 'Normal font size'),
    bold: font.bold,
    italic: font.italic,
  });
}
/** The caller binds this read-only rendering profile to its accepted owner/head.
 * Font measurements do not certify the availability of the requested font. */
export function validateGeometryProfile(value, expectedFont) {
  object(value, 'geometry profile');
  exact(value, ['schema', 'font', 'maximumDigitWidth', 'source', 'availability'], ['renderedFontFamily']);
  if (
    value.schema !== 1 ||
    !['local-font', 'substituted-font', 'reference-fallback'].includes(value.source) ||
    !['confirmed', 'unverified', 'substituted'].includes(value.availability)
  )
    throw refusal('internal-invalid-field', 'Invalid geometry profile provenance');
  const font = normalGeometryFont({ normalFont: value.font });
  if (expectedFont) {
    const expected = normalGeometryFont({ normalFont: expectedFont });
    if (Object.keys(font).some((key) => font[key] !== expected[key]))
      throw refusal('layout-bad-input', 'Geometry profile does not match the accepted Normal font');
  }
  if (
    (value.source === 'substituted-font' && value.availability !== 'substituted') ||
    (value.source === 'reference-fallback' && value.availability === 'confirmed')
  )
    throw refusal('internal-invalid-field', 'Invalid geometry profile availability');
  return Object.freeze({
    schema: 1,
    font,
    maximumDigitWidth: integer(value.maximumDigitWidth, 1, 4096, 'maximum digit width'),
    source: value.source,
    availability: value.availability,
    ...(value.renderedFontFamily === undefined ? {} : { renderedFontFamily: family(value.renderedFontFamily) }),
  });
}
/** Named reference fallback; it is never a universal font-parity claim. */
export function referenceGeometryProfile(font = DEFAULT_NORMAL_FONT) {
  return validateGeometryProfile({
    schema: 1,
    font,
    maximumDigitWidth: 7,
    source: 'reference-fallback',
    availability: 'unverified',
  });
}
export function xmlColumnWidthPixels(width, maximumDigitWidth) {
  // A baseColWidth of255 excludes up to5px padding; its derived encoded
  // width can reach260 at a one-pixel digit. Source XML attributes are
  // separately validated at255 by createWorksheetDimensions.
  finite(width, 0, 260, 'encoded column width');
  integer(maximumDigitWidth, 1, 4096, 'maximum digit width');
  return Math.floor(((256 * width + Math.floor(128 / maximumDigitWidth)) / 256) * maximumDigitWidth);
}
/** baseColWidth excludes the five pixels of margin/gridline padding; unlike an
 * explicit/defaultColWidth XML value it must first be encoded. */
export function encodedColumnWidth(characters, maximumDigitWidth) {
  finite(characters, 0, 255, 'base column width');
  integer(maximumDigitWidth, 1, 4096, 'maximum digit width');
  return Math.floor((characters + 5 / maximumDigitWidth) * 256) / 256;
}
export function worksheetDefaultRowHeight(sheet) {
  return finite(sheet?.sheetFormat?.defaultRowHeight ?? 15, 0, 409.5, 'default row height');
}
export function worksheetDefaultColumnWidth(sheet, profile) {
  const p = validateGeometryProfile(
    profile ?? referenceGeometryProfile(normalGeometryFont(sheet)),
    normalGeometryFont(sheet)
  );
  const format = sheet?.sheetFormat;
  return format?.defaultColumnWidth === undefined
    ? encodedColumnWidth(format?.baseColumnWidth ?? 8, p.maximumDigitWidth)
    : finite(format.defaultColumnWidth, 0, 255, 'default column width');
}
export function createWorksheetDimensions(sheet, profile) {
  const font = normalGeometryFont(sheet),
    p = validateGeometryProfile(profile ?? referenceGeometryProfile(font), font),
    defaultRowHeight = worksheetDefaultRowHeight(sheet),
    defaultColumnWidth = worksheetDefaultColumnWidth(sheet, p);
  const heights = new Map(),
    hidden = new Set(),
    visible = new Set();
  for (const item of sheet?.rowHeights ?? []) {
    const row = integer(item.row, 0, WORKSHEET_MAX_ROWS - 1, 'row height index'),
      height = finite(item.height, 0, 409.5, 'row height');
    if (!heights.has(row)) heights.set(row, height);
  }
  for (const row of sheet?.hiddenRows ?? []) hidden.add(integer(row, 0, WORKSHEET_MAX_ROWS - 1, 'hidden row'));
  for (const row of sheet?.visibleRows ?? []) visible.add(integer(row, 0, WORKSHEET_MAX_ROWS - 1, 'visible row'));
  const hiddenDefault = sheet?.sheetFormat?.rowsHiddenByDefault ?? false;
  if (typeof hiddenDefault !== 'boolean') throw refusal('internal-invalid-field', 'Invalid default row visibility');
  const rowHidden = (row) => {
    integer(row, 0, WORKSHEET_MAX_ROWS - 1, 'row');
    return hidden.has(row) || (hiddenDefault && !visible.has(row));
  };
  const rowHeightPt = (row) => (rowHidden(row) ? 0 : (heights.get(row) ?? defaultRowHeight));
  const widths = sheet?.widths ?? [];
  for (const item of widths) {
    integer(item.min, 1, WORKSHEET_MAX_COLUMNS, 'column interval');
    integer(item.max, item.min, WORKSHEET_MAX_COLUMNS, 'column interval');
    finite(item.width, 0, 255, 'column width');
    if (typeof item.hidden !== 'boolean') throw refusal('internal-invalid-field', 'Invalid column visibility');
  }
  const columnPixels = (col) => {
    integer(col, 0, WORKSHEET_MAX_COLUMNS - 1, 'column');
    const item = widths.find((item) => col + 1 >= item.min && col + 1 <= item.max);
    return item?.hidden ? 0 : xmlColumnWidthPixels(item?.width ?? defaultColumnWidth, p.maximumDigitWidth);
  };
  const rowDistancePt = (first, last) => {
    integer(first, 0, WORKSHEET_MAX_ROWS, 'row offset');
    integer(last, 0, WORKSHEET_MAX_ROWS, 'row offset');
    const a = Math.min(first, last),
      b = Math.max(first, last);
    let result = hiddenDefault ? 0 : (b - a) * defaultRowHeight;
    if (hiddenDefault) {
      for (const row of visible)
        if (row >= a && row < b && !hidden.has(row)) result += heights.get(row) ?? defaultRowHeight;
    } else {
      for (const [row, height] of heights) if (row >= a && row < b) result += height - defaultRowHeight;
      for (const row of hidden) if (row >= a && row < b) result -= heights.get(row) ?? defaultRowHeight;
    }
    return (first <= last ? 1 : -1) * result;
  };
  return Object.freeze({
    profile: p,
    font,
    defaultRowHeight,
    defaultColumnWidth,
    defaultColumnPixels: xmlColumnWidthPixels(defaultColumnWidth, p.maximumDigitWidth),
    rowHidden,
    rowHeightPt,
    columnPixels,
    rowDistancePt,
  });
}
export function worksheetRowOffsets(sheet, count, profile) {
  integer(count, 0, WORKSHEET_MAX_ROWS, 'row count');
  const model = createWorksheetDimensions(sheet, profile),
    offsets = new Float64Array(count + 1);
  let points = 0;
  for (let row = 0; row < count; row++) {
    points += model.rowHeightPt(row);
    offsets[row + 1] = points * PIXELS_PER_POINT;
  }
  return offsets;
}
export function worksheetColumnOffsets(sheet, count, profile) {
  integer(count, 0, WORKSHEET_MAX_COLUMNS, 'column count');
  const model = createWorksheetDimensions(sheet, profile),
    sizes = new Float64Array(count).fill(model.defaultColumnPixels),
    next = new Uint32Array(count + 1);
  for (let i = 0; i <= count; i++) next[i] = i;
  const free = (start) => {
    let index = start;
    while (next[index] !== index) index = next[index];
    while (next[start] !== start) {
      const old = next[start];
      next[start] = index;
      start = old;
    }
    return index;
  };
  for (const item of sheet?.widths ?? []) {
    const end = Math.min(count, item.max),
      pixels = item.hidden ? 0 : xmlColumnWidthPixels(item.width, model.profile.maximumDigitWidth);
    for (let index = free(Math.min(count, item.min - 1)); index < end; index = free(index)) {
      sizes[index] = pixels;
      next[index] = free(index + 1);
    }
  }
  const offsets = new Float64Array(count + 1);
  for (let i = 0; i < count; i++) offsets[i + 1] = offsets[i] + sizes[i];
  return offsets;
}

export function rowHeightPoints(sheet, row) {
  return createWorksheetDimensions(sheet).rowHeightPt(row);
}
export function columnWidthPixels(sheet, col, profile) {
  return createWorksheetDimensions(sheet, profile).columnPixels(col);
}
export function rowOffsetPixels(sheet, row) {
  return createWorksheetDimensions(sheet).rowDistancePt(0, row) * PIXELS_PER_POINT;
}
export function columnOffsetPixels(sheet, col, profile) {
  return worksheetColumnOffsets(sheet, col, profile)[col];
}
