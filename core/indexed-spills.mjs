/** Metadata serialization cannot create ownership. Groups come only from the
 * native export plan for the accepted immutable workbook version. */
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { buildSpillMetadata } from './spill-metadata.mjs';

export function indexedSpillMetadata(bytes, calculation, groups) {
  const parts = unzipSync(bytes);
  for (const [part, xml] of Object.entries(calculation.replacements)) parts[part] = strToU8(xml);
  for (const part of calculation.removed) delete parts[part];
  return buildSpillMetadata(zipSync(parts), groups);
}

export function withSpillDimensions(view, result) {
  if (result.calculationPolicy) view = { ...view, calculationPolicy: result.calculationPolicy };
  if (!result.spillDimensions?.length) return view;
  return {
    ...view,
    sheets: view.sheets.map((sheet) => {
      const dimensions = result.spillDimensions.find((item) => item.path === sheet.path);
      return dimensions
        ? {
            ...sheet,
            rows: Math.max(sheet.rows, dimensions.rows),
            columns: Math.max(sheet.columns, dimensions.columns),
          }
        : sheet;
    }),
  };
}
