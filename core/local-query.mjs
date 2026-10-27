import { FORMAT, matchesFormat } from './format.mjs';
/** Closed literal select/filter pipeline over immutable bounded delimited data. */
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { delimitedFields, inspectDelimitedSource } from './delimited-import.mjs';
import {
  PASTE_BYTE_LIMIT,
  PASTE_POSITION_LIMIT,
  PASTE_PAGE_BYTES,
  PASTE_PAGE_CELLS,
  pasteBytes,
  pasteId,
  validatePasteShape,
  validatePasteCells,
  validatePasteDescriptor,
  validatePasteState,
} from './paste-protocol.mjs';
import { assertXmlText, cellPosition, columnName } from './workbook-commands.mjs';
import { convertTextField, TEXT_CONVERSION_POLICIES } from './text-conversion.mjs';
import { hasExactNumericValue } from './calculation.mjs';
import { validateTypedPasteConversion, TYPED_PASTE_POLICY } from './paste-protocol.mjs';
import { refusal } from './refusal-codes.mjs';
export const QUERY_COLUMNS = 256,
  QUERY_ROW_BYTES = 1024 * 1024,
  QUERY_PREVIEW_BYTES = 128 * 1024;
const exact = (value, keys) =>
  value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).every((key) => keys.includes(key));
export function validateLocalQuery(value) {
  if (
    !exact(value, ['columns', 'filters', 'keepHeader', 'policies']) ||
    !Array.isArray(value.columns) ||
    value.columns.length > QUERY_COLUMNS ||
    new Set(value.columns).size !== value.columns.length ||
    value.columns.some((column) => !Number.isInteger(column) || column < 1 || column > 16384) ||
    !Array.isArray(value.filters) ||
    value.filters.length > 8 ||
    typeof value.keepHeader !== 'boolean'
  )
    throw refusal('internal-invalid-field', 'Invalid bounded local query');
  if (
    value.policies !== undefined &&
    (!Array.isArray(value.policies) ||
      value.policies.length > QUERY_COLUMNS ||
      Array.from(value.policies).some((policy) => !TEXT_CONVERSION_POLICIES.includes(policy)))
  )
    throw refusal('internal-invalid-field', 'Invalid query column conversion policies');
  const filters = value.filters.map((filter) => {
    if (
      !exact(filter, ['column', 'op', 'value', 'caseSensitive']) ||
      !Number.isInteger(filter.column) ||
      filter.column < 1 ||
      filter.column > 16384 ||
      !['equals', 'notEquals', 'contains', 'isEmpty', 'notEmpty'].includes(filter.op) ||
      typeof filter.value !== 'string' ||
      filter.value.length > 4096 ||
      typeof filter.caseSensitive !== 'boolean'
    )
      throw refusal('internal-invalid-field', 'Invalid literal query filter');
    assertXmlText(filter.value);
    return { ...filter };
  });
  return {
    columns: [...value.columns],
    filters,
    keepHeader: value.keepHeader,
    ...(value.policies !== undefined ? { policies: [...value.policies] } : {}),
  };
}
const fingerprint = () => {
  const hash = sha256.create();
  return {
    add(index, kind, value) {
      hash.update(new TextEncoder().encode(JSON.stringify([index, kind, value]) + '\n'));
    },
    finish() {
      return bytesToHex(hash.digest());
    },
    close() {
      hash.destroy();
    },
  };
};
export function queryCellFingerprint() {
  return fingerprint();
}
export function queryCellIdentity(cell) {
  if (cell?.formula !== undefined || cell?.formulaType !== undefined || cell?.spill)
    return ['protected', JSON.stringify([cell.formula ?? null, cell.formulaType ?? null, cell.spill ?? null])];
  if (cell?.calculation?.status === 'error' || cell?.error !== undefined)
    return ['error', JSON.stringify(cell.calculation?.error ?? cell.error)];
  const value = cell?.value;
  if (value === undefined || value === null || value === '') return ['text', ''];
  if (typeof value === 'string') return ['text', value];
  if (typeof value === 'boolean') return ['boolean', value];
  if (typeof value === 'number' && Number.isFinite(value))
    return ['number', Object.is(value, -0) ? '-0' : String(value)];
  return ['protected', JSON.stringify(value) ?? 'unknown'];
}
/** V1 intentionally folds blank/empty text. Typed managed output uses an
 * explicit profile, validating numeric meaning rather than raw XML spelling. */
export const QUERY_TYPED_FINGERPRINT = 'typed-semantic-v1';
export const queryHasConversions = (query) => (query.policies ?? []).some((policy) => policy !== 'text');
export function typedQueryCellIdentity(cell, { conversion, shape, index }) {
  const row = Math.floor(index / shape.columns),
    col = index % shape.columns,
    policy =
      row < conversion.outputRows && col < conversion.outputColumns && !(conversion.keepHeader && row === 0)
        ? conversion.policies[col]
        : 'text';
  if (
    cell?.formula !== undefined ||
    cell?.formulaType !== undefined ||
    cell?.spill ||
    cell?.calculation?.pending ||
    ['stored', 'unresolved', 'circular'].includes(cell?.calculation?.status)
  )
    return ['protected', JSON.stringify(cell)];
  if (cell?.dataType === 'error' || cell?.error !== undefined || cell?.calculation?.status === 'error')
    return ['error', cell.error ?? cell.calculation?.error ?? cell.rawValue];
  const value = cell?.value;
  if (value === undefined || value === null) return ['blank', ''];
  if (typeof value === 'string') return ['text', value];
  if (typeof value === 'boolean') return ['boolean', value];
  if (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    (!Number.isInteger(value) || Number.isSafeInteger(value)) &&
    hasExactNumericValue(value, cell?.rawValue)
  ) {
    if (policy === 'iso-date' && cell?.format !== 14)
      return ['protected', JSON.stringify(['date-format', cell?.format ?? null])];
    return [policy === 'iso-date' ? 'date' : 'number', Object.is(value, -0) ? '-0' : String(value)];
  }
  return ['protected', JSON.stringify([value, cell?.rawValue ?? null])];
}
export const queryConversionDisplay = (result) =>
  result.status === 'converted'
    ? typeof result.value === 'number'
      ? Object.is(result.value, -0)
        ? '-0'
        : String(result.value)
      : typeof result.value === 'string'
        ? JSON.stringify(result.value)
        : String(result.value)
    : undefined;
export function localQueryConversionIssue({
  index,
  raw,
  sourceRow,
  sourceColumn,
  outputRow,
  outputColumn,
  target,
  policy,
  result,
}) {
  const start = cellPosition(target.start);
  return {
    index,
    sourceRow,
    sourceColumn,
    outputRow,
    outputColumn,
    sourceRef: columnName(sourceColumn - 1) + sourceRow,
    destinationRef: columnName(start.col + outputColumn - 1) + (start.row + outputRow),
    raw,
    policy,
    status: result.status,
    ...(result.status === 'converted'
      ? { kind: result.kind, display: queryConversionDisplay(result) }
      : { reason: result.reason }),
    notices: [...result.notices],
  };
}
/** Replacement review binds exact authored metadata. It is deliberately
 * distinct from the cross-Save semantic identity of prior managed output. */
export function queryReplacementIdentity(cell) {
  const encode = (value) => {
    if (value === undefined) return ['undefined'];
    if (value === null) return ['null'];
    if (typeof value === 'number')
      return [
        'number',
        Object.is(value, -0)
          ? '-0'
          : Number.isNaN(value)
            ? 'NaN'
            : value === Infinity
              ? 'Infinity'
              : value === -Infinity
                ? '-Infinity'
                : String(value),
      ];
    if (typeof value === 'string' || typeof value === 'boolean') return [typeof value, value];
    if (Array.isArray(value)) return ['array', Array.from(value, encode)];
    if (typeof value === 'object')
      return [
        'object',
        Object.keys(value)
          .sort()
          .map((key) => [key, encode(value[key])]),
      ];
    throw refusal('query-unsupported', 'Unsupported destination cell metadata');
  };
  return ['replacement', JSON.stringify(encode(cell))];
}
export function queryReplacementCurrent(cell) {
  const scalar = (value) =>
    typeof value === 'number'
      ? Object.is(value, -0)
        ? '-0'
        : Number.isFinite(value)
          ? String(value)
          : String(value)
      : typeof value === 'string'
        ? JSON.stringify(value)
        : typeof value === 'boolean'
          ? String(value)
          : 'Blank';
  if (cell !== undefined && (!cell || typeof cell !== 'object' || Array.isArray(cell)))
    throw refusal('internal-invalid-field', 'Invalid destination cell record');
  if (
    (cell?.value !== undefined &&
      cell.value !== null &&
      !['string', 'number', 'boolean'].includes(typeof cell.value)) ||
    (cell?.dataType !== undefined && !['blank', 'text', 'number', 'boolean', 'error'].includes(cell.dataType)) ||
    (cell?.rawValue !== undefined && cell.rawValue !== null && typeof cell.rawValue !== 'string') ||
    (cell?.format !== undefined && !['string', 'number'].includes(typeof cell.format))
  )
    throw refusal('internal-invalid-field', 'Invalid destination scalar or metadata');
  const current = {
    kind: 'blank',
    display: 'Blank',
    rawValue: typeof cell?.rawValue === 'string' ? cell.rawValue : null,
    formula: typeof cell?.formula === 'string' ? cell.formula : null,
    format: cell?.format ?? null,
  };
  const special =
    cell?.spill ||
    cell?.scenarioResult ||
    cell?.sharedFormulaId !== undefined ||
    (cell?.formulaType !== undefined && cell.formulaType !== 'normal') ||
    (!cell?.formula &&
      (cell?.calculation?.pending || ['stored', 'unresolved', 'circular'].includes(cell?.calculation?.status)));
  if (current.formula !== null) {
    current.kind = 'formula';
    current.display = '=' + current.formula;
    if (cell?.calculation?.pending || cell?.calculation?.status === 'stored') {
      current.pending = true;
      const stored = cell?.cachedResult;
      current.storedDisplay =
        stored?.authority === 'stored'
          ? stored.status === 'error'
            ? 'Stored error: ' + String(stored.error)
            : 'Stored last result: ' + scalar(stored.value)
          : 'Stored result is pending calculation.';
    }
  } else if (cell?.dataType === 'error' || cell?.error !== undefined) {
    current.kind = 'error';
    current.display = String(cell.error ?? cell.rawValue ?? cell.value);
  } else if (typeof cell?.value === 'number') {
    current.kind = 'number';
    const unsafe =
      !Number.isFinite(cell.value) ||
      (Number.isInteger(cell.value) && !Number.isSafeInteger(cell.value)) ||
      !hasExactNumericValue(cell.value, cell.rawValue);
    current.display = unsafe && current.rawValue !== null ? current.rawValue : scalar(cell.value);
    if (unsafe) current.reason = 'Stored numeric precision; the original raw value is shown.';
  } else if (typeof cell?.value === 'string') {
    current.kind = 'text';
    current.display = JSON.stringify(cell.value);
  } else if (typeof cell?.value === 'boolean') {
    current.kind = 'boolean';
    current.display = String(cell.value);
  }
  if (special) {
    current.kind = 'protected';
    current.reason = 'This destination contains a protected array, shared formula, spill or stored object result.';
    if (!current.formula) current.display = current.rawValue !== null ? current.rawValue : 'Protected stored result';
  }
  const occupied = Boolean(
    special ||
      current.formula !== null ||
      (cell?.value !== undefined && cell.value !== null) ||
      cell?.error !== undefined ||
      cell?.dataType === 'error'
  );
  return { current, occupied, blocked: Boolean(special) };
}
const predicate = (value, filter) => {
  let wanted = filter.value;
  if (!filter.caseSensitive) {
    value = value.toLowerCase();
    wanted = wanted.toLowerCase();
  }
  switch (filter.op) {
    case 'equals':
      return value === wanted;
    case 'notEquals':
      return value !== wanted;
    case 'contains':
      return value.includes(wanted);
    case 'isEmpty':
      return value === '';
    case 'notEmpty':
      return value !== '';
  }
};
async function* rows(source, inspection, query, { signal, onProgress = () => {}, proof } = {}) {
  const columns = query.columns.length
    ? query.columns
    : Array.from({ length: inspection.shape.columns }, (_, index) => index + 1);
  if (
    !columns.length ||
    columns.length > QUERY_COLUMNS ||
    [...columns, ...query.filters.map((filter) => filter.column)].some((column) => column > inspection.shape.columns)
  )
    throw refusal('query-limit', 'Query columns are outside the source or exceed 256 selected columns');
  const needed = new Set([...columns, ...query.filters.map((filter) => filter.column)]);
  let index = -1,
    values = new Map(),
    bytes = 0,
    count = 0;
  const output = () =>
    (query.keepHeader && index === 0) ||
    query.filters.every((filter) => predicate(values.get(filter.column) ?? '', filter))
      ? {
          values: columns.map((column) => values.get(column) ?? ''),
          sourceRow: index + 1,
          padded: columns.map((column) => !values.has(column)),
        }
      : null;
  for await (const field of delimitedFields(source, inspection.delimiter, { signal, proof })) {
    if (field.row >= inspection.shape.rows || field.col >= inspection.shape.columns || ++count > inspection.count)
      throw refusal('query-stale', 'Query source shape changed');
    if (index !== field.row) {
      if (index >= 0) {
        const result = output();
        if (result) yield result;
        onProgress({ phase: 'querying-file', completed: index + 1, total: inspection.shape.rows });
      }
      index = field.row;
      values = new Map();
      bytes = 0;
    }
    if (needed.has(field.col + 1)) {
      bytes += pasteBytes(JSON.stringify(field.text));
      if (bytes > QUERY_ROW_BYTES) throw refusal('query-limit', 'Selected query row exceeds the 1 MiB buffer');
      values.set(field.col + 1, field.text);
    }
  }
  if (index >= 0) {
    const result = output();
    if (result) yield result;
    onProgress({ phase: 'querying-file', completed: index + 1, total: inspection.shape.rows });
  }
  if (
    count !== inspection.count ||
    index + 1 !== inspection.shape.rows ||
    proof.sha256 !== inspection.source.sha256 ||
    proof.bytes !== inspection.source.bytes ||
    proof.encoding !== inspection.encoding ||
    proof.bom !== inspection.bom
  )
    throw refusal('query-stale', 'Query immutable source changed between passes');
}
export async function inspectLocalQuery(
  source,
  {
    sheet,
    start,
    query,
    previousShape,
    signal,
    onProgress = () => {},
    date1904,
    onCell = async () => {},
    onLayout = async () => {},
  }
) {
  query = validateLocalQuery(query);
  const raw = await inspectDelimitedSource(source, {
      sheet,
      start: 'A1',
      delimiter: source.name.toLowerCase().endsWith('.tsv') ? '\t' : 'auto',
      signal,
      onProgress,
    }),
    columns = query.columns.length || raw.shape.columns;
  if (columns > QUERY_COLUMNS) throw refusal('query-limit', 'Select at most 256 query columns');
  if ((query.policies?.length ?? 0) > columns)
    throw refusal('query-limit', 'Conversion policies exceed output columns');
  const typed = queryHasConversions(query),
    policies = Array.from({ length: columns }, (_, index) => query.policies?.[index] ?? 'text');
  if (typed && typeof date1904 !== 'boolean')
    throw refusal('query-bad-input', 'Typed query requires the accepted workbook date calendar');
  let outputRows = 0,
    preview = [],
    previewBytes = 0,
    truncated = false,
    rejected = 0,
    changed = 0,
    issues = 0;
  const proof = {},
    reviewHash = fingerprint(),
    sourceColumns = query.columns.length ? query.columns : Array.from({ length: columns }, (_, index) => index + 1);
  try {
    await onLayout({ outputColumns: columns, sourceRows: raw.shape.rows, columns: sourceColumns });
    for await (const entry of rows(source, raw, query, { signal, onProgress, proof })) {
      const row = entry.values;
      outputRows++;
      if (outputRows * columns > PASTE_POSITION_LIMIT)
        throw refusal('query-limit', 'Query result exceeds 1,000,000 positions');
      for (let col = 0; col < columns; col++) {
        const policy = query.keepHeader && outputRows === 1 ? 'text' : policies[col],
          result = convertTextField(row[col], policy, { date1904, padded: entry.padded[col] });
        if (result.status === 'rejected') rejected++;
        else if (result.notices.length) changed++;
        if (result.status === 'rejected' || result.notices.length) issues++;
        reviewHash.add(
          (outputRows - 1) * columns + col,
          'raw',
          JSON.stringify([
            entry.sourceRow,
            row[col],
            policy,
            result.status === 'converted'
              ? [result.kind, queryConversionDisplay(result), result.numberFormat ?? null]
              : result.reason,
            result.notices,
          ])
        );
        await onCell({
          index: (outputRows - 1) * columns + col,
          raw: row[col],
          sourceRow: entry.sourceRow,
          sourceColumn: sourceColumns[col],
          outputRow: outputRows,
          outputColumn: col + 1,
          target: { sheet, start },
          policy,
          result,
        });
      }
      if (preview.length < 20 && columns <= 50) {
        const values = row.map((value) => {
            if (value.length > 256) truncated = true;
            return value.length > 256 ? value.slice(0, 255) + '…' : value;
          }),
          size = pasteBytes(JSON.stringify(values));
        if (previewBytes + size <= QUERY_PREVIEW_BYTES) {
          preview.push(values);
          previewBytes += size;
        }
      }
    }
    if (previousShape) {
      previousShape = validatePasteShape(previousShape);
      if (previousShape.sheet !== sheet || previousShape.start !== start)
        throw refusal('query-refused', 'Query target moved; prior output is kept in its original location');
    }
    const shape = validatePasteShape({
      sheet,
      start,
      rows: Math.max(1, outputRows, previousShape?.rows ?? 0),
      columns: Math.max(columns, previousShape?.columns ?? 0),
    });
    return {
      schema: FORMAT.SCHEMA_LOCAL_QUERY_V1,
      query,
      raw,
      source: raw.source,
      outputRows,
      outputColumns: columns,
      dataRows: Math.max(0, outputRows - (query.keepHeader ? 1 : 0)),
      shape,
      count: shape.rows * shape.columns,
      policy: typed ? TYPED_PASTE_POLICY : 'text',
      ...(typed
        ? {
            conversion: validateTypedPasteConversion(
              { schema: 1, policies, date1904, keepHeader: query.keepHeader, outputRows, outputColumns: columns },
              shape
            ),
          }
        : {}),
      conversionSummary: { rejected, changed, issues, fingerprint: reviewHash.finish() },
      preview,
      previewTruncated: truncated || preview.length < outputRows || columns > 50,
    };
  } finally {
    reviewHash.close();
  }
}
/** Re-stream literal fields against one complete inspection. The final proof
 * is supplied only after the complete immutable source and row count agree. */
export async function* localQueryFields(
  source,
  inspection,
  { signal, onProgress = () => {}, onComplete = () => {} } = {}
) {
  if (
    !matchesFormat(inspection?.schema, 'SCHEMA_LOCAL_QUERY_V1') ||
    inspection.source.id !== source.id ||
    (source.sha256 && inspection.source.sha256 !== source.sha256)
  )
    throw refusal('query-stale', 'Query inspection belongs to another immutable source');
  const query = validateLocalQuery(inspection.query),
    columns = query.columns.length ? query.columns : Array.from({ length: inspection.outputColumns }, (_, i) => i + 1),
    proof = {},
    hash = fingerprint();
  let rowIndex = 0;
  try {
    for await (const entry of rows(source, inspection.raw, query, { signal, onProgress, proof })) {
      if (rowIndex >= inspection.outputRows || entry.values.length !== inspection.outputColumns)
        throw refusal('query-stale', 'Query result shape changed');
      for (let col = 0; col < inspection.outputColumns; col++) {
        signal?.throwIfAborted();
        const policy = query.keepHeader && rowIndex === 0 ? 'text' : (query.policies?.[col] ?? 'text'),
          raw = entry.values[col],
          result = convertTextField(raw, policy, {
            date1904: inspection.conversion?.date1904,
            padded: entry.padded[col],
          }),
          index = rowIndex * inspection.outputColumns + col;
        hash.add(
          index,
          'raw',
          JSON.stringify([
            entry.sourceRow,
            raw,
            policy,
            result.status === 'converted'
              ? [result.kind, queryConversionDisplay(result), result.numberFormat ?? null]
              : result.reason,
            result.notices,
          ])
        );
        yield {
          index,
          raw,
          sourceRow: entry.sourceRow,
          sourceColumn: columns[col],
          outputRow: rowIndex + 1,
          outputColumn: col + 1,
          target: { sheet: inspection.shape.sheet, start: inspection.shape.start },
          policy,
          result,
        };
      }
      rowIndex++;
    }
    if (rowIndex !== inspection.outputRows) throw refusal('query-stale', 'Query result row count changed');
    await onComplete({
      source: inspection.source,
      outputRows: rowIndex,
      outputColumns: inspection.outputColumns,
      shape: inspection.shape,
      fingerprint: hash.finish(),
    });
  } finally {
    hash.close();
  }
}
export async function* localQueryPages(source, inspection, { signal, onProgress = () => {} } = {}) {
  if (
    !matchesFormat(inspection?.schema, 'SCHEMA_LOCAL_QUERY_V1') ||
    inspection.source.id !== source.id ||
    (source.sha256 && inspection.source.sha256 !== source.sha256)
  )
    throw refusal('query-stale', 'Query inspection belongs to another immutable source');
  const shape = validatePasteShape(inspection.shape),
    query = validateLocalQuery(inspection.query),
    hash = fingerprint(),
    proof = {};
  let cells = [],
    bytes = 0,
    totalBytes = 0,
    count = 0,
    sequence = 0,
    rowIndex = 0,
    previous = -1;
  try {
    const entry = async function* (text, column) {
      const value = [rowIndex * shape.columns + column, text],
        size = pasteBytes(JSON.stringify(value) + '\n');
      if (cells.length && (cells.length >= PASTE_PAGE_CELLS || bytes + size > PASTE_PAGE_BYTES)) {
        validatePasteCells(cells, shape, previous);
        previous = cells.at(-1)[0];
        yield { sequence: sequence++, cells };
        cells = [];
        bytes = 0;
      }
      if (size > PASTE_PAGE_BYTES || (totalBytes += size) > PASTE_BYTE_LIMIT)
        throw refusal('query-limit', 'Query attachment exceeds its 64 MiB budget');
      cells.push(value);
      bytes += size;
      count++;
      const result = inspection.conversion
        ? convertTextField(
            text,
            rowIndex >= inspection.outputRows ||
              column >= inspection.outputColumns ||
              (inspection.conversion.keepHeader && rowIndex === 0)
              ? 'text'
              : inspection.conversion.policies[column],
            { date1904: inspection.conversion.date1904 }
          )
        : null;
      if (result?.status === 'rejected')
        throw refusal('query-stale', 'Query conversion changed before staging: ' + result.reason);
      hash.add(
        value[0],
        result?.kind ?? 'text',
        result && typeof result.value === 'number'
          ? Object.is(result.value, -0)
            ? '-0'
            : String(result.value)
          : (result?.value ?? text)
      );
    };
    for await (const rowEntry of rows(source, inspection.raw, query, { signal, onProgress, proof })) {
      const row = rowEntry.values;
      if (rowIndex >= inspection.outputRows || row.length !== inspection.outputColumns)
        throw refusal('query-stale', 'Query result shape changed');
      for (let col = 0; col < shape.columns; col++) yield* entry(row[col] ?? '', col);
      rowIndex++;
    }
    if (rowIndex !== inspection.outputRows) throw refusal('query-stale', 'Query result row count changed');
    for (; rowIndex < shape.rows; rowIndex++) for (let col = 0; col < shape.columns; col++) yield* entry('', col);
    if (count !== inspection.count) throw refusal('query-stale', 'Query dense replacement changed');
    inspection.outputFingerprint = hash.finish();
    // No final page is released until source and output shape/hash are complete.
    if (cells.length) {
      validatePasteCells(cells, shape, previous);
      yield { sequence, cells };
    }
  } finally {
    hash.close();
  }
}
export async function stageLocalQuery(
  source,
  inspection,
  { state, signal, checkCurrent = () => {}, sink, onProgress = () => {} }
) {
  validatePasteState(state);
  if (inspection.conversionSummary?.rejected)
    throw refusal('query-refused', 'Query has rejected conversions; choose Text or correct the source before Refresh');
  if (!sink || ['begin', 'upload', 'seal', 'cancel', 'prune'].some((key) => typeof sink[key] !== 'function'))
    throw refusal('query-bad-input', 'Query staging requires a trusted owner');
  let id,
    sealed = false;
  const guard = async () => {
    signal?.throwIfAborted();
    await source.check();
    await checkCurrent();
    signal?.throwIfAborted();
  };
  try {
    await guard();
    const begin = await sink.begin({
      state,
      shape: inspection.shape,
      count: inspection.count,
      policy: inspection.policy,
      ...(inspection.conversion ? { conversion: inspection.conversion } : {}),
    });
    id = begin.id ?? begin.uploadId;
    if (!pasteId(id) || (begin.sequence ?? begin.nextSequence ?? 0) !== 0)
      throw refusal('internal-invalid-field', 'Invalid query upload grant');
    let count = 0,
      sequence = 0;
    for await (const page of localQueryPages(source, inspection, { signal, onProgress })) {
      await guard();
      onProgress({ phase: 'uploading-query', completed: count, total: inspection.count });
      const receipt = await sink.upload({ id, state, ...page });
      count += page.cells.length;
      sequence++;
      if (
        (receipt.id ?? id) !== id ||
        (receipt.sequence ?? receipt.nextSequence) !== sequence ||
        (receipt.count ?? receipt.accepted) !== count
      )
        throw refusal('query-stale', 'Query upload acknowledgement changed');
      onProgress({ phase: 'uploading-query', completed: count, total: inspection.count });
    }
    await guard();
    const reply = await sink.seal({ id, state });
    sealed = true;
    const descriptor = validatePasteDescriptor(reply.descriptor ?? reply);
    if (
      descriptor.id !== id ||
      descriptor.count !== inspection.count ||
      descriptor.policy !== inspection.policy ||
      JSON.stringify(descriptor.shape) !== JSON.stringify(inspection.shape) ||
      JSON.stringify(descriptor.conversion) !== JSON.stringify(inspection.conversion)
    )
      throw refusal('query-stale', 'Query sealed descriptor changed');
    await guard();
    return { descriptor, outputFingerprint: inspection.outputFingerprint };
  } catch (error) {
    if (id)
      try {
        if (sealed) await sink.prune({ state });
        else
          try {
            await sink.cancel({ id, state });
          } catch {
            await sink.prune({ state });
          }
      } catch {
        /* Trusted pruning must preserve accepted/uncertain attachments. */
      }
    throw error;
  }
}
