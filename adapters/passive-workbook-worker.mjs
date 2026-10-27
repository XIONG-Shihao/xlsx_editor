/** Isolated bounded parser owner. Only typed bounded projections leave this worker. */
import { parentPort, workerData } from 'node:worker_threads';
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import {
  admitPassiveContainer,
  validatePassiveCfb,
  PASSIVE_FORMAT_LIMITS as LIMITS,
} from '../core/passive-container.mjs';
import { passiveWorkbookModel } from '../core/passive-model.mjs';
import { inspectVbaSources, VBA_SOURCE_LIMITS } from '../core/vba-source.mjs';
import { pureVbaFunctions, evaluatePureVba, PURE_VBA_LIMITS } from '../core/pure-vba.mjs';
import { dataOnlyWorkbook } from '../core/passive-data-export.mjs';
import { cellPosition, rangeBounds } from '../core/workbook-commands.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const XLSX = createRequire(import.meta.url)('xlsx');
if (XLSX.version !== '0.20.3') throw refusal('internal-mismatch', 'Passive parser dependency identity disagrees');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  fail = (message) => {
    throw refusal('session-refused', 'Passive workbook: ' + message);
  };
const allowed = new Set([
  'info',
  'read',
  'scan',
  'query',
  'summary',
  'calculation-summary',
  'vba-list',
  'vba-code',
  'vba-functions',
  'vba-evaluate',
  'data-export',
  'objects',
]);
const checkFields = (input, fields) => {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input) ||
    Object.keys(input).some((key) => !fields.includes(key))
  )
    fail('invalid request fields');
};
const byteBound = (value) => {
  if (Buffer.byteLength(JSON.stringify(value)) > LIMITS.replyBytes) fail('reply exceeds1MiB; request a smaller range');
  return value;
};
const file = await fs.open(workerData.source, constants.O_RDONLY | constants.O_NOFOLLOW);
let bytes;
try {
  const stat = await file.stat();
  if (!stat.isFile() || stat.size > LIMITS.sourceBytes || stat.uid !== process.getuid())
    fail('source snapshot identity disagrees');
  bytes = await file.readFile();
} finally {
  await file.close();
}
if (hash(bytes) !== workerData.revision) fail('immutable source snapshot hash disagrees');
const admission = admitPassiveContainer(bytes, workerData.format);
if (workerData.format === 'xls') validatePassiveCfb(XLSX.CFB.read(bytes, { type: 'buffer' }), { workbook: true });
const parsed = XLSX.read(bytes, {
  type: 'buffer',
  dense: false,
  cellFormula: true,
  cellNF: true,
  cellStyles: true,
  cellDates: false,
  sheetStubs: true,
  bookVBA: true,
  bookDeps: false,
  WTF: true,
});
const vba = admission.vbaParts.length ? admission.parts[admission.vbaParts[0]] : parsed.vbaraw;
const model = passiveWorkbookModel(parsed, {
  format: workerData.format,
  parts: admission.parts,
  revision: workerData.revision,
});
const view = {
  revision: workerData.revision,
  format: workerData.format,
  originalFormat: workerData.format,
  readOnly: true,
  passiveNotice: admission.notice,
  limitations: [
    admission.notice,
    'Cell view only; drawings, embedded objects and unsupported formatting remain preserved in the original same-format copy.',
    'VBA source inspection is passive; no macro, external data, refresh or workbook recalculation runs.',
  ],
  calculation: 'Stored values only',
  calculationEngine: 'passive-sheetjs/0.20.3',
  structuralReason: 'This original format is read-only.',
  sheets: model.sheets,
  features: { passive: true, vbaPresent: !!vba },
  storage: { kind: 'rust-indexed-v1', sessionId: workerData.id, state: '0' },
};
let project = null,
  projectError = null;
function sheet(name) {
  if (typeof name !== 'string' || !model.records.has(name)) fail('unknown worksheet');
  return model.records.get(name);
}
function afterPosition(value) {
  if (value === undefined || value === null) return null;
  if (typeof value !== 'string') fail('invalid page cursor');
  return cellPosition(value);
}
const beyond = (entry, position) =>
  !position || entry.row > position.row || (entry.row === position.row && entry.col > position.col);
const inside = (entry, bounds) =>
  entry.row >= bounds.first.row &&
  entry.row <= bounds.last.row &&
  entry.col >= bounds.first.col &&
  entry.col <= bounds.last.col;
function range(value, max = Infinity) {
  if (typeof value !== 'string') fail('invalid cell range');
  const bounds = rangeBounds(value);
  if (bounds.size > max) fail('range exceeds its position budget');
  return bounds;
}
function page(entries, { after, bounds, limit = LIMITS.readPositions, workLimit = 8192, match } = {}) {
  const cursor = afterPosition(after),
    cells = {};
  let scanned = 0,
    matches = 0,
    low = 0,
    high = entries.length;
  if (cursor)
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (beyond(entries[middle], cursor)) high = middle;
      else low = middle + 1;
    }
  let index = low;
  while (index < entries.length && scanned < workLimit && matches < limit) {
    const entry = entries[index++];
    scanned++;
    if (bounds && !inside(entry, bounds)) continue;
    if (!match || match(entry)) {
      cells[entry.ref] = entry.cell;
      matches++;
    }
  }
  const nextCursor = index < entries.length ? entries[index - 1].ref : null;
  return byteBound({ cells, matches, scanned, nextCursor });
}
function projectSources() {
  if (!vba) return null;
  if (projectError) throw Error(projectError);
  if (!project) {
    try {
      const cfb = validatePassiveCfb(XLSX.CFB.read(vba, { type: 'buffer' })),
        raw = inspectVbaSources(cfb);
      project = {
        ...raw,
        modules: raw.modules.map((module, index) => ({
          ...module,
          id: hash(Buffer.from(workerData.id + '\0' + workerData.revision + '\0' + index + '\0' + module.name)),
          sha256: hash(Buffer.from(module.code, 'utf8')),
        })),
      };
    } catch (error) {
      projectError = 'VBA source inspection refused: ' + String(error.message ?? error).slice(0, 4096);
      throw Error(projectError);
    }
  }
  return project;
}
const integer = (value, low, high, label) => {
  if (!Number.isSafeInteger(value) || value < low || value > high) fail('invalid ' + label);
  return value;
};
function sourceModule(moduleId) {
  if (typeof moduleId !== 'string' || !/^[a-f0-9]{64}$/.test(moduleId)) fail('invalid module identity');
  const module = projectSources()?.modules.find((item) => item.id === moduleId);
  if (!module) fail('unknown module identity');
  return module;
}
async function handle(request) {
  checkFields(request, ['id', 'operation', 'args']);
  if (!Number.isSafeInteger(request.id) || !allowed.has(request.operation)) fail('unsupported request');
  const args = request.args ?? {};
  if (request.operation === 'info') {
    checkFields(args, []);
    return byteBound(view);
  }
  if (request.operation === 'calculation-summary') {
    checkFields(args, []);
    return { revision: workerData.revision, engine: view.calculationEngine, ...model.summary };
  }
  if (request.operation === 'objects') {
    checkFields(args, ['sheet', 'after', 'limit']);
    sheet(args.sheet);
    if (args.after !== undefined && typeof args.after !== 'string') fail('invalid object cursor');
    if (args.limit !== undefined) integer(args.limit, 1, 100, 'object page limit');
    return {
      objects: [],
      total: 0,
      nextCursor: null,
      protectedReason: 'Objects remain in the original bytes; this passive cell view does not render their graphs.',
    };
  }
  if (request.operation === 'read') {
    checkFields(args, ['sheet', 'range']);
    const entries = sheet(args.sheet),
      bounds = range(args.range, LIMITS.readPositions),
      cells = {};
    for (const entry of entries) if (inside(entry, bounds)) cells[entry.ref] = entry.cell;
    return byteBound({ sheet: args.sheet, range: args.range, cells });
  }
  if (request.operation === 'scan') {
    checkFields(args, ['sheet', 'after']);
    const result = page(sheet(args.sheet), { after: args.after });
    return { cells: result.cells, nextCursor: result.nextCursor };
  }
  if (request.operation === 'query') {
    checkFields(args, ['sheet', 'query', 'mode', 'after', 'range', 'matchLimit']);
    if (
      typeof args.query !== 'string' ||
      Buffer.byteLength(args.query) > 32768 ||
      !['display', 'raw', 'formula'].includes(args.mode)
    )
      fail('invalid search');
    const needle = args.query.toLowerCase(),
      limit =
        args.matchLimit === undefined
          ? LIMITS.readPositions
          : integer(args.matchLimit, 1, LIMITS.readPositions, 'match limit');
    return page(sheet(args.sheet), {
      after: args.after,
      bounds: range(args.range ?? 'A1:XFD1048576'),
      limit,
      match: (entry) => {
        const cell = entry.cell,
          value =
            args.mode === 'formula'
              ? (cell.formula ?? '')
              : args.mode === 'raw'
                ? (cell.rawValue ?? cell.value)
                : cell.calculation?.status === 'stored'
                  ? cell.calculation.value
                  : cell.value;
        return String(value ?? '')
          .toLowerCase()
          .includes(needle);
      },
    });
  }
  if (request.operation === 'summary') {
    checkFields(args, ['sheet', 'range']);
    const bounds = range(args.range, LIMITS.storedCells);
    let count = 0,
      numericCount = 0,
      errors = 0,
      unresolved = 0,
      unsafeNumbers = 0,
      sum = 0,
      correction = 0,
      min = null,
      max = null;
    for (const entry of sheet(args.sheet)) {
      if (!inside(entry, bounds)) continue;
      const cell = entry.cell,
        formula = cell.formula !== undefined || cell.formulaType !== undefined;
      if (!formula && (cell.value === null || cell.value === '')) continue;
      count++;
      if (cell.dataType === 'error' || cell.error) {
        errors++;
        continue;
      }
      if (formula && cell.calculation?.status !== 'stored') {
        unresolved++;
        continue;
      }
      const value = formula ? cell.calculation.value : cell.value;
      if (typeof value !== 'number') continue;
      if (
        !Number.isFinite(value) ||
        (typeof cell.rawValue === 'string' &&
          /^[+-]?\d+$/.test(cell.rawValue) &&
          Math.abs(value) > Number.MAX_SAFE_INTEGER)
      ) {
        unsafeNumbers++;
        continue;
      }
      numericCount++;
      const adjusted = value - correction,
        next = sum + adjusted;
      correction = next - sum - adjusted;
      sum = next;
      if (!Number.isFinite(sum)) unsafeNumbers++;
      min = min === null ? value : Math.min(min, value);
      max = max === null ? value : Math.max(max, value);
    }
    const safe = unsafeNumbers === 0;
    return {
      count,
      numericCount,
      sum: safe ? sum : null,
      average: safe && numericCount ? sum / numericCount : null,
      min: safe ? min : null,
      max: safe ? max : null,
      errors,
      unresolved,
      unsafeNumbers,
      policy: 'Stored numeric caches and literals; no workbook recalculation',
    };
  }
  if (request.operation === 'vba-list') {
    checkFields(args, ['offset', 'limit']);
    const offset = integer(args.offset ?? 0, 0, VBA_SOURCE_LIMITS.modules, 'module offset'),
      limit = integer(args.limit ?? 100, 1, 100, 'module page limit');
    let value;
    try {
      value = projectSources();
    } catch (error) {
      return { modules: [], total: 0, nextOffset: null, protectedReason: error.message, execution: 'disabled' };
    }
    if (!value) return { modules: [], total: 0, nextOffset: null, execution: 'disabled' };
    const modules = value.modules
      .slice(offset, offset + limit)
      .map(({ code, ...metadata }) => ({ ...metadata, length: code.length }));
    return byteBound({
      modules,
      total: value.modules.length,
      nextOffset: offset + limit < value.modules.length ? offset + limit : null,
      projectName: value.projectName,
      codepage: value.codepage,
      execution: 'disabled',
    });
  }
  if (request.operation === 'vba-code') {
    checkFields(args, ['moduleId', 'offset', 'limit']);
    const module = sourceModule(args.moduleId);
    const offset = integer(args.offset ?? 0, 0, module.code.length, 'source offset'),
      limit = integer(args.limit ?? VBA_SOURCE_LIMITS.chunkText, 1, VBA_SOURCE_LIMITS.chunkText, 'source chunk size');
    return byteBound({
      id: module.id,
      name: module.name,
      code: module.code.slice(offset, offset + limit),
      offset,
      total: module.code.length,
      nextOffset: offset + limit < module.code.length ? offset + limit : null,
      sha256: module.sha256,
      execution: 'disabled',
    });
  }
  if (request.operation === 'vba-functions') {
    checkFields(args, ['moduleId']);
    const module = sourceModule(args.moduleId);
    let functions = [],
      protectedReason;
    try {
      if (module.kind !== 'standard') fail('document/class modules cannot be evaluated');
      functions = pureVbaFunctions(module.code);
    } catch (error) {
      protectedReason = String(error.message ?? error).slice(0, 4096);
    }
    return {
      moduleId: module.id,
      sourceSha256: module.sha256,
      functions,
      limits: PURE_VBA_LIMITS,
      execution: 'explicit-pure-numeric-only',
      ...(protectedReason ? { protectedReason } : {}),
    };
  }
  if (request.operation === 'vba-evaluate') {
    checkFields(args, ['moduleId', 'functionName', 'args']);
    const module = sourceModule(args.moduleId);
    if (module.kind !== 'standard') fail('document/class modules cannot be evaluated');
    return {
      ...evaluatePureVba(module.code, args.functionName, args.args),
      moduleId: module.id,
      sourceSha256: module.sha256,
      sourceRevision: workerData.revision,
      execution: 'explicit-pure-numeric-only',
    };
  }
  if (request.operation === 'data-export') {
    checkFields(args, ['format', 'exportId']);
    if (
      !['xls', 'xlsb', 'xlsx'].includes(args.format) ||
      typeof args.exportId !== 'string' ||
      !/^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(args.exportId)
    )
      fail('invalid fresh export identity');
    const { bytes, ...metadata } = dataOnlyWorkbook(XLSX, model, args.format),
      filename = 'data-' + args.exportId + '.' + args.format,
      output = await fs.open(
        path.join(path.dirname(workerData.source), filename),
        constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
        0o600
      );
    try {
      await output.writeFile(bytes);
      await output.sync();
      await output.chmod(0o400);
    } finally {
      await output.close();
    }
    return {
      ...metadata,
      filename,
      revision: hash(bytes),
      bytes: bytes.byteLength,
      sourceRevision: workerData.revision,
    };
  }
  fail('unsupported operation');
}
parentPort.on('message', async (request) => {
  try {
    if (Buffer.byteLength(JSON.stringify(request)) > 65536) fail('request byte budget exceeded');
    const result = await handle(request);
    parentPort.postMessage({ id: request.id, result });
  } catch (error) {
    parentPort.postMessage({ id: request?.id, error: String(error.message ?? error).slice(0, 8192) });
  }
});
