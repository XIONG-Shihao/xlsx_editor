/** Portable calculation adapter. Runtime bindings are supplied by the browser or
 * Node entrypoint. Stored values/caches remain intact; results are separate. */
export const CALCULATION_ENGINE = 'univer-oss/0.25.2';
import { QUALIFIED_DATE_FUNCTIONS, calculateDateScalar } from './date-calculation.mjs';
import {
  FINANCIAL_FUNCTIONS,
  FINANCIAL_REASON,
  financialCalls,
  checkFinancialFormula,
  calculateFinancialScalar,
  rewriteFinancialFormula,
} from './financial-calculation.mjs';
import { ARRAY_CALCULATION_REASON, arrayFormulaContext } from './array-calculation.mjs';
import { refusal } from './refusal-codes.mjs';
const ERROR_VALUES = new Set([
  '#DIV/0!',
  '#NAME?',
  '#VALUE!',
  '#NUM!',
  '#N/A',
  '#CYCLE!',
  '#REF!',
  '#SPILL!',
  '#CALC!',
  '#ERROR!',
  '#GETTING_DATA',
  '#NULL!',
]);
// The pinned lexer calls error-reference literals FUNCTION nodes. Accept only
// the complete local #REF! spelling here; calls, external/3D prefixes and text
// still pass through their existing function/name/type admission rules.
const brokenReferenceToken = (token) => /^(?:#REF!|(?:[A-Z0-9_.]+|'(?:[^'\[\]:]|'')+')!#REF!)$/i.test(token);
const key = (sheet, row, column) => `${sheet}:${row}:${column}`;
const isFormulaCell = (cell) => Boolean(cell.formula || cell.formulaType);
const metadataName = (name) => /^_xlnm\.(?:Print_Area|Print_Titles|_FilterDatabase)$/i.test(name);
/** The pinned lexer stops at a comma in a grouped multi-area reference. Inspect
 * authored syntax before dependency parsing so its truncated prefix cannot
 * certify a complete dependency graph or a guessed #VALUE! result. */
function referenceUnion(formula) {
  const text = String(formula),
    frames = [];
  let quote = '',
    square = 0,
    array = 0;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quote) {
      if (char === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '[') {
      square++;
      continue;
    }
    if (char === ']') {
      square = Math.max(0, square - 1);
      continue;
    }
    if (square) continue;
    if (char === '{') {
      array++;
      continue;
    }
    if (char === '}') {
      array = Math.max(0, array - 1);
      continue;
    }
    if (array) continue;
    if (char === '(') {
      let end = i - 1;
      while (end >= 0 && /\s/.test(text[end])) end--;
      let start = end;
      while (start >= 0 && /[A-Za-z0-9_.\\]/.test(text[start])) start--;
      frames.push(/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(text.slice(start + 1, end + 1)));
    } else if (char === ')') frames.pop();
    else if (char === ',' && !frames.at(-1)) return true;
  }
  return false;
}
function quotedDateArgument(formula) {
  const text = String(formula),
    frames = [];
  let quote = '',
    square = 0;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quote) {
      if (char === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (char === '"') {
      if (QUALIFIED_DATE_FUNCTIONS.has([...frames].reverse().find(Boolean))) return true;
      quote = char;
      continue;
    }
    if (char === "'") {
      quote = char;
      continue;
    }
    if (char === '[') {
      square++;
      continue;
    }
    if (char === ']') {
      square = Math.max(0, square - 1);
      continue;
    }
    if (square) continue;
    if (char === '(') {
      let end = i - 1;
      while (end >= 0 && /\s/.test(text[end])) end--;
      let start = end;
      while (start >= 0 && /[A-Za-z0-9_.\\]/.test(text[start])) start--;
      frames.push(text.slice(start + 1, end + 1).toUpperCase());
    } else if (char === ')') frames.pop();
  }
  return false;
}
/** Conservative dependency closure over authored names. Unused print/filter or
 * external names must not prevent unrelated formulas from calculating. Same-name
 * bindings are retained together so ambiguous scopes still fail closed. */
function calculationNames(view) {
  const names = (view.namedRanges ?? []).filter((name) => !name.scopeInvalid),
    used = new Set();
  const formulas = view.sheets.flatMap((sheet) =>
    Object.values(sheet.cells)
      .filter(isFormulaCell)
      .map((cell) => cell.formula ?? '')
  );
  if (formulas.some((formula) => /\b(?:INDIRECT|OFFSET)\s*\(/i.test(formula))) return names;
  const patterns = names.map(
    (name) =>
      new RegExp(
        `(?<![\\p{L}\\p{N}_.])${name.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}\\p{N}_.])(?!\\s*\\()`,
        'iu'
      )
  );
  for (let index = 0; index < formulas.length; index++) {
    const text = String(formulas[index]).replace(/"(?:[^"]|"")*"/g, '');
    names.forEach((name, i) => {
      if (!used.has(i) && patterns[i].test(text)) {
        used.add(i);
        formulas.push(name.formula);
      }
    });
  }
  return names.filter((_, i) => used.has(i));
}
function containsFormulas(view) {
  for (const sheet of view.sheets) {
    for (const ref in sheet.cells) if (isFormulaCell(sheet.cells[ref])) return true;
  }
  return false;
}
function mapCells(cells, mapper) {
  const result = {};
  for (const ref in cells) result[ref] = mapper(cells[ref], ref);
  return result;
}
function numericMeaning(value) {
  const match = /^([+-]?)(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(String(value));
  if (!match) return null;
  let digits = (match[2] + (match[3] ?? '')).replace(/^0+/, '');
  let exponent = Number(match[4] ?? 0) - (match[3]?.length ?? 0);
  if (!digits) return '0';
  while (digits.endsWith('0')) {
    digits = digits.slice(0, -1);
    exponent++;
  }
  return `${match[1] === '-' ? '-' : ''}${digits}e${exponent}`;
}
export const hasExactNumericValue = (value, rawValue) =>
  rawValue == null || numericMeaning(rawValue) === numericMeaning(value);
function position(ref) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(ref);
  if (!match) throw refusal('internal-invalid-field', `Invalid cell reference: ${ref}`);
  return {
    row: Number(match[2]) - 1,
    column: [...match[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
  };
}
/** Retained Excel scalar oracle exceptions; keep upstream coercion/array rules.
 * Register executors so dependent formulas see corrected values too. */
function excelCompatibilityOverrides(libs, typedErrors, dateSystems, dateMisses, financialMisses, financialProofs) {
  const scalar = (value) =>
    value?.isArray() && value.getRowCount() === 1 && value.getColumnCount() === 1 ? value.get(0, 0) : value;
  const Trunc = libs.functionMath.find(([, name]) => name === 'TRUNC')[0];
  const Sum = libs.functionMath.find(([, name]) => name === 'SUM')[0];
  const Roundup = libs.functionMath.find(([, name]) => name === 'ROUNDUP')[0];
  class ExcelRoundup extends Roundup {
    calculate(number, digits) {
      const result = super.calculate(number, digits),
        input = scalar(number),
        precision = scalar(digits),
        rounded = scalar(result);
      if (!input?.isNumber() || !precision?.isNumber() || !rounded?.isNumber()) return result;
      const places = precision.getValue(),
        value = rounded.getValue();
      if (
        !Number.isFinite(input.getValue()) ||
        !Number.isFinite(value) ||
        !Number.isInteger(places) ||
        Math.abs(places) > 12
      )
        return result;
      // Keep Univer's decimal/epsilon selection of rounded units, then restore
      // their binary64 quantum by multiplication. Division loses the Excel
      // 16.113 oracle bit in ROUNDUP(-2.71,1): -28 * 0.1 != -28 / 10.
      const scaled = value * 10 ** places,
        units = Math.round(scaled);
      if (!Number.isSafeInteger(units) || Math.abs(scaled - units) > Number.EPSILON * Math.max(1, Math.abs(scaled)))
        return result;
      return libs.NumberValueObject.create(units * 10 ** -places);
    }
  }
  class TypedResult extends libs.BaseFunction {
    constructor(...args) {
      super(...args);
      this.minParams = 1;
      this.maxParams = 1;
    }
    calculate(value) {
      const cellKey = key(this.subUnitId, this.row, this.column),
        result = scalar(value);
      typedErrors.delete(cellKey);
      if (result?.isError()) typedErrors.set(cellKey, result.getValue());
      return value;
    }
  }
  class ExcelTrunc extends Trunc {
    calculate(number, digits) {
      const input = scalar(number),
        precision = scalar(digits);
      if (input.isNumber() && (!precision || precision.isNumber())) {
        const value = input.getValue();
        if (Number.isFinite(value)) {
          if (!precision || precision.getValue() === 0) return libs.NumberValueObject.create(Math.trunc(value) || 0);
          if (value < 0) {
            // The pinned executor adds a positive epsilon before truncating. Apply
            // its decimal precision logic to the magnitude, then restore the sign.
            const result = super.calculate(libs.NumberValueObject.create(-value), precision);
            if (!result.isNumber()) return result;
            const signed = -result.getValue();
            return libs.NumberValueObject.create(signed === 0 ? 0 : signed);
          }
        }
      }
      return super.calculate(number, digits);
    }
  }
  const dates = [...QUALIFIED_DATE_FUNCTIONS].map((name) => {
    const Original = libs.functionDate.find(([, key]) => key === name)[0];
    return [
      class extends Original {
        calculate(...args) {
          const values = args.map(scalar),
            failure = values.find((value) => value?.isError());
          if (failure) return failure;
          const result = dateSystems.has(this.subUnitId)
            ? calculateDateScalar(
                name,
                values.map((value) => (value?.isNumber() && !value.isDateFormat() ? value.getValue() : undefined)),
                dateSystems.get(this.subUnitId)
              )
            : { unresolved: 'Date calculation context is unavailable' };
          if (result.unresolved) {
            // Record the semantic miss at its owner before IFERROR can conceal the
            // engine sentinel. Sibling calls never erase a miss in the same cell.
            dateMisses.set(key(this.subUnitId, this.row, this.column), {
              id: this.subUnitId,
              row: this.row,
              column: this.column,
              reason: result.unresolved,
            });
            return libs.ErrorValueObject.create('#NAME?');
          }
          return result.error
            ? libs.ErrorValueObject.create(result.error)
            : libs.NumberValueObject.create(result.value);
        }
      },
      name,
    ];
  });
  const financial = [...FINANCIAL_FUNCTIONS].map((name) => [
    class extends libs.BaseFunction {
      constructor(...args) {
        super(...args);
        this.minParams = 3;
        this.maxParams = 5;
      }
      calculate(...args) {
        const values = args.map(scalar),
          cellKey = key(this.subUnitId, this.row, this.column);
        // The pinned lexer classifies long numeric tokens as strings. Recover only
        // a closed authored numeric token with owner-local input proof, never a text
        // argument/reference or a value inferred from provider coercion.
        const proof = financialProofs
          .get(cellKey)
          ?.find(
            (call) =>
              call.name === name &&
              call.values.length === values.length &&
              values.every((value, index) =>
                value?.isNumber()
                  ? Object.is(value.getValue(), call.values[index])
                  : value?.isString() &&
                    call.args[index].kind === 'number' &&
                    hasExactNumericValue(call.values[index], value.getValue())
              )
          );
        const result = proof
          ? calculateFinancialScalar(name, proof.values)
          : { unresolved: FINANCIAL_REASON + ': runtime input lacks authored numeric proof' };
        if (result.unresolved) {
          financialMisses.set(cellKey, {
            id: this.subUnitId,
            row: this.row,
            column: this.column,
            reason: result.unresolved,
          });
          return libs.ErrorValueObject.create('#NAME?');
        }
        return libs.NumberValueObject.create(result.value);
      }
    },
    name,
  ]);
  class FinancialCall extends libs.BaseFunction {
    constructor(...args) {
      super(...args);
      this.minParams = 1;
      this.maxParams = 33;
    }
    calculate(index, ...dependencies) {
      const cellKey = key(this.subUnitId, this.row, this.column),
        ordinal = scalar(index);
      // The private call-site token selects the independently proved authored
      // tree. Provider numeric equality never proves an argument expression.
      // Reference leaves remain arguments solely to retain complete graph edges.
      const proof =
        ordinal?.isNumber() && Number.isInteger(ordinal.getValue())
          ? financialProofs.get(cellKey)?.find((call) => call.index === ordinal.getValue())
          : undefined;
      const result = proof
        ? calculateFinancialScalar(proof.name, proof.values)
        : { unresolved: FINANCIAL_REASON + ': missing authored call-site proof' };
      if (result.unresolved) {
        financialMisses.set(cellKey, {
          id: this.subUnitId,
          row: this.row,
          column: this.column,
          reason: result.unresolved,
        });
        return libs.ErrorValueObject.create('#NAME?');
      }
      return libs.NumberValueObject.create(result.value);
    }
  }
  class ExcelSum extends Sum {
    calculate(...variants) {
      let total = 0;
      for (let value of variants) {
        if (value.isString()) value = value.convertToNumberObjectValue();
        if (value.isError()) return value;
        if (value.isArray()) {
          let error,
            unsupported = false;
          value.iterator((item) => {
            if (item == null || item.isString() || item.isBoolean() || item.isNull()) return true;
            if (item.isError()) {
              error = item;
              return false;
            }
            if (!item.isNumber()) {
              unsupported = true;
              return false;
            }
            total += item.getValue();
            return true;
          });
          if (error) return error;
          if (unsupported) return super.calculate(...variants);
        } else if (value.isNumber() || value.isBoolean()) total += Number(value.getValue());
        else if (!value.isNull()) return super.calculate(...variants);
      }
      return libs.NumberValueObject.create(total);
    }
  }
  // B8 scalar operator conversion: preserve array/function-specific rules.
  const arithmetic = (value) => {
    value = scalar(value);
    if (value?.isNumber()) return { n: value.getValue(), coerced: false };
    if (value?.isBoolean()) return { n: Number(value.getValue()), coerced: true };
    if (value?.isNull()) return { n: 0, coerced: false };
    if (!value?.isString()) return;
    let text = value.getValue().replace(/^ +| +$/g, '');
    if (!text || /^[A-Za-z]+$/.test(text) || /^[-+]{2}/.test(text)) return { error: '#VALUE!' };
    const percent = text.endsWith('%');
    if (percent) text = text.slice(0, -1);
    if (!/^[+\-0-9.eE,]+$/.test(text) || (text.includes(',') && /[eE]/.test(text))) return;
    if (text.includes(',')) {
      if (!/^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?$/.test(text)) return { error: '#VALUE!' };
      text = text.replaceAll(',', '');
    }
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) || (text.match(/\d/g) ?? []).length > 15) return;
    const n = Number(text) / (percent ? 100 : 1);
    if (!Number.isFinite(n) || Math.abs(n) > 999999999999999) return;
    return { n, coerced: true };
  };
  const operands = (a, b) => {
    const x = arithmetic(a),
      y = arithmetic(b);
    if (!x || !y) return;
    if (x.error || y.error) return { error: x.error ?? y.error };
    if ((x.coerced || y.coerced) && (Math.abs(x.n) > 999999999999999 || Math.abs(y.n) > 999999999999999)) return;
    return { x: x.n, y: y.n, coerced: x.coerced || y.coerced };
  };
  // B6 overrides operator executors, not the shared NumberValueObject
  // prototype: other function internals retain their independently pinned rules.
  const binary = (name, operation) => {
    const Original = libs.functionMeta.find(([, key]) => key === name)[0];
    return [
      class extends Original {
        calculate(a, b) {
          const pair = operands(a, b);
          if (pair?.error) return libs.ErrorValueObject.create(pair.error);
          if (pair && (name !== 'CUBE' || (pair.coerced && Number.isInteger(pair.x) && Number.isInteger(pair.y)))) {
            if (name === 'DIVIDED' && pair.y === 0) return libs.ErrorValueObject.create('#DIV/0!');
            return libs.NumberValueObject.create(operation(pair.x, pair.y));
          }
          return super.calculate(a, b);
        }
      },
      name,
    ];
  };
  const final = (name, subtract) => [
    class extends libs.BaseFunction {
      constructor(...args) {
        super(...args);
        this.minParams = 2;
        this.maxParams = 2;
      }
      calculate(a, b) {
        const pair = operands(a, b);
        if (pair?.error) return libs.ErrorValueObject.create(pair.error);
        if (pair) return libs.NumberValueObject.create(compensated(pair.x, subtract ? -pair.y : pair.y));
        const Original = libs.functionMeta.find(([, key]) => key === (subtract ? 'MINUS' : 'PLUS'))[0];
        return new Original(subtract ? 'MINUS' : 'PLUS').calculate(a, b);
      }
    },
    name,
  ];
  return [
    [ExcelTrunc, 'TRUNC'],
    ...dates,
    ...financial,
    [FinancialCall, 'PLATFORM_FINANCIAL_CALL'],
    [ExcelSum, 'SUM'],
    [ExcelRoundup, 'ROUNDUP'],
    [TypedResult, 'PLATFORM_TYPED_RESULT'],
    binary('PLUS', (a, b) => a + b),
    binary('MINUS', (a, b) => a - b),
    binary('MULTIPLY', (a, b) => a * b),
    binary('DIVIDED', (a, b) => a / b),
    binary('CUBE', (a, b) => a ** b),
    final('PLATFORM_B7_ADD', false),
    final('PLATFORM_B7_SUB', true),
  ];
}
// Calculation-only root wrapper. Stored and transformed formulas retain their
// authored text, including parentheses that suppress Excel's final compensation.
function calculationFormula(formula) {
  const expression = compensatedFormula(formula);
  // CellValueType has no ERROR member. Observe only quoted-text roots before
  // serialization conflates ErrorValueObject with an error-looking string.
  return /"(?:[^"]|"")*"/.test(expression) ? `PLATFORM_TYPED_RESULT(${expression})` : expression;
}
function compensatedFormula(formula) {
  const text = formula.replace(/^=/, '').trim();
  let depth = 0,
    brackets = 0,
    quote = '',
    at = -1;
  if (/^-\s*"(?:[^"]|"")*"$/.test(text)) return `PLATFORM_B7_SUB(0,${text.slice(1)})`;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quote) {
      if (c === quote) {
        if (text[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (c === '"' || c === "'") {
      quote = c;
      continue;
    }
    if (c === '[') {
      brackets++;
      continue;
    }
    if (c === ']') {
      brackets--;
      continue;
    }
    if (brackets) continue;
    if (c === '(') {
      depth++;
      continue;
    }
    if (c === ')') {
      depth--;
      continue;
    }
    if (depth) continue;
    if ('=<>&'.includes(c)) return text;
    if ((c === '+' || c === '-') && i > 0) {
      const before = text.slice(0, i).trimEnd(),
        prev = before.at(-1);
      if (!'+-*/^(,'.includes(prev) && !/[0-9][eE]$/.test(before)) at = i;
    }
  }
  return at < 0 ? text : `PLATFORM_B7_${text[at] === '+' ? 'ADD' : 'SUB'}(${text.slice(0, at)},${text.slice(at + 1)})`;
}
function compensated(a, b) {
  const n = a + b;
  if (Math.sign(a) !== Math.sign(b) && a !== 0 && Number.isFinite(a)) {
    const bytes = new ArrayBuffer(8),
      v = new DataView(bytes);
    v.setFloat64(0, Math.abs(a));
    v.setBigUint64(0, v.getBigUint64(0) + 1n);
    const ulp = v.getFloat64(0) - Math.abs(a);
    if (Number.isFinite(ulp) && Math.abs(n) < 8 * ulp) return 0;
  }
  return n;
}
// Separately loaded package copies can share one hoisted Univer AST cache.
// Namespace each copy; these IDs isolate calculation state, not authorization.
const calculationOwnerNamespace =
  globalThis.crypto?.randomUUID?.() ??
  (globalThis.crypto?.getRandomValues
    ? Array.from(globalThis.crypto.getRandomValues(new Uint32Array(4)), (n) => n.toString(16).padStart(8, '0')).join('')
    : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}-${Math.random().toString(36).slice(2)}`);
let calculationOwnerSequence = 0;
function createEngine(libs, { manual = false } = {}) {
  const {
    Univer,
    LocaleType,
    FUniver,
    UniverFormulaEnginePlugin,
    UniverSheetsPlugin,
    UniverSheetsFormulaPlugin,
    UniverSheetsTablePlugin,
    CalculationMode,
  } = libs;
  const univer = new Univer({ locale: LocaleType.EN_US, locales: { [LocaleType.EN_US]: {} } });
  const typedErrors = new Map(),
    dateSystems = new Map(),
    dateMisses = new Map(),
    financialMisses = new Map(),
    financialProofs = new Map(),
    financialAuthoredMisses = new Map();
  univer.onDispose(() => {
    typedErrors.clear();
    dateSystems.clear();
    dateMisses.clear();
    financialMisses.clear();
    financialProofs.clear();
    financialAuthoredMisses.clear();
  });
  univer.registerPlugin(UniverFormulaEnginePlugin, {
    function: excelCompatibilityOverrides(libs, typedErrors, dateSystems, dateMisses, financialMisses, financialProofs),
  });
  univer.registerPlugin(UniverSheetsPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin, {
    initialFormulaComputing: manual ? CalculationMode.NO_CALCULATION : CalculationMode.FORCED,
  });
  univer.registerPlugin(UniverSheetsTablePlugin);
  // The pinned engine caches AST nodes globally by unit ID and formula. A
  // private ID keeps cached executors/proofs inside their owning live engine.
  // The trailing delimiter also avoids its prefix-based disposal matching 1/10.
  return {
    univer,
    api: FUniver.newAPI(univer),
    typedErrors,
    dateSystems,
    dateMisses,
    financialMisses,
    financialProofs,
    financialAuthoredMisses,
    workbookId: `workbook-${calculationOwnerNamespace}-${++calculationOwnerSequence}:`,
  };
}
function literal(cell, libs) {
  const value = cell.value;
  const t =
    cell.dataType === 'error'
      ? libs.CellValueType.ERROR
      : cell.dataType === 'boolean' || typeof value === 'boolean'
        ? libs.CellValueType.BOOLEAN
        : cell.dataType === 'text' || typeof value === 'string'
          ? libs.CellValueType.FORCE_STRING
          : libs.CellValueType.NUMBER;
  return { v: value ?? null, t };
}
function rangeIndex(ranges) {
  ranges.sort((a, b) => a.startRow - b.startRow);
  const build = (start, end) => {
    if (start >= end) return null;
    const middle = (start + end) >> 1,
      range = ranges[middle];
    const left = build(start, middle),
      right = build(middle + 1, end);
    return { range, left, right, maxRow: Math.max(range.endRow, left?.maxRow ?? -1, right?.maxRow ?? -1) };
  };
  return build(0, ranges.length);
}
function containsPoint(node, row, column) {
  if (!node || node.maxRow < row) return false;
  if (containsPoint(node.left, row, column)) return true;
  const range = node.range;
  if (range.startRow > row) return false;
  if (row <= range.endRow && column >= range.startColumn && column <= range.endColumn) return true;
  return containsPoint(node.right, row, column);
}
/** This only reduces the engine's copy, never the accepted workbook. All formula
 * cells remain present. Unproved grammar falls back to the complete snapshot. */
function inputScope(engine, libs, view, enabled, omitBlocked = false) {
  const full = (reason) => ({ mode: 'full', reason, includes: () => true });
  if (!enabled) return full('Complete input requested');
  const scopes = new Map(view.sheets.map((sheet) => [sheet.name, { exact: new Set(), ranges: new Map() }]));
  const names = calculationNames(view),
    tables = view.tables ?? [];
  const tableNames = new Set(tables.map((table) => table.name.toUpperCase()));
  const parser = engine.api.getFormula();
  // Executors register during workbook creation. The pinned exported catalog
  // only selects the optimization here; actual support is checked after load.
  const knownFunctions = new Set(
    Object.entries(libs)
      .filter(([name]) => name.startsWith('FUNCTION_NAMES_'))
      .flatMap(([, names]) => Object.values(names))
      .filter((name) => typeof name === 'string')
      .map((name) => name.toUpperCase())
  );
  // SUMIF/AVERAGEIF resize a supplied result range to match their criteria
  // range, so their actual inputs can extend beyond the parsed references.
  const contextualFunctions = new Set(['CELL', 'INFO', 'SUMIF', 'AVERAGEIF']);
  function add(reference, owners) {
    if (reference.unitId || reference.sheetName?.includes(':'))
      throw refusal('formula-refused', 'External or three-dimensional reference');
    let { startRow, endRow, startColumn, endColumn } = reference.range;
    if (Number.isNaN(startRow) && Number.isNaN(endRow)) {
      startRow = 0;
      endRow = Infinity;
    }
    if (Number.isNaN(startColumn) && Number.isNaN(endColumn)) {
      startColumn = 0;
      endColumn = Infinity;
    }
    if (
      ![startRow, startColumn].every((value) => Number.isInteger(value) && value >= 0) ||
      ![endRow, endColumn].every((value) => value === Infinity || (Number.isInteger(value) && value >= 0)) ||
      endRow < startRow ||
      endColumn < startColumn
    )
      throw refusal('formula-malformed', 'Unproved range bounds');
    for (const owner of reference.sheetName ? [reference.sheetName] : owners) {
      const scope = scopes.get(owner);
      if (!scope) throw refusal('formula-unsupported', 'Unresolved sheet reference');
      if (startRow === endRow && startColumn === endColumn) scope.exact.add(address(startRow, startColumn));
      else
        scope.ranges.set(`${startRow}:${endRow}:${startColumn}:${endColumn}`, {
          startRow,
          endRow,
          startColumn,
          endColumn,
        });
    }
  }
  function inspect(formula, owners) {
    if (referenceUnion(formula))
      throw refusal('formula-unsupported', 'Multi-area reference unions require qualified native calculation');
    if (arrayFormulaContext(formula).conservative) throw refusal('formula-refused', ARRAY_CALCULATION_REASON);
    if (volatileFormula(formula)) throw refusal('formula-refused', 'Dynamic or volatile reference context');
    const nodes = parser.sequenceNodesBuilder('=' + String(formula).replace(/^=/, ''));
    const knownNames = new Set(
      names
        .filter((name) => !name.scopeInvalid && (!name.sheet || owners.includes(name.sheet)))
        .map((name) => name.name.toUpperCase())
    );
    for (let index = 0; index < nodes.length; index++) {
      const node = nodes[index];
      // A1:INDEX(...) constructs a larger range than its individual tokens. Spill
      // and implicit-intersection contexts likewise retain the complete input.
      if (typeof node === 'string') {
        if ([':', '#', '@'].includes(node)) throw refusal('formula-refused', 'Computed reference operator');
        continue;
      }
      if (node.nodeType === libs.sequenceNodeType.REFERENCE) add(libs.deserializeRangeWithSheet(node.token), owners);
      else if (node.nodeType === libs.sequenceNodeType.FUNCTION) {
        const token = node.token.toUpperCase();
        if (nodes[index + 1] === '(') {
          if (!knownFunctions.has(token) || contextualFunctions.has(token))
            throw refusal('formula-malformed', 'Unproved function context');
        } else if (
          !knownNames.has(token) &&
          !['TRUE', 'FALSE'].includes(token) &&
          !tableNames.has(token.split('[')[0])
        ) {
          throw refusal('formula-unsupported', 'Unresolved name or table');
        }
      } else if (node.nodeType === libs.sequenceNodeType.DEFINED_NAME) {
        if (!knownNames.has(node.token.toUpperCase())) throw refusal('formula-unsupported', 'Unresolved defined name');
      } else if (node.nodeType === libs.sequenceNodeType.TABLE) {
        if (!tableNames.has(node.token.toUpperCase().split('[')[0]))
          throw refusal('formula-unsupported', 'Unresolved table');
      } else if (
        ![libs.sequenceNodeType.NUMBER, libs.sequenceNodeType.STRING, libs.sequenceNodeType.ARRAY].includes(
          node.nodeType
        )
      ) {
        throw refusal('formula-malformed', 'Unproved parser node');
      }
    }
  }
  try {
    // Include every named/table range, even when unused. Workbook-relative names
    // conservatively retain their ranges on every sheet; local names keep scope.
    for (const name of names)
      if (!omitBlocked || !volatileFormula(name.formula))
        inspect(name.formula, name.sheet ? [name.sheet] : [...scopes.keys()]);
    for (const table of tables) add(libs.deserializeRangeWithSheet(table.ref), [table.sheet]);
    for (const sheet of view.sheets) {
      for (const ref in sheet.cells) {
        const cell = sheet.cells[ref];
        if (cell.formula && (!omitBlocked || (!cell.formulaType && !volatileFormula(cell.formula))))
          inspect(cell.formula, [sheet.name]);
      }
    }
    for (const scope of scopes.values()) scope.index = rangeIndex([...scope.ranges.values()]);
    return {
      mode: 'static-references',
      reason: 'All formulas and occupied static/name/table references retained',
      ranges: [...scopes].flatMap(([sheet, scope]) =>
        [...scope.exact]
          .map((ref) => ({ sheet, ref }))
          .concat([...scope.ranges.values()].map((range) => ({ sheet, ...range })))
      ),
      includes: (sheet, ref, row, column) => {
        const scope = scopes.get(sheet);
        return scope.exact.has(ref.toUpperCase()) || containsPoint(scope.index, row, column);
      },
    };
  } catch (error) {
    return full(error.message);
  }
}
/** Index-backed owners use the same conservative dependency selection as the
 * evaluator, before fetching literal inputs from their native cell store. */
export function calculationInputRangesWith(libs, view) {
  const engine = createEngine(libs, { manual: true });
  try {
    const { includes, ...scope } = inputScope(engine, libs, view, true, true);
    return scope;
  } finally {
    engine.univer.dispose();
  }
}
/** Reuse one parser/engine while proving each formula's static dependencies. */
export function calculationDependenciesWith(libs, view, { conservative = false } = {}) {
  const engine = createEngine(libs, { manual: true });
  try {
    const result = [];
    for (const sheet of view.sheets)
      for (const [ref, cell] of Object.entries(sheet.cells))
        if (cell.formula || cell.formulaType) {
          const isolated = {
            ...view,
            sheets: view.sheets.map((s) => ({ ...s, cells: s.name === sheet.name ? { [ref]: cell } : {} })),
          };
          const scope = cell.formulaType ? { mode: 'full' } : inputScope(engine, libs, isolated, true);
          if (scope.mode !== 'static-references' && !conservative) return null;
          // An unproved formula watches every cell, without discarding proven edges
          // belonging to other formulas. Use full grid bounds, including future cells.
          result.push({
            sheet: sheet.name,
            ref,
            ranges:
              scope.mode === 'static-references'
                ? scope.ranges
                : view.sheets.map((s) => ({
                    sheet: s.name,
                    startRow: 0,
                    endRow: 1048575,
                    startColumn: 0,
                    endColumn: 16383,
                  })),
          });
        }
    return result;
  } finally {
    engine.univer.dispose();
  }
}
async function load(
  engine,
  libs,
  view,
  { projectReferences = false, groupNamedExpressions = false, onInputProjection } = {}
) {
  for (const sheet of view.sheets)
    if (sheet.date1904 !== undefined && typeof sheet.date1904 !== 'boolean')
      throw refusal('internal-invalid-field', 'Invalid workbook date-system metadata');
  const scope = inputScope(engine, libs, view, projectReferences, projectReferences),
    projection = [];
  const ids = new Map(view.sheets.map((sheet, index) => [sheet.name, `sheet-${index}`]));
  for (const sheet of view.sheets) engine.dateSystems.set(ids.get(sheet.name), !!sheet.date1904);
  const unsafe = [];
  const unsupported = new Map();
  const dateFunctions = new Set(
    Object.values(libs.FUNCTION_NAMES_DATE).filter((name) => !QUALIFIED_DATE_FUNCTIONS.has(name))
  );
  const containsDateFunction = (formula) =>
    [...String(formula).matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].some((match) =>
      dateFunctions.has(match[1].toUpperCase())
    );
  const dateNames = (view.namedRanges ?? []).some((name) => containsDateFunction(name.formula));
  const volatileNames = (view.namedRanges ?? []).some((name) => volatileFormula(name.formula));
  const snapshot = { id: engine.workbookId, name: 'Workbook', sheetOrder: [...ids.values()], sheets: {} };
  for (const sheet of view.sheets) {
    const id = ids.get(sheet.name),
      cellData = {};
    let rowCount = Math.max(sheet.rows ?? 20, 20),
      columnCount = Math.max(sheet.columns ?? 8, 8);
    let sourceCells = 0,
      engineCells = 0;
    for (const ref in sheet.cells) {
      const cell = sheet.cells[ref];
      sourceCells++;
      const { row, column } = position(ref);
      rowCount = Math.max(rowCount, row + 1);
      columnCount = Math.max(columnCount, column + 1);
      if (!isFormulaCell(cell) && !scope.includes(sheet.name, ref, row, column)) continue;
      engineCells++;
      const financial = cell.formula ? checkFinancialFormula(cell.formula, sheet.cells) : undefined;
      const authoredKey = key(id, row, column);
      if (financial?.calls?.length) engine.financialProofs.set(authoredKey, financial.calls);
      let callSites = [];
      try {
        if (cell.formula) callSites = financialCalls(cell.formula);
      } catch {}
      const executionFormula = callSites.length ? rewriteFinancialFormula(cell.formula, callSites) : cell.formula;
      (cellData[row] ??= {})[column] = cell.formulaType
        ? { f: '=PLATFORM_UNSUPPORTED_SPECIAL_FORMULA()' }
        : cell.formula
          ? {
              f:
                '=' +
                (groupNamedExpressions ? calculationFormula(executionFormula) : executionFormula.replace(/^=/, '')),
            }
          : literal(cell, libs);
      if (cell.formula && arrayFormulaContext(cell.formula).blocked)
        cellData[row][column] = { f: '=PLATFORM_UNSUPPORTED_ARRAY_FORMULA()' };
      if (cell.formulaType) unsafe.push({ id, row, column });
      if (cell.formulaType === 'array' && cell.formulaReference) {
        const [first, last = first] = cell.formulaReference.split(':'),
          a = position(first),
          b = position(last);
        unsafe.push({
          id,
          row,
          column,
          range: { startRow: a.row, endRow: b.row, startColumn: a.column, endColumn: b.column },
        });
      }
      if (cell.formula && arrayFormulaContext(cell.formula).blocked)
        unsafe.push({
          id,
          row,
          column,
          range: { startRow: row, endRow: 1048575, startColumn: column, endColumn: 16383 },
        });
      if (sheet.date1904 && cell.formula && (dateNames || containsDateFunction(cell.formula))) {
        unsupported.set(key(id, row, column), 'Date functions in the 1904 date system are not qualified');
        unsafe.push({ id, row, column });
      }
      if (cell.formula) {
        const result = checkFinancialFormula(cell.formula, sheet.cells);
        if (result.calls?.length) engine.financialProofs.set(key(id, row, column), result.calls);
        if (result.unresolved)
          engine.financialAuthoredMisses.set(key(id, row, column), { id, row, column, reason: result.unresolved });
      }
      if (cell.formula && quotedDateArgument(cell.formula)) {
        unsupported.set(key(id, row, column), 'Quoted date arguments require separate coercion/locale qualification');
        unsafe.push({ id, row, column });
      }
      if (cell.formula && (volatileNames || volatileFormula(cell.formula))) {
        unsupported.set(key(id, row, column), 'Volatile calculation requires a qualified clock and random seed');
        unsafe.push({ id, row, column });
        cellData[row][column] = { f: '=PLATFORM_UNSUPPORTED_VOLATILE_FORMULA()' };
      }
      // Univer converts error-looking text to errors, even FORCE_STRING; precision
      // outside safe integers likewise cannot be certified from JS numeric values.
      if (
        !cell.formula &&
        (((cell.dataType === 'text' || (!cell.dataType && typeof cell.value === 'string')) &&
          ERROR_VALUES.has(cell.value)) ||
          (cell.dataType === 'number' &&
            typeof cell.rawValue === 'string' &&
            numericMeaning(cell.rawValue) !== numericMeaning(Number(cell.rawValue))))
      )
        unsafe.push({ id, row, column });
    }
    snapshot.sheets[id] = { id, name: sheet.name, rowCount, columnCount, cellData };
    projection.push({ sheet: sheet.name, sourceCells, engineCells, rowCount, columnCount });
  }
  onInputProjection?.({ mode: scope.mode, reason: scope.reason, sheets: projection });
  const book = engine.api.createWorkbook(snapshot);
  // No Univer UI is mounted: advance its pinned plugin lifecycle so browser
  // calculation and reference controllers initialize identically to Node.
  engine.univer.__getInjector().get(libs.LifecycleService).stage = libs.LifecycleStages.Rendered;
  for (const name of groupNamedExpressions ? calculationNames(view) : (view.namedRanges ?? [])) {
    if (name.scopeInvalid) {
      if (!groupNamedExpressions)
        throw refusal('internal-invalid-field', 'Invalid defined-name scope cannot be transformed');
      continue;
    }
    // Print/filter metadata stays in the OOXML model. The pinned engine treats
    // duplicate built-in names on different sheets as duplicate user names.
    if (groupNamedExpressions && metadataName(name.name)) continue;
    const target = name.sheet ? book.getSheetBySheetId(ids.get(name.sheet)) : book;
    if (!target) throw refusal('internal-unknown-reference', `Unknown defined-name sheet: ${name.sheet}`);
    let formula = name.formula;
    if (arrayFormulaContext(formula).blocked) formula = 'PLATFORM_UNSUPPORTED_ARRAY_FORMULA()';
    // The pinned engine substitutes a defined expression into its caller without
    // grouping: Rate=A1+1 makes Rate*3 become A1+1*3. Group compound expressions
    // for evaluation only; reference-transform loads retain the authored text.
    if (groupNamedExpressions) {
      const expression = String(formula).replace(/^=/, '');
      const nodes = engine.api.getFormula().sequenceNodesBuilder('=' + expression);
      if (nodes.length > 1) formula = `(${expression})`;
    }
    target.insertDefinedName(name.name, formula);
  }
  for (const [index, table] of (view.tables ?? []).entries()) {
    if ((table.headerRowCount ?? 1) !== 1 || (table.totalsRowCount ?? 0) !== 0)
      throw refusal('formula-unsupported', `Table ${table.name} has an unsupported header or totals configuration`);
    const sheet = book.getSheetBySheetId(ids.get(table.sheet));
    if (!sheet) throw refusal('internal-unknown-reference', `Unknown table sheet: ${table.sheet}`);
    const columns = (table.columns ?? []).map((column, i) => ({
      id: `column-${i}`,
      displayName: typeof column === 'string' ? column : column.name,
    }));
    if (
      !(await sheet.addTable(
        table.name,
        sheet.getRange(table.ref).getRange(),
        `table-${index}`,
        columns.length ? { columns } : undefined
      ))
    )
      throw refusal('formula-bad-input', `Could not map table ${table.name}`);
  }
  // An unknown function can be hidden by IFERROR/ISERROR. Check the parsed
  // expression against the installed registry before accepting any result.
  const functions = engine.univer.__getInjector().get(libs.IFunctionService),
    parser = engine.api.getFormula();
  const definedNames = view.namedRanges ?? [];
  const knownNamesFor = (sheet) =>
    new Set(
      definedNames
        .filter((name) => !name.scopeInvalid && (!name.sheet || name.sheet === sheet))
        .map((name) => name.name.toUpperCase())
    );
  const tableNames = new Set((view.tables ?? []).map((table) => table.name.toUpperCase()));
  const unknown = (formula, sheet) => {
    const knownNames = knownNamesFor(sheet),
      nodes = parser.sequenceNodesBuilder('=' + String(formula).replace(/^=/, ''));
    return nodes.some((node, index) => {
      if (typeof node !== 'object') return false;
      const token = node.token.toUpperCase();
      if (node.nodeType === libs.sequenceNodeType.FUNCTION) {
        if (nodes[index + 1] === '(') return token.startsWith('PLATFORM_') || !functions.hasExecutor(token);
        if (brokenReferenceToken(token)) return false;
        // The pinned lexer labels unresolved names and structured table prefixes as
        // FUNCTION until their evaluation context resolves them.
        return !knownNames.has(token) && !['TRUE', 'FALSE'].includes(token) && !tableNames.has(token.split('[')[0]);
      }
      return (
        node.nodeType === libs.sequenceNodeType.DEFINED_NAME && !brokenReferenceToken(token) && !knownNames.has(token)
      );
    });
  };
  const unsafeNames = new Set(
    definedNames
      .filter(
        (name) =>
          name.scopeInvalid ||
          metadataName(name.name) ||
          (() => {
            try {
              return financialCalls(name.formula).length > 0;
            } catch {
              return true;
            }
          })() ||
          referenceUnion(name.formula) ||
          arrayFormulaContext(name.formula).blocked ||
          quotedDateArgument(name.formula) ||
          unknown(name.formula, name.sheet)
      )
      .map((name) => name.name.toUpperCase())
  );
  const arrayNames = new Set(
    definedNames.filter((name) => arrayFormulaContext(name.formula).blocked).map((name) => name.name.toUpperCase())
  );
  const usesArrayName = (formula) =>
    parser
      .sequenceNodesBuilder('=' + String(formula).replace(/^=/, ''))
      .some((node) => typeof node === 'object' && arrayNames.has(node.token.toUpperCase()));
  for (let previous = -1; previous !== arrayNames.size; ) {
    previous = arrayNames.size;
    for (const name of definedNames) if (usesArrayName(name.formula)) arrayNames.add(name.name.toUpperCase());
  }
  const usesUnsafeName = (formula) =>
    parser
      .sequenceNodesBuilder('=' + String(formula).replace(/^=/, ''))
      .some((node) => typeof node === 'object' && unsafeNames.has(node.token.toUpperCase()));
  for (let previous = -1; previous !== unsafeNames.size; ) {
    previous = unsafeNames.size;
    for (const name of definedNames) if (usesUnsafeName(name.formula)) unsafeNames.add(name.name.toUpperCase());
  }
  for (const sheet of view.sheets)
    for (const ref in sheet.cells) {
      const cell = sheet.cells[ref];
      if (cell.formula && !cell.formulaType) {
        if (arrayFormulaContext(cell.formula).blocked || usesArrayName(cell.formula)) {
          const id = ids.get(sheet.name),
            { row, column } = position(ref);
          unsupported.set(key(id, row, column), ARRAY_CALCULATION_REASON);
          // With no qualified shape, a spill could occupy any position down/right
          // from its master. Protect reads there, including absent follower cells.
          unsafe.push({
            id,
            row,
            column,
            range: { startRow: row, endRow: 1048575, startColumn: column, endColumn: 16383 },
          });
          if (!arrayFormulaContext(cell.formula).blocked)
            await engine.api.executeCommand(
              libs.SetRangeValuesMutation.id,
              {
                unitId: engine.workbookId,
                subUnitId: id,
                cellValue: { [row]: { [column]: { f: '=PLATFORM_UNSUPPORTED_ARRAY_FORMULA()' } } },
              },
              { onlyLocal: true }
            );
        } else if (referenceUnion(cell.formula)) {
          const id = ids.get(sheet.name),
            { row, column } = position(ref);
          unsupported.set(key(id, row, column), 'Multi-area reference unions require qualified native calculation');
          unsafe.push({ id, row, column });
        } else if (unknown(cell.formula, sheet.name) || usesUnsafeName(cell.formula)) {
          const id = ids.get(sheet.name),
            { row, column } = position(ref);
          unsupported.set(key(id, row, column), 'Formula contains an unsupported function or defined name');
          unsafe.push({ id, row, column });
        }
      }
    }
  return { book, ids, unsafe, unsupported };
}
// Removing every acyclic leaf leaves precisely cycles and formulas depending on
// them. The graph comes from Univer's parser, including names and table ranges.
function circularClosure(trees) {
  const all = new Map(trees.map((tree) => [tree.treeId, tree]));
  const degree = new Map(trees.map((tree) => [tree.treeId, new Set(tree.children.filter((id) => all.has(id))).size]));
  const queue = trees.filter((tree) => degree.get(tree.treeId) === 0).map((tree) => tree.treeId);
  for (let i = 0; i < queue.length; i++)
    for (const id of all.get(queue[i]).parents) {
      if (!degree.has(id)) continue;
      degree.set(id, degree.get(id) - 1);
      if (degree.get(id) === 0) queue.push(id);
    }
  return new Set(
    trees.filter((tree) => degree.get(tree.treeId) > 0).map((tree) => key(tree.subUnitId, tree.row, tree.column))
  );
}
function unsafeClosure(trees, unsafe) {
  const all = new Map(trees.map((tree) => [tree.treeId, tree])),
    affected = new Set();
  for (const tree of trees)
    for (const source of unsafe)
      if (
        tree.rangeList.some(({ sheetId, range }) => {
          const blocked = source.range ?? {
            startRow: source.row,
            endRow: source.row,
            startColumn: source.column,
            endColumn: source.column,
          };
          return (
            sheetId === source.id &&
            blocked.startRow <= range.endRow &&
            blocked.endRow >= range.startRow &&
            blocked.startColumn <= range.endColumn &&
            blocked.endColumn >= range.startColumn
          );
        })
      ) {
        affected.add(tree.treeId);
        break;
      }
  const queue = [...affected];
  for (let i = 0; i < queue.length; i++)
    for (const id of all.get(queue[i])?.parents ?? [])
      if (!affected.has(id)) {
        affected.add(id);
        queue.push(id);
      }
  return new Set(
    trees.filter((tree) => affected.has(tree.treeId)).map((tree) => key(tree.subUnitId, tree.row, tree.column))
  );
}
function fallback(view, error, takeOwnership = false) {
  if (takeOwnership) {
    view.calculationEngine = CALCULATION_ENGINE;
    for (const sheet of view.sheets)
      for (const ref in sheet.cells) {
        const cell = sheet.cells[ref];
        if (isFormulaCell(cell)) cell.calculation = { status: 'unresolved', error, revision: view.revision };
      }
    return view;
  }
  return {
    ...view,
    calculationEngine: CALCULATION_ENGINE,
    sheets: view.sheets.map((sheet) => ({
      ...sheet,
      cells: mapCells(sheet.cells, (cell) =>
        isFormulaCell(cell)
          ? { ...cell, calculation: { status: 'unresolved', error, revision: view.revision } }
          : { ...cell }
      ),
    })),
  };
}
// Functions whose result depends on the clock, random state or dynamically chosen
// references are not qualified for reproducible command replay.
const volatileFormula = (formula) =>
  /\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT)\s*\(/i.test(
    String(formula).replace(/"(?:[^"]|"")*"/g, '')
  );
function arraySources(trees, arrayRanges) {
  const sources = [];
  // Read only known formula masters. Upstream owns this transient map; retain
  // no follower values or per-workbook closure after engine disposal.
  for (const tree of trees) {
    const range = arrayRanges?.[tree.subUnitId]?.[tree.row]?.[tree.column];
    if (range && (range.endRow > range.startRow || range.endColumn > range.startColumn)) {
      sources.push({ id: tree.subUnitId, row: tree.row, column: tree.column, range: { ...range } });
    }
  }
  return sources;
}
function arrayRanges(engine, libs) {
  return engine.univer.__getInjector().get(libs.FormulaDataModel).getArrayFormulaRange()[engine.workbookId];
}
function interpretation(
  libs,
  trees,
  unsafe,
  unsupported,
  read,
  typedErrors,
  dateMisses,
  financialMisses,
  financialAuthoredMisses,
  ranges
) {
  const shapes = arraySources(trees, ranges),
    arrays = new Set(shapes.map((source) => key(source.id, source.row, source.column)));
  const circular = circularClosure(trees),
    textErrorFormulas = new Set(),
    sources = [
      ...unsafe,
      ...dateMisses.values(),
      ...financialMisses.values(),
      ...financialAuthoredMisses.values(),
      ...shapes,
    ];
  for (const tree of trees) {
    if (!/"(?:[^"]|"")*"/.test(tree.formula)) continue;
    const result = read(tree.subUnitId, tree.row, tree.column);
    if (ERROR_VALUES.has(result?.v) && typedErrors.get(key(tree.subUnitId, tree.row, tree.column)) !== result.v) {
      textErrorFormulas.add(key(tree.subUnitId, tree.row, tree.column));
      sources.push({ id: tree.subUnitId, row: tree.row, column: tree.column });
    }
  }
  const unresolved = unsafeClosure(trees, sources);
  return (cell, id, row, column, revision) => {
    const cellKey = key(id, row, column),
      result = read(id, row, column);
    let calculation;
    if (cell.formulaType)
      calculation = { status: 'unresolved', error: `Unsupported ${cell.formulaType} formula semantics` };
    else if (unsupported.has(cellKey)) calculation = { status: 'unresolved', error: unsupported.get(cellKey) };
    else if (financialAuthoredMisses.has(cellKey))
      calculation = { status: 'unresolved', error: financialAuthoredMisses.get(cellKey).reason };
    else if (financialMisses.has(cellKey))
      calculation = { status: 'unresolved', error: financialMisses.get(cellKey).reason };
    else if (dateMisses.has(cellKey)) calculation = { status: 'unresolved', error: dateMisses.get(cellKey).reason };
    else if (arrays.has(cellKey)) calculation = { status: 'unresolved', error: ARRAY_CALCULATION_REASON };
    else if (circular.has(cellKey)) calculation = { status: 'circular', error: '#CYCLE!' };
    else if (unresolved.has(cellKey))
      calculation = {
        status: 'unresolved',
        error: 'Dependency cannot be calculated without changing its stored meaning',
      };
    else if (textErrorFormulas.has(cellKey)) {
      const literal = /^\s*=?\s*"((?:[^"]|"")*)"\s*$/.exec(cell.formula);
      calculation = literal
        ? { status: 'calculated', value: literal[1].replace(/""/g, '"') }
        : { status: 'unresolved', error: 'Engine cannot distinguish error-looking formula text from an error' };
    } else if (result?.v === undefined || result?.v === null)
      calculation = { status: 'unresolved', error: 'No calculation result' };
    else if (ERROR_VALUES.has(result.v))
      calculation = {
        status:
          result.v === '#NAME?' || result.v === '#GETTING_DATA'
            ? 'unresolved'
            : result.v === '#CYCLE!'
              ? 'circular'
              : 'error',
        error: result.v,
      };
    else if (typeof result.v === 'number' && !Number.isFinite(result.v))
      calculation = { status: 'error', error: '#NUM!' };
    else
      calculation = {
        status: 'calculated',
        value: result.t === libs.CellValueType.BOOLEAN ? Boolean(result.v) : result.v,
      };
    return { ...calculation, revision };
  };
}
async function manuallyCalculate(engine, libs, dirtyRanges, timeoutMs) {
  await engine.api.executeCommand(
    libs.SetFormulaCalculationStartMutation.id,
    { forceCalculation: dirtyRanges === undefined, dirtyRanges: dirtyRanges ?? [] },
    { onlyLocal: true }
  );
  await engine.api.getFormula().onCalculationResultApplied(timeoutMs);
}
function manualScheduler(engine, libs) {
  // Pinned upstream scheduling controller adds a 100 ms debounce. This owner
  // serializes command batches and explicitly starts calculation instead.
  engine.univer.__getInjector().get(libs.TriggerCalculationController).dispose();
}
export async function calculateWorkbookWith(
  libs,
  view,
  { timeoutMs = 10000, takeOwnership = false, inputProjection = 'static-references', onInputProjection } = {}
) {
  // The default keeps caller-owned views immutable. The package service may
  // transfer a freshly parsed, unpublished view so calculation need not clone
  // every literal cell while its engine graph is still resident.
  if (!containsFormulas(view))
    return takeOwnership
      ? Object.assign(view, { calculationEngine: CALCULATION_ENGINE })
      : { ...view, calculationEngine: CALCULATION_ENGINE };
  const engine = createEngine(libs, { manual: true });
  try {
    const { book, ids, unsafe, unsupported } = await load(engine, libs, view, {
      projectReferences: inputProjection !== 'full',
      groupNamedExpressions: true,
      onInputProjection,
    });
    manualScheduler(engine, libs);
    await manuallyCalculate(engine, libs, undefined, timeoutMs);
    const trees = await engine.api.getFormula().getAllDependencyTrees(timeoutMs);
    const sheets = new Map([...ids.values()].map((id) => [id, book.getSheetBySheetId(id)]));
    // A workbook snapshot copies all input cells. Interpretation only reads
    // formula results, including error-looking text, through bounded ranges.
    const interpret = interpretation(
      libs,
      trees,
      unsafe,
      unsupported,
      (id, row, column) => sheets.get(id)?.getRange(address(row, column)).getCellDatas()[0]?.[0],
      engine.typedErrors,
      engine.dateMisses,
      engine.financialMisses,
      engine.financialAuthoredMisses,
      arrayRanges(engine, libs)
    );
    if (takeOwnership) {
      view.calculationEngine = CALCULATION_ENGINE;
      for (const sheet of view.sheets) {
        sheet.calculationRevision = view.revision;
        for (const ref in sheet.cells) {
          const cell = sheet.cells[ref];
          if (!isFormulaCell(cell)) continue;
          const { row, column } = position(ref);
          cell.calculation = interpret(cell, ids.get(sheet.name), row, column, view.revision);
        }
      }
      return view;
    }
    return {
      ...view,
      calculationEngine: CALCULATION_ENGINE,
      sheets: view.sheets.map((sheet) => ({
        ...sheet,
        calculationRevision: view.revision,
        cells: mapCells(sheet.cells, (cell, ref) => {
          if (!isFormulaCell(cell)) return { ...cell };
          const { row, column } = position(ref);
          return { ...cell, calculation: interpret(cell, ids.get(sheet.name), row, column, view.revision) };
        }),
      })),
    };
  } catch (error) {
    return fallback(view, `Calculation unavailable: ${error.message}`, takeOwnership);
  } finally {
    engine.univer.dispose();
  }
}
/** Retains only the engine graph. The caller supplies immutable projected views;
 * literal updates return bounded result patches. Formula/structural changes must
 * dispose and create a new session, so dependencies can never silently go stale. */
export async function createCalculationSessionWith(
  libs,
  view,
  { timeoutMs = 10000, takeOwnership = false, inputProjection = 'static-references', onInputProjection } = {}
) {
  const engine = createEngine(libs, { manual: true });
  let disposed = false;
  try {
    const loaded = await load(engine, libs, view, {
        projectReferences: inputProjection !== 'full',
        groupNamedExpressions: true,
        onInputProjection,
      }),
      { book, ids, unsupported } = loaded;
    manualScheduler(engine, libs);
    await manuallyCalculate(engine, libs, undefined, timeoutMs);
    const trees = await engine.api.getFormula().getAllDependencyTrees(timeoutMs);
    const sheets = new Map([...ids].map(([name, id]) => [id, book.getSheetBySheetId(id)]));
    const read = (id, row, column) => sheets.get(id).getRange(address(row, column)).getCellDatas()[0]?.[0];
    const unsafe = new Map(loaded.unsafe.map((source) => [key(source.id, source.row, source.column), source]));
    const formulas = new Map();
    for (const sheet of view.sheets)
      for (const ref in sheet.cells) {
        const cell = sheet.cells[ref];
        if (isFormulaCell(cell)) formulas.set(`${sheet.name}!${ref}`, cell.formula ?? cell.formulaType);
      }
    const initialInterpret = interpretation(
      libs,
      trees,
      [...unsafe.values()],
      unsupported,
      read,
      engine.typedErrors,
      engine.dateMisses,
      engine.financialMisses,
      engine.financialAuthoredMisses,
      arrayRanges(engine, libs)
    );
    let initial;
    if (takeOwnership) {
      view.calculationEngine = CALCULATION_ENGINE;
      for (const sheet of view.sheets) {
        sheet.calculationRevision = view.revision;
        for (const ref in sheet.cells) {
          const cell = sheet.cells[ref];
          if (!isFormulaCell(cell)) continue;
          const { row, column } = position(ref);
          cell.calculation = initialInterpret(cell, ids.get(sheet.name), row, column, view.revision);
        }
      }
      initial = view;
    } else
      initial = {
        ...view,
        calculationEngine: CALCULATION_ENGINE,
        sheets: view.sheets.map((sheet) => ({
          ...sheet,
          calculationRevision: view.revision,
          cells: mapCells(sheet.cells, (cell, ref) => {
            const { row, column } = position(ref);
            return isFormulaCell(cell)
              ? { ...cell, calculation: initialInterpret(cell, ids.get(sheet.name), row, column, view.revision) }
              : { ...cell };
          }),
        })),
      };
    return {
      view: initial,
      async update(next, changes) {
        if (disposed) throw refusal('formula-refused', 'Calculation session is disposed');
        if (
          next.sheets.length !== ids.size ||
          next.sheets.some(
            (sheet) =>
              !ids.has(sheet.name) ||
              (sheet.date1904 !== undefined && typeof sheet.date1904 !== 'boolean') ||
              engine.dateSystems.get(ids.get(sheet.name)) !== !!sheet.date1904
          )
        )
          throw refusal('formula-unsupported', 'Incremental calculation requires unchanged valid date-system metadata');
        const dirty = [],
          priorArrays = arraySources(trees, arrayRanges(engine, libs));
        for (const change of changes) {
          const sheet = next.sheets.find((item) => item.name === change.sheet),
            cell = sheet?.cells[change.cell];
          if (!sheet || !cell || isFormulaCell(cell) || formulas.has(`${change.sheet}!${change.cell}`))
            throw refusal('formula-unsupported', 'Incremental calculation only accepts existing literal coordinates');
          const id = ids.get(sheet.name),
            { row, column } = position(change.cell),
            sourceKey = key(id, row, column);
          unsafe.delete(sourceKey);
          if (
            ((cell.dataType === 'text' || typeof cell.value === 'string') && ERROR_VALUES.has(cell.value)) ||
            (typeof cell.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue))
          )
            unsafe.set(sourceKey, { id, row, column });
          await engine.api.executeCommand(
            libs.SetRangeValuesMutation.id,
            { unitId: engine.workbookId, subUnitId: id, cellValue: { [row]: { [column]: literal(cell, libs) } } },
            { onlyLocal: true }
          );
          dirty.push({
            unitId: engine.workbookId,
            sheetId: id,
            range: { startRow: row, endRow: row, startColumn: column, endColumn: column },
          });
        }
        const affected = unsafeClosure(
          trees,
          dirty.map((item) => ({ id: item.sheetId, row: item.range.startRow, column: item.range.startColumn }))
        );
        // Financial proofs own only their own misses. A valid Financial call must
        // never erase date, volatile, source-format or complete array geometry.
        for (const sheet of next.sheets)
          for (const [ref, cell] of Object.entries(sheet.cells))
            if (cell.formula) {
              const { row, column } = position(ref),
                id = ids.get(sheet.name),
                cellKey = key(id, row, column);
              const result = checkFinancialFormula(cell.formula, sheet.cells);
              engine.financialProofs.delete(cellKey);
              engine.financialAuthoredMisses.delete(cellKey);
              if (result.calls?.length) engine.financialProofs.set(cellKey, result.calls);
              if (result.unresolved)
                engine.financialAuthoredMisses.set(cellKey, { id, row, column, reason: result.unresolved });
            }
        for (const cellKey of affected) engine.financialMisses.delete(cellKey);
        for (const cellKey of affected) engine.dateMisses.delete(cellKey);
        if (dirty.length) await manuallyCalculate(engine, libs, dirty, timeoutMs);
        const shapes = [...priorArrays, ...arraySources(trees, arrayRanges(engine, libs))];
        const consumers = unsafeClosure(trees, shapes);
        if (dirty.length && [...consumers].some((cellKey) => !affected.has(cellKey))) {
          await manuallyCalculate(
            engine,
            libs,
            shapes.map((source) => ({ unitId: engine.workbookId, sheetId: source.id, range: source.range })),
            timeoutMs
          );
        }
        for (const cellKey of consumers) affected.add(cellKey);
        // Interpretation uses the same safety graph as a complete calculation. Only
        // formula cells depending on the edited literals need new result objects.
        const interpret = interpretation(
            libs,
            trees,
            [...unsafe.values()],
            unsupported,
            read,
            engine.typedErrors,
            engine.dateMisses,
            engine.financialMisses,
            engine.financialAuthoredMisses,
            arrayRanges(engine, libs)
          ),
          patches = [];
        for (const tree of trees)
          if (affected.has(key(tree.subUnitId, tree.row, tree.column))) {
            const name = [...ids].find(([, id]) => id === tree.subUnitId)?.[0],
              ref = address(tree.row, tree.column),
              cell = next.sheets.find((sheet) => sheet.name === name)?.cells[ref];
            if (cell)
              patches.push({
                sheet: name,
                cell: ref,
                value: { ...cell, calculation: interpret(cell, tree.subUnitId, tree.row, tree.column, next.revision) },
              });
          }
        return patches;
      },
      dispose() {
        if (!disposed) {
          disposed = true;
          engine.univer.dispose();
        }
      },
    };
  } catch (error) {
    engine.univer.dispose();
    throw error;
  }
}
export function shiftFormulaWith(libs, formula, columnOffset, rowOffset) {
  const engine = createEngine(libs);
  try {
    return engine.api
      .getFormula()
      .moveFormulaRefOffset('=' + formula.replace(/^=/, ''), columnOffset, rowOffset)
      .replace(/^=/, '');
  } finally {
    engine.univer.dispose();
  }
}
/** Use the engine's reference-aware row insertion to prepare the logical view.
 * The caller still enforces the package's structural preservation policy. */
export async function transformWorkbookWith(libs, view, edits) {
  if (view.sheets.some((sheet) => Object.values(sheet.cells).some((cell) => cell.formulaType)))
    throw refusal(
      'formula-unsupported',
      'Structural editing of shared, array, or data-table formulas is not qualified'
    );
  const engine = createEngine(libs);
  try {
    const { book, ids } = await load(engine, libs, view),
      parser = engine.api.getFormula();
    let projected = structuredClone(view);
    const serialize = (reference) =>
      reference.sheetName
        ? libs.serializeRangeWithSheet(reference.sheetName, reference.range)
        : libs.serializeRange(reference.range);
    for (const edit of edits) {
      const sheet = book.getSheetBySheetId(ids.get(edit.sheet));
      if (!sheet) throw refusal('formula-not-found', `Sheet not found: ${edit.sheet}`);
      if (edit.operation !== 'insertRow')
        throw refusal('formula-unsupported', `Unsupported structural operation: ${edit.operation}`);
      const count = edit.count ?? 1;
      if (
        !Number.isSafeInteger(edit.index) ||
        edit.index < 0 ||
        !Number.isSafeInteger(count) ||
        count < 1 ||
        edit.index + count > 1048576
      )
        throw refusal('internal-invalid-field', 'Invalid inserted row range');
      const original = projected.sheets.find((item) => item.name === edit.sheet),
        totalPatches = [],
        namePatches = [];
      if ((original.rows ?? 0) + count > 1048576)
        throw refusal('formula-limit', 'Inserted rows exceed the XLSX row boundary');
      for (const ref in original.cells)
        if (position(ref).row >= edit.index && position(ref).row + count >= 1048576)
          throw refusal('formula-conflict', 'Inserted rows would move an occupied cell outside XLSX');
      const boundary = (token, ownerSheet = edit.sheet) => {
        if (token?.nodeType !== libs.sequenceNodeType.REFERENCE) return null;
        const ref = libs.deserializeRangeWithSheet(token.token);
        if (ref.unitId || (ref.sheetName || ownerSheet) !== edit.sheet || ref.range.endRow !== edit.index - 1)
          return null;
        return { ...ref, range: { ...ref.range, endRow: ref.range.endRow + count } };
      };
      for (const [ref, cell] of Object.entries(original.cells))
        if (cell.formula && position(ref).row === edit.index) {
          const nodes = parser.sequenceNodesBuilder('=' + cell.formula.replace(/^=/, '')),
            affected = nodes.map((token) => boundary(token)).filter(Boolean);
          if (!affected.length) continue;
          if (
            nodes.length !== 4 ||
            nodes[0]?.nodeType !== libs.sequenceNodeType.FUNCTION ||
            !['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT'].includes(nodes[0].token.toUpperCase()) ||
            nodes[1] !== '(' ||
            nodes[3] !== ')' ||
            !boundary(nodes[2])
          )
            throw refusal(
              'formula-unsupported',
              `Complex total at insertion boundary is not qualified: ${edit.sheet}!${ref}`
            );
          totalPatches.push({
            column: position(ref).column,
            formula: `${nodes[0].token}(${serialize(boundary(nodes[2]))})`,
          });
        }
      for (const originalName of projected.namedRanges ?? []) {
        const owner = originalName.sheet ? book.getSheetBySheetId(ids.get(originalName.sheet)) : book;
        const name = owner.getDefinedNames().find((item) => item.getName() === originalName.name);
        if (!name) throw refusal('formula-refused', 'Defined name disappeared');
        const nodes = parser.sequenceNodesBuilder('=' + name.getFormulaOrRefString().replace(/^=/, ''));
        // A scalar name is a fixed cell identity, not a growing budget region.
        // Unqualified references inherit the name's scope, never the edited sheet.
        const affected = nodes
          .filter((token) => typeof token === 'object' && token.token.includes(':'))
          .map((token) => boundary(token, originalName.sheet ?? null))
          .filter(Boolean);
        if (affected.length) {
          if (nodes.length !== 1)
            throw refusal('formula-unsupported', 'Complex named formula at insertion boundary is not qualified');
          namePatches.push({ name, formula: serialize(affected[0]) });
        }
      }
      const tablePatches = book
        .getTableList()
        .filter((table) => table.subUnitId === ids.get(edit.sheet) && table.range.endRow === edit.index - 1)
        .map((table) => ({ id: table.id, range: { ...table.range, endRow: table.range.endRow + count } }));
      sheet.insertRowsBefore(edit.index, count);
      for (const patch of tablePatches)
        if (!(await sheet.setTableRange(patch.id, patch.range)))
          throw refusal('formula-refused', 'Table range extension failed');
      for (const patch of namePatches) patch.name.setRef(patch.formula);
      const inherited = {};
      if (edit.index > 0)
        for (const [ref, cell] of Object.entries(original.cells))
          if (position(ref).row === edit.index - 1) {
            const column = position(ref).column,
              sourceFormula = sheet.getRange(address(edit.index - 1, column)).getCellDatas()[0][0]?.f;
            const style = {};
            for (const field of ['styleId', 'format', 'bold', 'italic'])
              if (cell[field] !== undefined) style[field] = cell[field];
            for (let offset = 0; offset < count; offset++) {
              const target = address(edit.index + offset, column),
                formula = sourceFormula
                  ? parser.moveFormulaRefOffset(sourceFormula, 0, offset + 1).replace(/^=/, '')
                  : undefined;
              inherited[target] = {
                ...style,
                value: null,
                dataType: 'blank',
                rawValue: null,
                cachedValue: null,
                ...(formula ? { formula } : {}),
              };
              sheet.getRange(target).setValue(formula ? { f: '=' + formula } : { v: null });
            }
          }
      for (const patch of totalPatches)
        sheet.getRange(address(edit.index + count, patch.column)).setValue({ f: '=' + patch.formula });
      const snapshot = book.save();
      const sheets = projected.sheets.map((previous) => {
        const current = snapshot.sheets[ids.get(previous.name)],
          cells = {};
        for (const [ref, cell] of Object.entries(previous.cells)) {
          let { row, column } = position(ref);
          if (previous.name === edit.sheet && row >= edit.index) row += count;
          const updated = current.cellData[row]?.[column];
          if (cell.formula && !updated?.f)
            throw refusal('formula-refused', `Formula disappeared during structural edit: ${previous.name}!${ref}`);
          cells[address(row, column)] = {
            ...cell,
            ...(cell.formula ? { formula: updated.f.replace(/^=/, ''), calculation: undefined } : {}),
          };
        }
        if (previous.name === edit.sheet) Object.assign(cells, inherited);
        return { ...previous, rows: current.rowCount, cells };
      });
      const namedRanges = (projected.namedRanges ?? []).map((previous) => {
        const owner = previous.sheet ? book.getSheetBySheetId(ids.get(previous.sheet)) : book;
        const name = owner.getDefinedNames().find((item) => item.getName() === previous.name);
        if (!name) throw refusal('formula-refused', 'Defined name disappeared');
        return { ...previous, formula: name.getFormulaOrRefString().replace(/^=/, '') };
      });
      const tables = (projected.tables ?? []).map((previous) => {
        const table = book.getTableInfoByName(previous.name);
        if (!table) throw refusal('formula-refused', 'Table disappeared');
        const { startRow, startColumn, endRow, endColumn } = table.range;
        return { ...previous, ref: `${address(startRow, startColumn)}:${address(endRow, endColumn)}` };
      });
      projected = { ...projected, sheets, namedRanges, tables };
    }
    return projected;
  } finally {
    engine.univer.dispose();
  }
}

function address(row, column) {
  let letters = '';
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters + (row + 1);
}
export function shiftFormulasWith(libs, requests) {
  const engine = createEngine(libs);
  try {
    return requests.map(({ formula, columnOffset, rowOffset }) =>
      engine.api
        .getFormula()
        .moveFormulaRefOffset('=' + formula.replace(/^=/, ''), columnOffset, rowOffset)
        .replace(/^=/, '')
    );
  } finally {
    engine.univer.dispose();
  }
}
export function formulaReferencesWith(libs, formula) {
  const engine = createEngine(libs);
  try {
    const nodes = engine.api.getFormula().sequenceNodesBuilder('=' + formula.replace(/^=/, '')),
      references = [];
    let unsupported = false;
    for (let index = 0; index < nodes.length; index++) {
      const node = nodes[index];
      if (typeof node === 'string') continue;
      if (node.nodeType === libs.sequenceNodeType.REFERENCE) {
        const reference = libs.deserializeRangeWithSheet(node.token);
        if (
          reference.unitId ||
          reference.sheetName?.includes(':') ||
          !Number.isFinite(reference.range.startRow) ||
          reference.range.startRow < 0
        )
          unsupported = true;
        references.push(reference);
      } else if (node.nodeType === libs.sequenceNodeType.FUNCTION) {
        if (nodes[index + 1] !== '(' || ['INDIRECT', 'OFFSET'].includes(node.token.toUpperCase())) unsupported = true;
      } else if ([libs.sequenceNodeType.DEFINED_NAME, libs.sequenceNodeType.TABLE].includes(node.nodeType))
        unsupported = true;
    }
    return { references, unsupported };
  } finally {
    engine.univer.dispose();
  }
}
