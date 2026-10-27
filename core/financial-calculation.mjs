/** PRIVATE arithmetic-operand candidate: bounded Financial mathematics, not universal Excel bits. */
import { refusal } from './refusal-codes.mjs';

export const FINANCIAL_FUNCTIONS = new Set(['PMT', 'PV', 'FV']);
export const FINANCIAL_REASON = 'Financial formula is outside the bounded numeric scalar profile';
const MAX = Number.MAX_SAFE_INTEGER;
const meaning = (value) => {
  const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(String(value));
  if (!m) return null;
  let d = (m[2] + (m[3] ?? '')).replace(/^0+/, '');
  let e = Number(m[4] ?? 0) - (m[3]?.length ?? 0);
  if (!d) return '0';
  while (d.endsWith('0')) {
    d = d.slice(0, -1);
    e++;
  }
  return `${m[1] === '-' ? '-' : ''}${d}e${e}`;
};
const miss = (reason) => {
  throw refusal('formula-unsupported', FINANCIAL_REASON + ': ' + reason);
};
function numeric(value) {
  if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > MAX || Object.is(value, -0))
    miss('nonfinite/unsafe/negative-zero operand');
  return value;
}
function metrics(arg, depth = 1) {
  if (depth > 8) miss('operand depth');
  if (arg.kind === 'binary') {
    const a = metrics(arg.left, depth + 1),
      b = metrics(arg.right, depth + 1);
    return { nodes: a.nodes + b.nodes + 1, refs: [...a.refs, ...b.refs] };
  }
  if (arg.kind === 'unary') {
    const child = metrics(arg.value, depth + 1);
    return { nodes: child.nodes + 1, refs: child.refs };
  }
  return { nodes: 1, refs: arg.kind === 'reference' ? [arg.reference] : [] };
}
export function financialArgument(text) {
  if (typeof text !== 'string' || new TextEncoder().encode(text).length > 1024) miss('operand byte bound');
  let at = 0;
  const space = () => {
    while (/[ \t\r\n\f\v]/.test(text[at] ?? '') && at < text.length) at++;
  };
  function unary(depth) {
    if (depth > 8) miss('operand depth');
    space();
    const c = text[at];
    if (c === '+' || c === '-') {
      at++;
      return { kind: 'unary', operator: c, value: unary(depth + 1) };
    }
    if (c === '(') {
      at++;
      const expr = additive(depth + 1);
      space();
      if (text[at++] !== ')') miss('unmatched parentheses');
      return expr;
    }
    const rest = text.slice(at),
      number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
    if (number) {
      at += number[0].length;
      const value = numeric(Number(number[0]));
      if (meaning(number[0]) !== meaning(value)) miss('numeric literal precision');
      return { kind: 'number', value };
    }
    const ref = /^\$?([A-Z]{1,3})\$?([1-9]\d*)/.exec(rest);
    if (ref) {
      let column = 0;
      for (const c of ref[1]) column = column * 26 + c.charCodeAt(0) - 64;
      if (column > 16384 || Number(ref[2]) > 1048576) miss('reference bounds');
      at += ref[0].length;
      return { kind: 'reference', reference: ref[1] + ref[2] };
    }
    miss('closed arithmetic operand grammar');
  }
  function multiplicative(depth) {
    let left = unary(depth);
    while (true) {
      space();
      const operator = text[at];
      if (operator !== '*' && operator !== '/') return left;
      at++;
      left = { kind: 'binary', operator, left, right: unary(depth) };
    }
  }
  function additive(depth) {
    let left = multiplicative(depth);
    while (true) {
      space();
      const operator = text[at];
      if (operator !== '+' && operator !== '-') return left;
      at++;
      left = { kind: 'binary', operator, left, right: multiplicative(depth) };
    }
  }
  const arg = additive(1);
  space();
  if (at !== text.length) miss('closed arithmetic operand grammar');
  const proof = metrics(arg);
  if (proof.nodes > 32) miss('operand node bound');
  return arg;
}
/** Authored call-site identity precedes fallback parsing, name expansion and coercion. */
export function financialCalls(formula) {
  const text = String(formula),
    calls = [];
  let quote = '',
    square = 0;
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
      square++;
      continue;
    }
    if (c === ']') {
      square--;
      continue;
    }
    if (square) continue;
    if (!/[A-Za-z_]/.test(c)) continue;
    const startName = i;
    let end = i + 1;
    while (/[A-Za-z0-9_.]/.test(text[end] ?? '')) end++;
    const name = text.slice(i, end).toUpperCase();
    let open = end;
    while (/[ \t\r\n\f\v]/.test(text[open] ?? '') && open < text.length) open++;
    i = end - 1;
    if ((!FINANCIAL_FUNCTIONS.has(name) && !/^_XLFN\.(PMT|PV|FV)$/.test(name)) || text[open] !== '(') continue;
    if (!FINANCIAL_FUNCTIONS.has(name)) miss('unqualified alias');
    const tokens = [];
    let start = open + 1,
      depth = 1,
      q = '',
      brace = 0,
      bracket = 0,
      close;
    for (let at = start; at < text.length; at++) {
      const ch = text[at];
      if (q) {
        if (ch === q) {
          if (text[at + 1] === q) at++;
          else q = '';
        }
        continue;
      }
      if (ch === '"' || ch === "'") {
        q = ch;
        continue;
      }
      if (ch === '[') bracket++;
      if (ch === ']') bracket--;
      if (ch === '{') brace++;
      if (ch === '}') brace--;
      if (ch === '(') depth++;
      if (ch === ')') {
        depth--;
        if (!depth) {
          tokens.push(text.slice(start, at));
          close = at;
          break;
        }
      }
      if (ch === ',' && depth === 1 && !bracket && !brace) {
        tokens.push(text.slice(start, at));
        start = at + 1;
      }
    }
    if (close === undefined || tokens.length < 3 || tokens.length > 5) miss('argument count');
    const args = tokens.map(financialArgument),
      shape = args.map((arg) => metrics(arg));
    const references = shape.flatMap((s) => s.refs);
    if (shape.reduce((n, s) => n + s.nodes, 0) > 128 || references.length > 32) miss('call node/reference bound');
    calls.push({
      name,
      args,
      index: calls.length,
      start: startName,
      end: close + 1,
      references: [...new Set(references)],
    });
    i = close;
  }
  return calls;
}
/** Closed authored syntax only. This does not inspect a workbook or prove numeric reference values. */
export function financialExpressionSyntax(raw) {
  const ast = financialArgument(raw),
    shape = metrics(ast);
  return {
    text: raw.replace(/^[ \t\r\n\f\v]+|[ \t\r\n\f\v]+$/g, ''),
    ast,
    nodes: shape.nodes,
    references: [...new Set(shape.refs)],
    requiresOwnerValidation: true,
  };
}
export function rewriteFinancialFormula(formula, calls) {
  let text = formula;
  for (const call of [...calls].reverse()) {
    const token = `PLATFORM_FINANCIAL_CALL(${[call.index, ...call.references].join(',')})`;
    text = text.slice(0, call.start) + token + text.slice(call.end);
  }
  return text;
}
export function calculateFinancialScalar(name, input) {
  if (!FINANCIAL_FUNCTIONS.has(name) || input.length < 3 || input.length > 5)
    return { unresolved: FINANCIAL_REASON + ': argument count' };
  const args = [...input];
  while (args.length < 5) args.push(0);
  if (
    args.some(
      (v, i) =>
        typeof v !== 'number' ||
        !Number.isFinite(v) ||
        Object.is(v, -0) ||
        (i === 0 && (v < 0 || v > 1)) ||
        (i === 1 && (!Number.isInteger(v) || v < 1 || v > 1200)) ||
        ((i === 2 || i === 3) && Math.abs(v) > 1e12) ||
        (i === 4 && v !== 0 && v !== 1)
    )
  )
    return { unresolved: FINANCIAL_REASON + ': input domain/type' };
  const [rate, nper, flow, balance, type] = args;
  const exponent = nper * Math.log1p(rate),
    growth = rate === 0 ? 1 : rate === 1 ? 2 ** nper : Math.exp(exponent),
    annuity = rate === 0 ? nper : rate === 1 ? growth - 1 : Math.expm1(exponent) / rate,
    due = 1 + rate * type;
  if (!Number.isFinite(growth) || !Number.isFinite(annuity))
    return { unresolved: FINANCIAL_REASON + ': intermediate overflow' };
  const result =
    name === 'PMT'
      ? -(flow * growth + balance) / (due * annuity)
      : name === 'PV'
        ? -(flow * due * annuity + balance) / growth
        : -(balance * growth + flow * due * annuity);
  if (!Number.isFinite(result) || Math.abs(result) > MAX)
    return { unresolved: FINANCIAL_REASON + ': result magnitude' };
  return { value: result === 0 ? 0 : result };
}
function inputValue(arg, cells) {
  let result;
  if (arg.kind === 'number') result = arg.value;
  else if (arg.kind === 'reference') {
    const cell = cells[arg.reference];
    if (
      !cell ||
      cell.formula ||
      cell.formulaType ||
      cell.spill ||
      cell.scenarioResult ||
      cell.financialInputBlocked ||
      cell.dataType !== 'number' ||
      typeof cell.value !== 'number' ||
      Object.is(Number(cell.rawValue), -0) ||
      meaning(cell.rawValue ?? cell.value) !== meaning(cell.value)
    )
      miss('reference needs an exact ordinary numeric literal');
    result = cell.value;
  } else if (arg.kind === 'unary') {
    const n = inputValue(arg.value, cells);
    result = arg.operator === '-' ? -n : n;
  } else {
    const a = inputValue(arg.left, cells),
      b = inputValue(arg.right, cells);
    if (arg.operator === '/' && b === 0) miss('zero-divide operand');
    result = arg.operator === '+' ? a + b : arg.operator === '-' ? a - b : arg.operator === '*' ? a * b : a / b;
    if ((arg.operator === '*' || arg.operator === '/') && a !== 0 && b !== 0 && result === 0) miss('underflow operand');
  }
  return numeric(result);
}
export function checkFinancialFormula(formula, cells) {
  try {
    const calls = financialCalls(formula);
    for (const call of calls) {
      const values = call.args.map((arg) => inputValue(arg, cells));
      const result = calculateFinancialScalar(call.name, values);
      if (result.unresolved) return result;
      call.values = values;
    }
    return { calls };
  } catch (error) {
    return { unresolved: String(error.message ?? error) };
  }
}
