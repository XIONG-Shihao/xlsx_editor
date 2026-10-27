/** Explicit pure numeric VBA subset. AST interpretation only; no eval or host bindings. */
import { refusal } from './refusal-codes.mjs';

export const PURE_VBA_LIMITS = Object.freeze({
  sourceBytes: 128 * 1024,
  functions: 32,
  parameters: 16,
  statements: 256,
  tokens: 2048,
  nodes: 1024,
  depth: 32,
  steps: 8192,
});
const fail = (message) => {
  throw refusal('workbook-refused', 'Pure VBA subset: ' + message);
};
const identifier = (value) => {
  if (
    !/^[A-Za-z][A-Za-z\d_]{0,63}$/.test(value) ||
    /^(?:application|thisworkbook|shell|createobject|declare|eval|execute|environ|call|sub|function|end|if|then|for|while|do|loop)$/i.test(
      value
    )
  )
    fail('unsupported identifier');
  return value.toLowerCase();
};
function expression(source, variables, double) {
  const tokens = [];
  let at = 0;
  while (at < source.length) {
    if (/\s/.test(source[at])) {
      at++;
      continue;
    }
    const rest = source.slice(at),
      number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?#?/.exec(rest),
      name = /^[A-Za-z][A-Za-z\d_]*/.exec(rest);
    if (number) {
      const value = Number(number[0].replace(/#$/, ''));
      if (!Number.isFinite(value)) fail('nonfinite literal');
      if (double && !number[0].endsWith('#')) fail('Double arithmetic literals require the # suffix');
      if (!double && (!/^\d+$/.test(number[0]) || value > 32767))
        fail('untyped constant must be a small Integer literal');
      tokens.push({ kind: 'number', value });
      at += number[0].length;
    } else if (name) {
      const value = identifier(name[0]);
      if (!variables.has(value)) fail('undeclared or external value');
      tokens.push({ kind: 'variable', value });
      at += name[0].length;
    } else if ('+-*/()'.includes(rest[0])) {
      tokens.push({ kind: rest[0] });
      at++;
    } else fail('operator, call or host syntax is unsupported');
    if (tokens.length > PURE_VBA_LIMITS.tokens) fail('expression token budget exceeded');
  }
  let cursor = 0,
    nodes = 0;
  const node = (value) => {
    if (++nodes > PURE_VBA_LIMITS.nodes) fail('expression node budget exceeded');
    return value;
  };
  function primary(depth) {
    if (depth > PURE_VBA_LIMITS.depth) fail('expression depth exceeded');
    const token = tokens[cursor++];
    if (!token) fail('incomplete expression');
    if (token.kind === 'number' || token.kind === 'variable') return node(token);
    if (token.kind === '+' || token.kind === '-')
      return node({ kind: 'unary', operator: token.kind, value: primary(depth + 1) });
    if (token.kind === '(') {
      const value = sum(depth + 1);
      if (tokens[cursor++]?.kind !== ')') fail('unclosed expression');
      return value;
    }
    fail('unsupported expression');
  }
  function product(depth) {
    let value = primary(depth);
    while (['*', '/'].includes(tokens[cursor]?.kind)) {
      const operator = tokens[cursor++].kind;
      value = node({ kind: 'binary', operator, left: value, right: primary(depth) });
    }
    return value;
  }
  function sum(depth) {
    let value = product(depth);
    while (['+', '-'].includes(tokens[cursor]?.kind)) {
      const operator = tokens[cursor++].kind;
      value = node({ kind: 'binary', operator, left: value, right: product(depth) });
    }
    return value;
  }
  const tree = sum(0);
  if (cursor !== tokens.length) fail('unexpected expression suffix');
  return tree;
}
function constant(tree) {
  return tree.kind === 'number' || (tree.kind === 'unary' && constant(tree.value));
}
export function compilePureVba(source) {
  if (
    typeof source !== 'string' ||
    new TextEncoder().encode(source).length > PURE_VBA_LIMITS.sourceBytes ||
    source.includes('\0')
  )
    fail('source budget or encoding refused');
  const functions = new Map();
  let current = null;
  for (const raw of source.replaceAll('\r\n', '\n').replaceAll('\r', '\n').split('\n')) {
    let line = raw.trim();
    if (!line || line.startsWith("'") || /^Rem(?:\s|$)/i.test(line)) continue;
    if (!current && /^Attribute VB_[A-Za-z\d_]+ = (?:"(?:[^"\n]|"")*"|True|False)$/i.test(line)) continue;
    line = line.split("'")[0].trim();
    if (!line) continue;
    if (!current && /^Option Explicit$/i.test(line)) continue;
    const header = /^(?:Public\s+)?Function\s+([A-Za-z][A-Za-z\d_]*)\s*\(([^()]*)\)\s*(?:As\s+(Double))?$/i.exec(line);
    if (header) {
      if (current) fail('nested function');
      const name = identifier(header[1]);
      if (functions.has(name) || functions.size >= PURE_VBA_LIMITS.functions)
        fail('duplicate function or function budget');
      const parameters = [];
      if (header[2].trim())
        for (const input of header[2].split(',')) {
          const parameter = /^\s*ByVal\s+([A-Za-z][A-Za-z\d_]*)\s+As\s+Double\s*$/i.exec(input);
          if (!parameter || !header[3]) fail('parameters require ByVal Double and a Double return');
          const parameterName = identifier(parameter[1]);
          if (
            parameterName === name ||
            parameters.includes(parameterName) ||
            parameters.length >= PURE_VBA_LIMITS.parameters
          )
            fail('parameter identity or budget refused');
          parameters.push(parameterName);
        }
      current = {
        name,
        displayName: header[1],
        parameters,
        double: !!header[3],
        variables: new Set([name, ...parameters]),
        locals: [],
        statements: [],
        returned: false,
      };
      continue;
    }
    if (/^End Function$/i.test(line)) {
      if (!current || !current.returned) fail('function has no qualified result');
      functions.set(current.name, current);
      current = null;
      continue;
    }
    if (!current) fail('module globals, Sub, events and declarations are unsupported');
    const local = /^Dim\s+([A-Za-z][A-Za-z\d_]*)\s+As\s+Double$/i.exec(line);
    if (local) {
      if (!current.double) fail('untyped functions only support a numeric constant return');
      const name = identifier(local[1]);
      if (current.variables.has(name)) fail('duplicate local variable');
      current.variables.add(name);
      current.locals.push(name);
      continue;
    }
    const assignment = /^([A-Za-z][A-Za-z\d_]*)\s*=\s*(.+)$/.exec(line);
    if (!assignment) fail('only local numeric assignments are supported');
    const name = identifier(assignment[1]);
    if (!current.variables.has(name) || current.parameters.includes(name))
      fail('assignment cannot affect parameters or host state');
    const tree = expression(assignment[2], current.variables, current.double);
    if (!current.double && (!constant(tree) || name !== current.name || current.statements.length))
      fail('untyped function only supports one numeric constant return');
    if (current.statements.length >= PURE_VBA_LIMITS.statements) fail('statement budget exceeded');
    current.statements.push({ name, tree });
    if (name === current.name) current.returned = true;
  }
  if (current || !functions.size) fail('incomplete or absent pure numeric functions');
  return functions;
}
export function pureVbaFunctions(source) {
  return [...compilePureVba(source).values()].map((fn) => ({
    name: fn.displayName,
    parameters: fn.parameters,
    returnType: fn.double ? 'Double' : 'numeric constant',
  }));
}
export function evaluatePureVba(source, functionName, args) {
  const functions = compilePureVba(source),
    fn = functions.get(identifier(functionName));
  if (!fn) fail('function is not declared');
  if (
    !Array.isArray(args) ||
    args.length !== fn.parameters.length ||
    args.some((value) => typeof value !== 'number' || !Number.isFinite(value))
  )
    fail('exact finite numeric arguments required');
  const variables = new Map([...fn.variables].map((name) => [name, 0]));
  fn.parameters.forEach((name, index) => variables.set(name, args[index]));
  let steps = 0;
  function evaluate(tree, depth = 0) {
    if (++steps > PURE_VBA_LIMITS.steps || depth > PURE_VBA_LIMITS.depth) fail('evaluation work budget exceeded');
    if (tree.kind === 'number') return tree.value;
    if (tree.kind === 'variable') return variables.get(tree.value);
    if (tree.kind === 'unary') {
      const value = evaluate(tree.value, depth + 1);
      return tree.operator === '-' ? -value : value;
    }
    const a = evaluate(tree.left, depth + 1),
      b = evaluate(tree.right, depth + 1);
    if (tree.operator === '/' && b === 0) fail('division by zero');
    const result =
      tree.operator === '+' ? a + b : tree.operator === '-' ? a - b : tree.operator === '*' ? a * b : a / b;
    if (!Number.isFinite(result)) fail('numeric overflow');
    return result;
  }
  for (const statement of fn.statements) variables.set(statement.name, evaluate(statement.tree));
  const value = variables.get(fn.name);
  if (!Number.isFinite(value)) fail('nonfinite result');
  return {
    functionName: fn.displayName,
    value,
    steps,
    policy: 'Explicit pure numeric subset; original workbook cells, macros and host state unchanged',
  };
}
