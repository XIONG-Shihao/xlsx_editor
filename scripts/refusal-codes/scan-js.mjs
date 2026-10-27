// Finds the places in JavaScript/TypeScript that create an error from a message:
// `Error('…')`, `new Error('…')`, `throw Error(\`…\`)`. TypeScript's parser is used so
// strings, templates and comments cannot fool it.
import ts from 'typescript';

/** Text of the message with every computed part replaced by `{}`. */
function staticText(node, source) {
  if (!node) return '<no message>';
  if (ts.isStringLiteralLike(node)) return node.text;
  if (ts.isTemplateExpression(node))
    return node.head.text + node.templateSpans.map((span) => '{}' + span.literal.text).join('');
  if (ts.isParenthesizedExpression(node)) return staticText(node.expression, source);
  if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken) {
    const left = staticText(node.left, source);
    const right = staticText(node.right, source);
    return left + right;
  }
  if (ts.isConditionalExpression(node)) return staticText(node.whenTrue, source);
  return '{}';
}

function isLiteralMessage(node) {
  if (ts.isStringLiteralLike(node) || ts.isTemplateExpression(node)) return true;
  if (ts.isParenthesizedExpression(node)) return isLiteralMessage(node.expression);
  if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken)
    return isLiteralMessage(node.left) || isLiteralMessage(node.right);
  if (ts.isConditionalExpression(node)) return isLiteralMessage(node.whenTrue) && isLiteralMessage(node.whenFalse);
  return false;
}

/**
 * Every `Error(...)` creation in the text.
 * Each site: { file, line, start, end, argStart, message, dynamic, legacyCode, extraArgs }
 *   start..argStart is the text `Error(` or `new Error(` that a coded site replaces with `refusal('code', `
 */
export function scanJs(file, text) {
  const kind = /\.tsx$/.test(file)
    ? ts.ScriptKind.TSX
    : /\.ts$|\.mts$/.test(file)
      ? ts.ScriptKind.TS
      : ts.ScriptKind.JS;
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, kind);
  const sites = [];
  const visit = (node) => {
    if (
      (ts.isCallExpression(node) || ts.isNewExpression(node)) &&
      ts.isIdentifier(node.expression) &&
      node.expression.text === 'Error'
    ) {
      const args = node.arguments ?? [];
      const parent = node.parent;
      // `Object.assign(Error('…'), { code: … })` already carries a code of its own.
      const legacyCode =
        ts.isCallExpression(parent) &&
        parent.arguments[0] === node &&
        parent.arguments.slice(1).some((arg) => /\bcode\b/.test(arg.getText(source)));
      const first = args[0];
      const open = text.indexOf('(', node.expression.end);
      sites.push({
        file,
        line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
        start: node.getStart(source),
        argStart: open + 1,
        end: node.end,
        message: staticText(first, source),
        dynamic: !first || !isLiteralMessage(first),
        argText: first ? first.getText(source).replace(/\s+/g, ' ') : '',
        legacyCode,
        extraArgs: args.length > 1,
        thrown: ts.isThrowStatement(parent),
      });
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return sites;
}

/** Every `refusal('code', message)` in the text, as { code, message, line }. */
export function scanJsCoded(file, text) {
  const kind = /\.tsx$/.test(file)
    ? ts.ScriptKind.TSX
    : /\.ts$|\.mts$/.test(file)
      ? ts.ScriptKind.TS
      : ts.ScriptKind.JS;
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, kind);
  const found = [];
  const visit = (node) => {
    if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === 'refusal') {
      const [code, message] = node.arguments;
      if (code && ts.isStringLiteralLike(code))
        found.push({
          file,
          code: code.text,
          message: staticText(message, source),
          dynamic: !message || !isLiteralMessage(message),
          line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
        });
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
}
