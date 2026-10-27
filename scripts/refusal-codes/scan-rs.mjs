// Finds the places in Rust that build an error sentence for the caller:
//   Err("…".into())   Err(format!("…", …))   .ok_or("…")   .ok_or_else(|| format!("…"))
//   .map_err(|_| "…".to_string())
// Errors forwarded from a library (`map_err(err)`, `Err(error)`) carry no sentence of ours and are not sites.
// A small tokenizer is enough here: it only has to know strings, comments and brackets.

export function tokenize(text) {
  const tokens = [];
  let i = 0;
  const n = text.length;
  const isIdStart = (c) => /[A-Za-z_]/.test(c);
  const isId = (c) => /[A-Za-z0-9_]/.test(c);
  while (i < n) {
    const c = text[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === '/' && text[i + 1] === '/') {
      while (i < n && text[i] !== '\n') i++;
      continue;
    }
    if (c === '/' && text[i + 1] === '*') {
      let depth = 1;
      i += 2;
      while (i < n && depth) {
        if (text[i] === '/' && text[i + 1] === '*') (depth++, (i += 2));
        else if (text[i] === '*' && text[i + 1] === '/') (depth--, (i += 2));
        else i++;
      }
      continue;
    }
    // raw strings r"…", r#"…"#, br#"…"#
    const raw = /^b?r(#*)"/.exec(text.slice(i, i + 12));
    if (raw && (i === 0 || !isId(text[i - 1]))) {
      const hashes = raw[1];
      const close = '"' + hashes;
      const bodyStart = i + raw[0].length;
      const end = text.indexOf(close, bodyStart);
      tokens.push({ t: 'str', start: i, end: end + close.length, value: text.slice(bodyStart, end), raw: true });
      i = end + close.length;
      continue;
    }
    if (c === '"' || (c === 'b' && text[i + 1] === '"' && (i === 0 || !isId(text[i - 1])))) {
      const start = i;
      i += c === 'b' ? 2 : 1;
      let value = '';
      while (i < n && text[i] !== '"') {
        if (text[i] === '\\') {
          const e = text[i + 1];
          if (e === 'n') value += '\n';
          else if (e === 't') value += '\t';
          else if (e === 'u') {
            const close = text.indexOf('}', i);
            value += String.fromCodePoint(parseInt(text.slice(i + 3, close), 16));
            i = close - 1;
            i += 1;
            continue;
          } else if (e === '\n') {
            i += 2;
            while (/\s/.test(text[i])) i++;
            continue;
          } else value += e;
          i += 2;
        } else value += text[i++];
      }
      i++;
      tokens.push({ t: 'str', start, end: i, value, raw: false });
      continue;
    }
    if (c === "'") {
      // a char literal ('a', '\n', '\u{1F}') or a lifetime ('a)
      const m = /^'(?:\\u\{[0-9a-fA-F]+\}|\\x[0-9a-fA-F]{2}|\\.|[^\\'])'/.exec(text.slice(i, i + 14));
      if (m) {
        tokens.push({ t: 'char', start: i, end: i + m[0].length });
        i += m[0].length;
        continue;
      }
      i++;
      while (i < n && isId(text[i])) i++;
      continue;
    }
    if (isIdStart(c)) {
      const start = i;
      while (i < n && isId(text[i])) i++;
      tokens.push({ t: 'id', start, end: i, value: text.slice(start, i) });
      continue;
    }
    if (/[0-9]/.test(c)) {
      const start = i;
      while (i < n && /[0-9A-Za-z_.]/.test(text[i]) && !(text[i] === '.' && !/[0-9]/.test(text[i + 1] ?? ''))) i++;
      tokens.push({ t: 'num', start, end: i });
      continue;
    }
    if (c === ':' && text[i + 1] === ':') {
      tokens.push({ t: 'p', start: i, end: i + 2, value: '::' });
      i += 2;
      continue;
    }
    if (c === '|' && text[i + 1] === '|') {
      tokens.push({ t: 'p', start: i, end: i + 2, value: '||' });
      i += 2;
      continue;
    }
    if (c === '=' && text[i + 1] === '>') {
      tokens.push({ t: 'p', start: i, end: i + 2, value: '=>' });
      i += 2;
      continue;
    }
    tokens.push({ t: 'p', start: i, end: i + 1, value: c });
    i++;
  }
  return tokens;
}

const OPEN = new Set(['(', '[', '{']);
const CLOSE = new Set([')', ']', '}']);

/** Index of the token that closes the bracket opened at `from`. */
function matching(tokens, from) {
  let depth = 0;
  for (let k = from; k < tokens.length; k++) {
    const v = tokens[k].t === 'p' ? tokens[k].value : '';
    if (OPEN.has(v)) depth++;
    else if (CLOSE.has(v) && --depth === 0) return k;
  }
  return -1;
}

const is = (token, value) => token && token.t === 'p' && token.value === value;
const isId = (token, value) => token && token.t === 'id' && (value === undefined || token.value === value);
const MAKERS = new Set(['into', 'to_string', 'to_owned']);

/**
 * Looks at tokens[from..to) as one expression. Returns
 *   { kind: 'literal', str, start, end }   "text" / "text".into() / String::from("text")
 *   { kind: 'format', str, start, end }     format!("text", …)
 *   { kind: 'coded' }                       already crate::coded(…)
 *   { kind: 'other' }                       anything else (a variable, a helper call, a block)
 */
function classify(tokens, from, to) {
  if (is(tokens[to - 1], ',') && to - 1 > from) to -= 1; // rustfmt leaves a trailing comma on split calls
  const first = tokens[from];
  if (!first || from >= to) return { kind: 'other' };
  const start = first.start;
  const end = tokens[to - 1].end;
  if (first.t === 'str') {
    if (to - from === 1) return { kind: 'literal', str: first, start, end, bare: true };
    if (
      is(tokens[from + 1], '.') &&
      isId(tokens[from + 2]) &&
      MAKERS.has(tokens[from + 2].value) &&
      is(tokens[from + 3], '(') &&
      is(tokens[from + 4], ')') &&
      to - from === 5
    )
      return { kind: 'literal', str: first, start, end };
    return { kind: 'other' };
  }
  if (
    isId(first, 'String') &&
    is(tokens[from + 1], '::') &&
    isId(tokens[from + 2], 'from') &&
    is(tokens[from + 3], '(') &&
    tokens[from + 4]?.t === 'str' &&
    is(tokens[from + 5], ')') &&
    to - from === 6
  )
    return { kind: 'literal', str: tokens[from + 4], start, end };
  if (
    isId(first, 'format') &&
    is(tokens[from + 1], '!') &&
    is(tokens[from + 2], '(') &&
    tokens[from + 3]?.t === 'str' &&
    matching(tokens, from + 2) === to - 1
  )
    return { kind: 'format', str: tokens[from + 3], start, end };
  // `format!("…").into()` / `.to_string()`
  if (
    isId(first, 'format') &&
    is(tokens[from + 1], '!') &&
    is(tokens[from + 2], '(') &&
    tokens[from + 3]?.t === 'str'
  ) {
    const close = matching(tokens, from + 2);
    if (is(tokens[close + 1], '.') && MAKERS.has(tokens[close + 2]?.value) && close + 5 === to)
      return { kind: 'format', str: tokens[from + 3], start, end, fmtEnd: tokens[close].end };
  }
  if (isId(first, 'crate') && is(tokens[from + 1], '::') && isId(tokens[from + 2], 'coded')) return { kind: 'coded' };
  if (isId(first, 'coded') && is(tokens[from + 1], '(')) return { kind: 'coded' };
  return { kind: 'other' };
}

/** Normalised message: `{…}` placeholders of format strings become `{}`. */
function normalise(strToken, format) {
  let value = strToken.value;
  if (!format) return value;
  return value
    .replace(/\{\{/g, '\u0001')
    .replace(/\}\}/g, '\u0002')
    .replace(/\{[^{}]*\}/g, '{}')
    .replace(/\u0001/g, '{')
    .replace(/\u0002/g, '}');
}

/** Token ranges of `#[cfg(test)] mod x { … }` blocks, which are not part of the product. */
function testRanges(tokens) {
  const ranges = [];
  for (let k = 0; k + 5 < tokens.length; k++) {
    if (
      is(tokens[k], '#') &&
      is(tokens[k + 1], '[') &&
      isId(tokens[k + 2], 'cfg') &&
      is(tokens[k + 3], '(') &&
      isId(tokens[k + 4], 'test')
    ) {
      // skip this attribute's `]`, then other attributes and `pub`: the item must be a `mod`
      let m = matching(tokens, k + 1) + 1;
      while (is(tokens[m], '#') && is(tokens[m + 1], '[')) m = matching(tokens, m + 1) + 1;
      if (isId(tokens[m], 'pub')) m += is(tokens[m + 1], '(') ? 1 + (matching(tokens, m + 1) - m) : 1;
      if (!isId(tokens[m], 'mod')) continue;
      let b = m;
      while (b < tokens.length && !is(tokens[b], '{') && !is(tokens[b], ';')) b++;
      if (is(tokens[b], '{')) ranges.push([tokens[k].start, tokens[matching(tokens, b)].end]);
    }
  }
  return ranges;
}

/**
 * Every sentence-carrying error construction in a Rust file.
 * A site: { file, line, form, message, dynamic, edit: { start, end, wrap, ... } }
 */
export function scanRs(file, text) {
  const tokens = tokenize(text);
  const skip = testRanges(tokens);
  const inTest = (pos) => skip.some(([a, b]) => pos >= a && pos < b);
  const lineOf = (pos) => {
    let line = 1;
    for (let k = 0; k < pos; k++) if (text.charCodeAt(k) === 10) line++;
    return line;
  };
  // line numbers: precompute offsets
  const offsets = [0];
  for (let k = 0; k < text.length; k++) if (text.charCodeAt(k) === 10) offsets.push(k + 1);
  const line = (pos) => {
    let lo = 0;
    let hi = offsets.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offsets[mid] <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo + 1;
  };
  void lineOf;
  const sites = [];
  const push = (form, cls, tokenIndex, extra = {}) => {
    const pos = tokens[tokenIndex].start;
    if (inTest(pos)) return;
    if (cls.kind === 'coded') return;
    if (cls.kind === 'other') {
      sites.push({
        file,
        line: line(pos),
        form,
        dynamic: true,
        message: '',
        argText: text.slice(extra.argStart, extra.argEnd).replace(/\s+/g, ' '),
        ...extra,
      });
      return;
    }
    sites.push({
      file,
      line: line(pos),
      form,
      dynamic: false,
      message: normalise(cls.str, cls.kind === 'format'),
      edit: {
        start: cls.start,
        end: cls.end,
        fmtEnd: cls.fmtEnd ?? cls.end,
        wrapFormat: cls.kind === 'format',
        strStart: cls.str.start,
        strEnd: cls.str.end,
        bare: cls.bare,
      },
      ...extra,
    });
  };
  for (let k = 0; k < tokens.length; k++) {
    const token = tokens[k];
    if (token.t !== 'id') continue;
    const prev = tokens[k - 1];
    // Err( … )
    if (token.value === 'Err' && is(tokens[k + 1], '(') && !(is(prev, '::') && false)) {
      const close = matching(tokens, k + 1);
      if (close < 0) continue;
      // `Err(x) =>` and `Err(x) if` are patterns, not constructions
      if (
        is(tokens[close + 1], '=>') ||
        isId(tokens[close + 1], 'if') ||
        is(tokens[close + 1], '=') ||
        isId(prev, 'let')
      )
        continue;
      const cls = classify(tokens, k + 2, close);
      // a bare one-word `Err("token")` is a `&str` marker for the caller to branch on, not a sentence for a person
      if (cls.kind === 'literal' && cls.bare && !/\s/.test(cls.str.value)) continue;
      push('Err', cls, k, { argStart: tokens[k + 2]?.start, argEnd: tokens[close - 1]?.end });
      continue;
    }
    // .ok_or( … )  /  .ok_or_else(|| …)  /  .map_err(|x| …)
    if (
      is(prev, '.') &&
      (token.value === 'ok_or' || token.value === 'ok_or_else' || token.value === 'map_err') &&
      is(tokens[k + 1], '(')
    ) {
      const close = matching(tokens, k + 1);
      if (close < 0) continue;
      if (token.value === 'ok_or') {
        const cls = classify(tokens, k + 2, close);
        if (cls.kind === 'other') continue;
        push('ok_or', cls, k, { nameStart: token.start, nameEnd: token.end });
      } else {
        // closure: `|| expr`, `|_| expr`, `|e| expr`
        let b = k + 2;
        let ok = false;
        if (is(tokens[b], '||')) ((ok = true), (b += 1));
        else if (is(tokens[b], '|')) {
          let e = b + 1;
          while (e < close && !is(tokens[e], '|')) e++;
          ok = true;
          b = e + 1;
        }
        if (!ok) continue;
        const cls = classify(tokens, b, close);
        if (
          cls.kind === 'other' ||
          cls.kind === 'coded' ||
          (cls.kind === 'literal' && cls.bare && !/\s/.test(cls.str.value))
        )
          continue;
        const closureArgs = text.slice(tokens[k + 2].start, tokens[b - 1].end);
        // `|error| format!("... {error}")` adds context to an error that may already carry a code
        const param = /^\|\s*([a-z_][a-z0-9_]*)\s*\|$/.exec(closureArgs)?.[1];
        const rewrap =
          param && param !== '_' && cls.kind === 'format' && cls.str.value.includes('{' + param + '}')
            ? param
            : undefined;
        push(token.value, cls, k, { closureArgs, rewrap });
      }
    }
  }
  return sites;
}

/** Every `coded("code", message)` and `rephrase("code", error, |said| format!("..."))` in the text. */
export function scanRsCoded(file, text) {
  const tokens = tokenize(text);
  const found = [];
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k];
    if (
      t.t !== 'id' ||
      (t.value !== 'coded' && t.value !== 'rephrase') ||
      !is(tokens[k + 1], '(') ||
      tokens[k + 2]?.t !== 'str'
    )
      continue;
    if (!is(tokens[k + 3], ',')) continue;
    const close = matching(tokens, k + 1);
    let cls;
    if (t.value === 'coded') cls = classify(tokens, k + 4, close);
    else {
      // rephrase("code", error, |said| format!("..."))
      let b = k + 4;
      while (b < close && !is(tokens[b], '|')) b++;
      b = tokens.findIndex((x, i) => i > b && is(x, '|')) + 1;
      cls = classify(tokens, b, close);
    }
    const line = text.slice(0, t.start).split('\n').length;
    if (cls.kind === 'literal' || cls.kind === 'format')
      found.push({
        file,
        code: tokens[k + 2].value,
        message: normalise(cls.str, cls.kind === 'format').replace(/\{said\}/g, '{}'),
        dynamic: false,
        line,
      });
    else found.push({ file, code: tokens[k + 2].value, message: '', dynamic: true, line });
  }
  return found;
}
