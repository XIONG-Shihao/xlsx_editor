/** Pure review of supplied accepted cells. Neither coverage, dictionary
 * availability nor write permission is inferred from a cached value. */
import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';
import { hasExactNumericValue } from './calculation.mjs';
import { refusal } from './refusal-codes.mjs';

export const WORKBOOK_REVIEW_LIMITS = Object.freeze({
  statisticsCells: 256,
  statisticsBytes: 1024 * 1024,
  textBytes: 1024 * 1024,
  wordCodePoints: 256,
  tokens: 100,
  workUnits: 32768,
  pageBytes: 32768,
  commandBytes: 1024 * 1024,
});
const kinds = ['blank', 'text', 'number', 'boolean', 'error', 'unavailable'],
  authorities = ['absent', 'literal', 'current', 'stored', 'pending', 'unresolved', 'circular', 'unqualified'],
  formulaStates = ['calculated', 'stored', 'pending', 'unresolved', 'error', 'circular', 'unqualified'];
const plain = (v) => v !== null && typeof v === 'object' && !Array.isArray(v),
  own = (v, k) => Object.hasOwn(v, k),
  bytes = (v) => new TextEncoder().encode(v).length;
function exact(value, keys, label) {
  if (!plain(value) || Object.keys(value).length !== keys.length || keys.some((key) => !own(value, key)))
    throw refusal('internal-invalid-field', 'Invalid ' + label + ' fields');
}
function text(value, max, label) {
  if (typeof value !== 'string' || !value || value.length > max || bytes(value) > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function xml(value) {
  if (typeof value !== 'string' || /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value))
    throw refusal('package-unsupported', 'Unsupported XML text in review');
  return value;
}
const boolean = (value, label) => {
  if (typeof value !== 'boolean') throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
};
function options(value) {
  if (!plain(value) || Object.keys(value).some((k) => !['revision', 'special', 'protected'].includes(k)))
    throw refusal('internal-invalid-field', 'Invalid review classification options');
  text(value.revision, 512, 'accepted revision');
  for (const key of ['special', 'protected']) if (own(value, key)) boolean(value[key], key);
  return value;
}
const scalarKind = (value) =>
  value === null || value === undefined
    ? 'blank'
    : typeof value === 'string'
      ? 'text'
      : typeof value === 'boolean'
        ? 'boolean'
        : typeof value === 'number' && Number.isFinite(value)
          ? 'number'
          : 'unavailable';
function storedKind(cell) {
  const cache = cell.cachedResult;
  if (cache?.authority === 'stored') {
    if (cache.status === 'error') return 'error';
    if (cache.status === 'calculated') return scalarKind(cache.value);
  }
  if (cell.cachedValue !== undefined && cell.cachedValue !== null)
    return cell.dataType === 'error' ? 'error' : scalarKind(cell.cachedValue);
  return 'unavailable';
}

export function classifyReviewCell(cell, context) {
  options(context);
  if (cell === undefined)
    return {
      presence: 'absent',
      authored: null,
      kind: 'blank',
      authority: 'absent',
      formulaState: null,
      special: false,
      protected: false,
      unsafeNumber: false,
      spellingEligible: false,
      reason: 'No authored record at this supplied position.',
    };
  if (!plain(cell)) throw refusal('internal-invalid-field', 'Invalid accepted review cell');
  const formula =
    (cell.formula !== undefined && cell.formula !== null) ||
    (cell.formulaType !== undefined && cell.formulaType !== null);
  const special =
    context.special === true ||
    (cell.formulaOutput !== undefined && cell.formulaOutput !== null) ||
    (cell.spill !== undefined && cell.spill !== null) ||
    cell.scenarioResult === true ||
    cell.scenario !== undefined ||
    cell.sharedFormulaId !== undefined ||
    (cell.formulaType !== undefined && cell.formulaType !== null && cell.formulaType !== 'normal');
  const protectedCell =
    context.protected === true || cell.protected === true || cell.readOnly === true || cell.protection !== undefined;
  const calculation = cell.calculation,
    derived = calculation !== undefined && calculation !== null;
  let kind = cell.dataType === 'error' || cell.error !== undefined ? 'error' : scalarKind(cell.value),
    authority = 'literal',
    formulaState = null,
    reason = null;
  const authored = formula ? 'formula' : cell.value === null || cell.value === undefined ? 'blank' : 'literal';
  const unsafeNumber =
    typeof cell.value === 'number' &&
    (!Number.isFinite(cell.value) ||
      (!Number.isSafeInteger(cell.value) && Number.isInteger(cell.value)) ||
      !hasExactNumericValue(cell.value, cell.rawValue));
  if (calculation?.pending === true) {
    authority = 'pending';
    kind = storedKind(cell);
    formulaState = formula ? 'pending' : null;
    reason = 'Stored result is pending calculation; it is not current.';
  } else if (formula || derived || special) {
    if (!calculation || calculation.status === 'stored') {
      authority = 'stored';
      kind = storedKind(cell);
      if (kind === 'unavailable' && special && cell.value !== undefined && cell.value !== null)
        kind = cell.dataType === 'error' ? 'error' : scalarKind(cell.value);
      formulaState = formula ? 'stored' : null;
      reason = 'Stored cache has not been certified as a current formula result.';
    } else if (calculation.revision !== context.revision) {
      authority = 'unqualified';
      kind = 'unavailable';
      formulaState = formula ? 'unqualified' : null;
      reason = 'Calculation does not match the captured accepted revision.';
    } else if (calculation.status === 'calculated') {
      kind = scalarKind(calculation.value);
      if (kind === 'unavailable' || kind === 'blank') {
        authority = 'unqualified';
        formulaState = formula ? 'unqualified' : null;
        reason = 'A current calculated scalar is missing or invalid.';
      } else {
        authority = 'current';
        formulaState = formula ? 'calculated' : null;
      }
    } else if (calculation.status === 'error') {
      authority = 'current';
      kind = typeof calculation.error === 'string' && calculation.error ? 'error' : 'unavailable';
      formulaState = formula ? 'error' : null;
      if (kind === 'unavailable') {
        authority = 'unqualified';
        formulaState = formula ? 'unqualified' : null;
        reason = 'Current error result is missing its typed code.';
      }
    } else if (['unresolved', 'circular'].includes(calculation.status)) {
      authority = calculation.status;
      kind = 'unavailable';
      formulaState = formula ? calculation.status : null;
      reason =
        calculation.status === 'circular'
          ? 'Circular formula; no current scalar.'
          : 'Unresolved formula; no current scalar.';
    } else {
      authority = 'unqualified';
      kind = 'unavailable';
      formulaState = formula ? 'unqualified' : null;
      reason = 'Unknown calculation result authority.';
    }
  } else if (cell.dataType !== undefined && cell.dataType !== kind) {
    authority = 'unqualified';
    kind = 'unavailable';
    reason = 'Authored scalar and declared type disagree.';
  }
  const literalText = !formula && !derived && kind === 'text' && authority === 'literal';
  let spellingEligible = literalText && !special && !protectedCell;
  if (spellingEligible && bytes(cell.value) > WORKBOOK_REVIEW_LIMITS.textBytes) {
    spellingEligible = false;
    reason = 'Literal text exceeds the spelling cell byte limit.';
  }
  if (protectedCell) reason = 'Known protected cell is excluded from correction.';
  else if (special) reason = 'Known array, shared, spill or special output is excluded from correction.';
  else if (!spellingEligible && !reason)
    reason = literalText ? 'Text is unavailable for spelling.' : 'Only authored literal text is eligible for spelling.';
  return {
    presence: 'authored',
    authored,
    kind,
    authority,
    formulaState,
    special: !!special,
    protected: protectedCell,
    unsafeNumber,
    spellingEligible,
    reason,
  };
}

const zeroKinds = () => Object.fromEntries(kinds.map((k) => [k, 0]));
export function reviewStatistics(cells, context) {
  options(context);
  const initial = {
    schema: 1,
    revision: context.revision,
    positions: 0,
    authoredCells: 0,
    absentPositions: 0,
    authoredBlankCells: 0,
    literalCells: 0,
    formulaCells: 0,
    specialCells: 0,
    protectedCells: 0,
    unsafeNumericCells: 0,
    spellingEligibleCells: 0,
    literalKinds: zeroKinds(),
    formulaStates: Object.fromEntries(formulaStates.map((k) => [k, 0])),
    results: Object.fromEntries(authorities.map((a) => [a, zeroKinds()])),
  };
  return addReviewStatistics(initial, cells, context);
}
export function addReviewStatistics(summary, cells, context) {
  options(context);
  exact(
    summary,
    [
      'schema',
      'revision',
      'positions',
      'authoredCells',
      'absentPositions',
      'authoredBlankCells',
      'literalCells',
      'formulaCells',
      'specialCells',
      'protectedCells',
      'unsafeNumericCells',
      'spellingEligibleCells',
      'literalKinds',
      'formulaStates',
      'results',
    ],
    'statistics'
  );
  if (
    summary.schema !== 1 ||
    summary.revision !== context.revision ||
    !Array.isArray(cells) ||
    cells.length > WORKBOOK_REVIEW_LIMITS.statisticsCells ||
    bytes(JSON.stringify(cells)) > WORKBOOK_REVIEW_LIMITS.statisticsBytes
  )
    throw refusal('workbook-limit', 'Statistics revision/page bound changed');
  for (const key of [
    'positions',
    'authoredCells',
    'absentPositions',
    'authoredBlankCells',
    'literalCells',
    'formulaCells',
    'specialCells',
    'protectedCells',
    'unsafeNumericCells',
    'spellingEligibleCells',
  ])
    if (!Number.isSafeInteger(summary[key]) || summary[key] < 0)
      throw refusal('internal-invalid-field', 'Invalid statistics count');
  for (const [value, keys] of [
    [summary.literalKinds, kinds],
    [summary.formulaStates, formulaStates],
    ...authorities.map((a) => [summary.results[a], kinds]),
  ]) {
    exact(value, keys, 'statistics counters');
    for (const n of Object.values(value))
      if (!Number.isSafeInteger(n) || n < 0) throw refusal('internal-invalid-field', 'Invalid statistics counter');
  }
  exact(summary.results, authorities, 'statistics result authorities');
  const sum = (value) => Object.values(value).reduce((n, v) => n + v, 0);
  if (
    summary.positions !== summary.authoredCells + summary.absentPositions ||
    summary.authoredCells !== summary.authoredBlankCells + summary.literalCells + summary.formulaCells ||
    sum(summary.literalKinds) !== summary.literalCells ||
    sum(summary.formulaStates) !== summary.formulaCells ||
    sum(Object.fromEntries(authorities.map((a) => [a, sum(summary.results[a])]))) !== summary.positions ||
    ['specialCells', 'protectedCells', 'unsafeNumericCells', 'spellingEligibleCells'].some(
      (k) => summary[k] > summary.authoredCells
    ) ||
    summary.positions + cells.length > Number.MAX_SAFE_INTEGER
  )
    throw refusal('workbook-unsupported', 'Inconsistent statistics population');
  const next = {
    ...summary,
    literalKinds: { ...summary.literalKinds },
    formulaStates: { ...summary.formulaStates },
    results: Object.fromEntries(authorities.map((a) => [a, { ...summary.results[a] }])),
  };
  for (const cell of cells) {
    const c = classifyReviewCell(cell, context);
    next.positions++;
    next.results[c.authority][c.kind]++;
    if (c.presence === 'absent') {
      next.absentPositions++;
      continue;
    }
    next.authoredCells++;
    if (c.authored === 'blank') next.authoredBlankCells++;
    else if (c.authored === 'literal') {
      next.literalCells++;
      next.literalKinds[c.kind]++;
    } else {
      next.formulaCells++;
      next.formulaStates[c.formulaState]++;
    }
    if (c.special) next.specialCells++;
    if (c.protected) next.protectedCells++;
    if (c.unsafeNumber) next.unsafeNumericCells++;
    if (c.spellingEligible) next.spellingEligibleCells++;
  }
  return next;
}

const letter = /\p{L}/u,
  number = /\p{N}/u,
  mark = /\p{M}/u,
  apostrophe = (c) => c === "'" || c === '\u2019';
const at = (value, i) => (i < value.length ? String.fromCodePoint(value.codePointAt(i)) : '');
function previous(value, i) {
  if (i === 0) return '';
  const n = value.charCodeAt(i - 1);
  return i > 1 && n >= 0xdc00 && n <= 0xdfff ? value.slice(i - 2, i) : value[i - 1];
}
const startChar = (c) => !!c && (letter.test(c) || number.test(c));
const continuation = (value, i) => {
  const c = at(value, i);
  return (
    !!c && (letter.test(c) || number.test(c) || mark.test(c) || (apostrophe(c) && startChar(at(value, i + c.length))))
  );
};
const boundary = (value, i) =>
  i === 0 ||
  i === value.length ||
  !(
    value.charCodeAt(i) >= 0xdc00 &&
    value.charCodeAt(i) <= 0xdfff &&
    value.charCodeAt(i - 1) >= 0xd800 &&
    value.charCodeAt(i - 1) <= 0xdbff
  );
function initialCursor() {
  return { offset: 0, start: null, points: 0, hasNumber: false, oversized: false, ready: false };
}
function cursor(value, raw) {
  const c = raw === undefined ? initialCursor() : raw;
  exact(c, ['offset', 'start', 'points', 'hasNumber', 'oversized', 'ready'], 'spelling cursor');
  if (
    !Number.isSafeInteger(c.offset) ||
    c.offset < 0 ||
    c.offset > value.length ||
    !boundary(value, c.offset) ||
    !Number.isSafeInteger(c.points) ||
    c.points < 0 ||
    c.points > 257 ||
    typeof c.hasNumber !== 'boolean' ||
    typeof c.oversized !== 'boolean' ||
    typeof c.ready !== 'boolean'
  )
    throw refusal('internal-invalid-field', 'Invalid spelling cursor');
  if (c.start === null) {
    const prior = previous(value, c.offset);
    if (
      c.points ||
      c.hasNumber ||
      c.oversized ||
      c.ready ||
      (continuation(value, c.offset) &&
        (startChar(prior) ||
          mark.test(prior) ||
          (apostrophe(prior) && c.offset > 1 && startChar(previous(value, c.offset - 1)))))
    )
      throw refusal('internal-invalid-field', 'Invalid empty spelling cursor');
  } else {
    if (
      !Number.isSafeInteger(c.start) ||
      c.start < 0 ||
      c.start >= c.offset ||
      !boundary(value, c.start) ||
      !startChar(at(value, c.start)) ||
      startChar(previous(value, c.start)) ||
      mark.test(previous(value, c.start)) ||
      (apostrophe(previous(value, c.start)) && c.start > 1 && startChar(previous(value, c.start - 1))) ||
      !c.points ||
      c.oversized !== (c.points === 257)
    )
      throw refusal('internal-invalid-field', 'Invalid partial spelling cursor');
    if (!c.oversized) {
      const chars = Array.from(value.slice(c.start, c.offset));
      if (
        chars.length !== c.points ||
        chars.length > 256 ||
        chars.some((char, index) => !startChar(char) && !mark.test(char) && !(apostrophe(char) && index > 0)) ||
        chars.some((char) => number.test(char)) !== c.hasNumber
      )
        throw refusal('workbook-bad-input', 'Spelling cursor does not match original text');
    }
    if (c.ready && continuation(value, c.offset))
      throw refusal('workbook-malformed', 'Ready spelling token ended inside a word');
  }
  return { ...c };
}
/** Original UTF-16 offsets; a pending or ready token carries across work/byte
 * pages. No normalized copy, substring fragments or silent skipped word. */
export function reviewTokens(value, rawCursor, limits = {}) {
  xml(value);
  if (!plain(limits) || Object.keys(limits).some((k) => !['maxTokens', 'maxWork', 'maxBytes'].includes(k)))
    throw refusal('internal-invalid-field', 'Invalid spelling page limits');
  const maxTokens = limits.maxTokens ?? 100,
    maxWork = limits.maxWork ?? 32768,
    maxBytes = limits.maxBytes ?? 32768;
  for (const [n, min, max] of [
    [maxTokens, 1, 100],
    [maxWork, 2, 32768],
    [maxBytes, 2048, 32768],
  ])
    if (!Number.isInteger(n) || n < min || n > max)
      throw refusal('internal-invalid-field', 'Invalid spelling page bound');
  let c = cursor(value, rawCursor),
    work = 0;
  const tokens = [],
    excluded = [];
  const make = (next, reason) => ({
    schema: 1,
    tokens,
    excluded,
    nextCursor: next,
    coverage: { complete: next === null && reason === null, reason, offset: next?.offset ?? c.offset, workUnits: work },
  });
  if (bytes(value) > WORKBOOK_REVIEW_LIMITS.textBytes)
    return {
      schema: 1,
      tokens,
      excluded: [{ start: 0, end: value.length, reason: 'cell-text-limit' }],
      nextCursor: null,
      coverage: { complete: false, reason: 'cell-text-limit', offset: 0, workUnits: 0 },
    };
  for (;;) {
    if (c.start !== null && (c.ready || c.offset === value.length || !continuation(value, c.offset))) {
      c.ready = true;
      const item = c.oversized
        ? { start: c.start, end: c.offset, reason: 'token-too-long' }
        : c.hasNumber
          ? { start: c.start, end: c.offset, reason: 'numeric-or-identifier' }
          : { text: value.slice(c.start, c.offset), start: c.start, end: c.offset };
      if (tokens.length >= maxTokens || excluded.length >= 100)
        return make(c, tokens.length >= maxTokens ? 'token-limit' : 'exclusion-limit');
      const list = 'reason' in item ? excluded : tokens;
      list.push(item);
      if (bytes(JSON.stringify(make(c, 'byte-limit'))) > maxBytes - 512) {
        list.pop();
        return make(c, 'byte-limit');
      }
      c = { offset: c.offset, start: null, points: 0, hasNumber: false, oversized: false, ready: false };
      continue;
    }
    if (c.offset === value.length) return make(null, null);
    const char = at(value, c.offset);
    if (work + char.length > maxWork) return make(c, 'work-limit');
    if (c.start === null && startChar(char)) c.start = c.offset;
    if (c.start !== null) {
      c.points = Math.min(257, c.points + 1);
      c.hasNumber ||= number.test(char);
      c.oversized = c.points === 257;
    }
    c.offset += char.length;
    work += char.length;
  }
}

export function validateReviewReceipt(receipt, expected) {
  exact(
    receipt,
    ['documentId', 'ownerId', 'state', 'revision', 'locale', 'providerId', 'dictionaryId'],
    'review receipt'
  );
  for (const key of ['documentId', 'ownerId', 'state', 'revision'])
    text(receipt[key], key === 'revision' ? 512 : 128, 'review ' + key);
  if (
    !['en-US', 'en-GB'].includes(receipt.locale) ||
    !['providerId', 'dictionaryId'].every(
      (key) => typeof receipt[key] === 'string' && /^[a-f0-9]{64}$/.test(receipt[key])
    )
  )
    throw refusal('workbook-malformed', 'Invalid verified spelling locale/provider identity');
  if (expected !== undefined) {
    validateReviewReceipt(expected);
    if (Object.keys(receipt).some((k) => receipt[k] !== expected[k]))
      throw refusal('workbook-stale', 'Captured review owner/head/locale/provider retired');
  }
  return { ...receipt };
}
function target(value) {
  exact(value, ['sheet', 'cell'], 'review target');
  text(value.sheet, 124, 'review worksheet');
  if (value.sheet.length > 31 || /[\\/*?:\[\]\u0000]/.test(value.sheet))
    throw refusal('internal-invalid-field', 'Invalid review worksheet');
  const m = typeof value.cell === 'string' && /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(value.cell);
  if (!m || Number(m[2]) > 1048576 || [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) > 16384)
    throw refusal('workbook-bad-input', 'Choose one canonical review cell');
  return { ...value };
}
function identity(cell) {
  xml(cell.value);
  if (cell.rawValue !== undefined && cell.rawValue !== null) xml(cell.rawValue);
  return bytesToHex(
    sha256(
      new TextEncoder().encode(
        JSON.stringify(['literal-review-v1', cell.value, cell.rawValue ?? null, cell.dataType ?? null])
      )
    )
  );
}
function token(value, t) {
  exact(t, ['text', 'start', 'end'], 'original spelling token');
  if (
    !Number.isSafeInteger(t.start) ||
    !Number.isSafeInteger(t.end) ||
    t.start < 0 ||
    t.end <= t.start ||
    t.end > value.length ||
    typeof t.text !== 'string' ||
    t.text.length > 512 ||
    !boundary(value, t.start) ||
    !boundary(value, t.end) ||
    value.slice(t.start, t.end) !== t.text ||
    Array.from(t.text).length > 256 ||
    !startChar(at(value, t.start)) ||
    startChar(previous(value, t.start)) ||
    mark.test(previous(value, t.start)) ||
    (apostrophe(previous(value, t.start)) && t.start > 1 && startChar(previous(value, t.start - 1)))
  )
    throw refusal('workbook-stale', 'Spelling token no longer matches the original text');
  let end = t.start;
  while (end < t.end && (end === t.start ? startChar(at(value, end)) : continuation(value, end)))
    end += at(value, end).length;
  if (end !== t.end || continuation(value, t.end) || Array.from(t.text).some((c) => number.test(c)))
    throw refusal('workbook-refused', 'Spelling token is partial or numeric');
  return { ...t };
}
export function captureSpellingFinding(input) {
  exact(input, ['receipt', 'target', 'cell', 'token'], 'spelling capture');
  const { receipt: rawReceipt, target: rawTarget, cell, token: rawToken } = input;
  const receipt = validateReviewReceipt(rawReceipt),
    where = target(rawTarget),
    c = classifyReviewCell(cell, { revision: receipt.revision });
  if (!c.spellingEligible) throw refusal('workbook-refused', c.reason);
  return { schema: 1, receipt, target: where, cellIdentity: identity(cell), token: token(cell.value, rawToken) };
}
export function planSpellingReplacement(finding, current) {
  exact(finding, ['schema', 'receipt', 'target', 'cellIdentity', 'token'], 'spelling finding');
  if (finding.schema !== 1 || !/^[a-f0-9]{64}$/.test(finding.cellIdentity))
    throw refusal('workbook-malformed', 'Invalid spelling finding identity');
  exact(
    current,
    [
      'receipt',
      'cell',
      'replacement',
      'writable',
      ...Object.keys(current ?? {}).filter((k) => ['special', 'protected'].includes(k)),
    ],
    'replacement context'
  );
  validateReviewReceipt(current.receipt, finding.receipt);
  if (current.writable !== true) throw refusal('workbook-protected', 'This captured workbook is read-only');
  const c = classifyReviewCell(current.cell, {
    revision: current.receipt.revision,
    ...(own(current, 'special') ? { special: current.special } : {}),
    ...(own(current, 'protected') ? { protected: current.protected } : {}),
  });
  if (!c.spellingEligible) throw refusal('workbook-refused', c.reason);
  const where = target(finding.target);
  if (identity(current.cell) !== finding.cellIdentity)
    throw refusal('workbook-stale', 'Original literal value or raw source changed');
  const span = token(current.cell.value, finding.token);
  text(current.replacement, 1024, 'spelling replacement');
  xml(current.replacement);
  if (Array.from(current.replacement).length > 256)
    throw refusal('workbook-limit', 'Spelling replacement exceeds its bound');
  const value = current.cell.value.slice(0, span.start) + current.replacement + current.cell.value.slice(span.end);
  if (value === current.cell.value) throw refusal('workbook-bad-input', 'Spelling replacement makes no change');
  const command = {
    type: 'setCells',
    sheet: where.sheet,
    cells: [
      {
        cell: where.cell,
        value,
        expectedValue: current.cell.value,
        expectedFormula: null,
        expectedRawValue: current.cell.rawValue ?? null,
      },
    ],
  };
  if (bytes(JSON.stringify(command)) > WORKBOOK_REVIEW_LIMITS.commandBytes)
    throw refusal('workbook-limit', 'Correction command exceeds its byte bound; this text is inspection-only');
  return command;
}
