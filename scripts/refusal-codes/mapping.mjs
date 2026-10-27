// Which stable code each refusal site gets.
//
// A site is found by where it is (file) and by the shape of its sentence (message with computed parts
// shown as `{}`). A code names the cause as a pair: WHAT part of the workbook or editor it concerns (the
// domain) and WHAT went wrong (the kind). Rules run in this order and the first match wins:
//   1. EXCEPTIONS: one site or small family with a code of its own (checked first)
//   2. the kind from KIND_RULES
//   3. the domain from FILE_DOMAINS (specialised files) or MESSAGE_DOMAINS (general files)
// Kinds that start with `internal-` have no domain: they mark broken bookkeeping, not a person's situation.
// The code is then `${domain}-${kind}`; every code must exist in core/refusal-codes.mjs.
//
// The mapping decides the code only from the sentence's role (what kind of problem it states), not by
// turning the sentence into a code. Rewording a sentence keeps its code as long as it still matches the
// same rule; a new sentence that no rule matches is reported by check.mjs until a rule covers it.

/**
 * Not refusals: the formula engine reports an Excel error value (#NUM!, #DIV/0!) as an `Err` so that the
 * calling formula can handle it (IFERROR and friends test for the `Excel error: ` prefix). These must keep
 * their exact text and carry no code.
 */
export const SIGNALS = [{ file: /\.rs$/, message: /^Excel error: / }];

/** Sites with a code of their own. { file?, message, code } */
export const EXCEPTIONS = [
  { message: /^The spreadsheet operation was stopped before its next commit/, code: 'operation-cancelled' },
];

/**
 * Sites whose message is not written there: `Error(message)`, `Err(reason)`. Checked in this order:
 *   forwarded: the error was made elsewhere (a worker reply, a caught error); it keeps whatever code it had
 *   legacy:    a helper that already attaches an older upper-case code of its own
 *   xml:       the XML parser's complaint about a part of the workbook file
 *   otherwise: a stored reason string; the code is the file's domain with the kind `refused`
 */
export const DYNAMIC = [
  {
    arg: /^(error|e|err|_|message\.error|reply\.error|projectError|err\(.*\)|e\.into\(\)|this\.#warning( \?\? .*)?|rusqlite::.*)$/,
    how: 'forwarded',
  },
  { file: /print-layout|print-snapshot/, arg: /replaceAll/, how: 'legacy' },
  {
    file: /worksheet-xml|workbook-theme|spill-metadata|indexed-tables|indexed-accessibility|indexed-copy|ink-stroke-xml|package-capabilities|workbook-package|note-zip-reader|workbook-spelling-jobs|indexed-notes/,
    arg: /^(message|m)$/,
    how: 'code',
    code: 'package-malformed',
  },
];

// ---------------------------------------------------------------------------------------------------------------
// Kinds
// ---------------------------------------------------------------------------------------------------------------
const FILE_WORDS =
  /\b(XML|attribute|relationship|content[- ]type|ZIP|shared[- ]string|namespace|root|tag|quote|span|element|archive|package|part\b|CDATA|entity|escape|record)/i;

export const KIND_RULES = [
  // Our own bookkeeping.
  ['internal-missing-field', /^Missing\b|\bis missing\b|\bare missing\b|\bmissing$/],
  ['internal-unknown-reference', /^Unknown\b/],
  // The workbook file itself is not shaped as expected.
  [
    'malformed',
    /^(Invalid|Malformed|Unexpected|Unbalanced|Truncated|Noncanonical|Unproved|Text outside)\b.*(?:XML|attribute|relationship|content[- ]type|ZIP|shared[- ]string|namespace|root|tag|quote|span|element|archive|part\b|CDATA|entity|escape|container|boundary|path|target|index)|^(Malformed|Unexpected|Unbalanced|Truncated|Text outside|Unproved|Unterminated|Unquoted|Unclosed|Escaping|Damaged|Degenerate|Cyclic|Reversed|Unordered)\b|\bmultiple .* roots\b|^Multiple\b.*\broots?\b|\bentities\b|\bdamaged\b|surrogate|ended (early|during|inside)|wrong location kind|^Mixed\b.*\btext|^Text in\b|\bhas multiple roots\b/i,
  ],
  ['internal-invalid-field', /^Invalid\b|\bis invalid\b|^Noncanonical\b/],
  // Sessions and operations.
  ['operation-cancelled', /cancel/i],
  ['operation-timed-out', /timed out|deadline/i],
  ['session-busy', /\bbusy\b|being updated|still stopping|in progress|already running|retry/i],
  [
    'session-closed',
    /\b(is|was|has been|already|been) (opened or )?closed\b|closed before|\bstopped\b|has stopped|not open|not been opened|opened or closed/i,
  ],
  // The shape of the problem.
  ['not-found', /not found|no such|does not exist|doesn't exist|could not find|cannot find/i],
  ['ambiguous', /\bAmbiguous\b/],
  ['duplicate', /^Duplicate\b|\bduplicate\b.*\b(key|name|id|member)\b|already (exists|used|has|in use)/],
  [
    'limit',
    /exceed|\bbound\b|\bbounds\b|budget|\blimit|overflow|too (large|many|long|big|deep)|at most|up to|more than|capacity|queue is full|\bcap\b|oversize|beyond/i,
  ],
  [
    'stale',
    /\bchanged\b|\bexpired\b|\bstale\b|superseded|retired|no longer|changed (while|since|after|between)|belongs to another|outdated|current revision|\bnewer\b/i,
  ],
  ['conflict', /would hide|overlap|collid|collision|occupied|intersect|in the way|would (replace|overwrite)/i],
  ['protected', /protect|read-only|readonly/i],
  ['unsupported', /^\{\}: /],
  [
    'internal-mismatch',
    /\bdiffers?\b|mismatch|disagree|no longer agrees|\blost an? |^Transfer lost|^Package rewrite (moved|changed)|\bchanged (an |the |its )?(existing|undeclared|source|stored|nonlink|immutable|different|formula|authored|literal|value|style)|\bchanges (an |the |its |a )?(authored|existing|stored|literal|formula|style|source)/i,
  ],
  [
    'unsupported',
    /unsupported|not supported|isn't supported|unavailable|not qualified|qualif|opaque|extended|foreign|private|noncanonical|unsafe|incomplete|unrelated|native-owned|\bgated\b|preserv|not admitted|outside|only (supports|one|a|the)|cannot|can't|not allowed|rich|imported|shared|special|nested|orphan|unbound|inconsistent|legacy|differs|mismatch|changed|disagree|unreliable|different|escapes|belongs|reserved|disconnected|excessive|deleted|unresolved|unconfined|non-numeric|nonfinite|existing|alternative|unused|degenerate/i,
  ],
  [
    'bad-input',
    /^(Choose|Select|Use|Enter|Preview|Pick|Open|Save|Type|Name|Specify|Provide|Wait|Keep|Make|Submit|Activate|Calculate|Recover|Close|Symbols|Make)\b|\bmust\b|\bneeds?\b|\brequires?\b|\bshould\b|\bempty\b|\bno\b|\bnot\b|\bonly\b/i,
  ],
];

// ---------------------------------------------------------------------------------------------------------------
// Domains
// ---------------------------------------------------------------------------------------------------------------
/** Specialised files: the file decides. */
export const FILE_DOMAINS = [
  [/pivot/i, 'pivot'],
  [/print|page-theme|page_style|page-style|header-footer/i, 'print'],
  [/spill|scalars|scalar-proof|sequence/i, 'spill'],
  [/hyperlink/i, 'hyperlink'],
  [/(^|[/_-])notes?([/_.-]|$)|note-|note_|accessibility/i, 'note'],
  [/rules|conditional|validation/i, 'rules'],
  [/filter|owned_filter/i, 'filter'],
  [/sort-keys|sort/i, 'sort'],
  [/merge/i, 'merge'],
  [/name_update|name_addition|defined-name|indexed-names|names/i, 'name'],
  [/\btables?\b|tables\/|indexed-tables|table/i, 'table'],
  [/chart/i, 'chart'],
  [/model3d|smartart|ink|shape|image|background|managed|object|drawing/i, 'object'],
  [/csv|delimited/i, 'csv'],
  [/query|local-actions|local-action/i, 'query'],
  [/paste/i, 'paste'],
  [/dedup|duplicate/i, 'duplicates'],
  [/fill|bulk|across|values/i, 'fill'],
  [/theme|style|font|format/i, 'style'],
  [/used_range|used-range|worksheet-geometry|layout/i, 'layout'],
  [/calculation|financial|formula|dependencies|trace|date|scenarios|decimal/i, 'formula'],
  [/spelling/i, 'spelling'],
  [
    /recovery|durability|quota|cache|checkpoint|storage|session|worker|owner|relay|client|runtime|gate|retirement|tab-registry/i,
    'session',
  ],
  [/zip|xml|package|raw_read|passive|import/i, 'package'],
];

/** General files: the sentence decides (first match). */
export const MESSAGE_DOMAINS = [
  [/pivot/i, 'pivot'],
  [/\bprint|header\/footer|page setup|page break|PRINT_/i, 'print'],
  [/spill|SEQUENCE|dynamic array|scalar array|array formula|dynamic metadata/i, 'spill'],
  [/hyperlink|\blink\b/i, 'hyperlink'],
  [/\bnotes?\b|comment/i, 'note'],
  [/conditional|validation|\brules?\b/i, 'rules'],
  [/filter|autofilter/i, 'filter'],
  [/\bsort/i, 'sort'],
  [/\bmerge/i, 'merge'],
  [/\bname/i, 'name'],
  [/\btable|header row/i, 'table'],
  [/chart/i, 'chart'],
  [/3D|GLB|SmartArt|\bink\b|stroke|shape|image|PNG|picture|drawing|WordArt|background|\bobject/i, 'object'],
  [/CSV|delimited/i, 'csv'],
  [/query|connection|Refresh All/i, 'query'],
  [/paste/i, 'paste'],
  [/theme|style|font|number format|color/i, 'style'],
  [/UsedRange|used range|row height|column width|freeze/i, 'layout'],
  [/formula|calculation|dependenc/i, 'formula'],
  [/recovery|session|checkpoint|SQLite|database|cache|durab/i, 'session'],
  [/ZIP|XML|relationship|content type|package|part\b/i, 'package'],
];

export const DEFAULT_DOMAIN = 'workbook';
