/**
 * The registry of refusal codes: every reason the editor can refuse, in one list.
 *
 * A refusal carries a stable code so that a caller (and a test) decides on the code and shows the sentence:
 * a sentence may be reworded, a code may not change meaning. The engine writes `[code] sentence` and its
 * helper sends the two apart; code written in JavaScript uses `refusal(code, sentence)`.
 *
 * The rule for what gets which code:
 *   - A code names the CAUSE and is shared by every place with that cause, as `<part>-<what went wrong>`:
 *     the part of the workbook or editor it concerns (pivot, table, print, spill, ...) and one of
 *     unsupported, protected, limit, stale, conflict, bad-input, not-found, malformed, ambiguous, duplicate,
 *     refused (the catch-all when no narrower kind fits).
 *   - `internal-*` codes are not a person's situation: the editor's own bookkeeping failed (a plan or
 *     stored field missing or malformed, an internal check disagreeing). They are shared by all such
 *     places. Show a person "something went wrong", log the sentence, and treat it as a bug to fix.
 *   - `operation-*` and `session-*` say the operation or workbook session itself ended or is busy.
 *   - A code is never made from a sentence's words. scripts/refusal-codes/mapping.mjs decides, by rule, which
 *     code each place gets; `node scripts/refusal-codes/check.mjs` lists places that have none, and
 *     tests/refusal-codes.mjs keeps the list, the code and the engine in step.
 *   - To add a code, add it here with its meaning. Never change what a code means; add a new one.
 */
export const REFUSAL_CODES = Object.freeze({
  // Codes that existed before this registry was completed.
  'needs-structural-writer': 'The preserving writer cannot make this change; the caller may try the general writer.',
  'session-disk-budget': 'The workbook session has used its disk allowance; save and reopen.',
  'filtered-column': 'A column with filter criteria cannot be deleted; clear its filter first.',
  'sheet-protected': 'The sheet is protected; Excel refuses the same change until the sheet is unprotected.',
  'table-or-protection': 'A table on the sheet, or sheet protection, blocks this change.',
  'pivot-source-deleted': "The rows or columns are all of a pivot table's data; delete the pivot table first.",
  'object-range-deleted': 'The change would delete everything a link, note, rule or chart refers to.',
  'object-not-movable': 'Something on the sheet cannot yet follow inserted or deleted rows and columns.',
  'kept-recovery-full': 'The list of kept recovery files is full; save or discard some first.',

  // Internal: the editor's own bookkeeping failed. Not a person's situation.
  'internal-invalid-field':
    'A plan, record or stored field the editor itself builds is malformed. A bug, not something a person can fix.',
  'internal-mismatch':
    'An internal consistency check failed: two things the editor keeps in step disagree. A bug, not something a person can fix.',
  'internal-missing-field':
    'A plan, record or stored field the editor itself builds is missing. A bug, not something a person can fix.',
  'internal-unknown-reference':
    'An internal reference (worksheet, version, record) does not resolve. A bug or a stale request, not something a person can fix.',
  // Operations and sessions
  'operation-cancelled': 'The operation was cancelled before it finished; nothing was committed.',
  'operation-timed-out': 'The operation ran past its deadline and was stopped.',
  'session-busy': 'The workbook is busy with another operation; retry when it has finished.',
  'session-closed': 'The workbook, owner or client was already closed or stopped; reopen it.',
  // chart
  'chart-ambiguous': 'Stored content for a chart can be read in more than one way; the editor will not guess.',
  'chart-bad-input': 'The request about a chart needs different input from the person (a choice, name or range).',
  'chart-conflict': 'The change to a chart would collide with content that is already there.',
  'chart-duplicate': 'Something that must be unique appears twice for a chart.',
  'chart-limit': 'A size, count or time limit would be exceeded for a chart.',
  'chart-malformed': 'Stored content for a chart is not shaped as expected and cannot be read safely.',
  'chart-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a chart.',
  'chart-refused': 'The editor declined a change to a chart; the sentence says why. The cause fits no narrower kind.',
  'chart-stale': 'The request about a chart refers to a state that has changed or expired; refresh and retry.',
  'chart-unsupported':
    'The editor will not change a chart this way: outside what it can change safely. The sentence says what.',

  // csv
  'csv-bad-input':
    'The request about a CSV or delimited file needs different input from the person (a choice, name or range).',
  'csv-duplicate': 'Something that must be unique appears twice for a CSV or delimited file.',
  'csv-limit': 'A size, count or time limit would be exceeded for a CSV or delimited file.',
  'csv-malformed': 'Stored content for a CSV or delimited file is not shaped as expected and cannot be read safely.',
  'csv-refused':
    'The editor declined a change to a CSV or delimited file; the sentence says why. The cause fits no narrower kind.',
  'csv-stale':
    'The request about a CSV or delimited file refers to a state that has changed or expired; refresh and retry.',
  'csv-unsupported':
    'The editor will not change a CSV or delimited file this way: outside what it can change safely. The sentence says what.',

  // duplicates
  'duplicates-ambiguous':
    'Stored content for removing duplicates can be read in more than one way; the editor will not guess.',
  'duplicates-bad-input':
    'The request about removing duplicates needs different input from the person (a choice, name or range).',
  'duplicates-conflict': 'The change to removing duplicates would collide with content that is already there.',
  'duplicates-duplicate': 'Something that must be unique appears twice for removing duplicates.',
  'duplicates-limit': 'A size, count or time limit would be exceeded for removing duplicates.',
  'duplicates-malformed': 'Stored content for removing duplicates is not shaped as expected and cannot be read safely.',
  'duplicates-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to removing duplicates.',
  'duplicates-refused':
    'The editor declined a change to removing duplicates; the sentence says why. The cause fits no narrower kind.',
  'duplicates-unsupported':
    'The editor will not change removing duplicates this way: outside what it can change safely. The sentence says what.',

  // fill
  'fill-ambiguous':
    'Stored content for a fill, bulk copy or Across Worksheets action can be read in more than one way; the editor will not guess.',
  'fill-bad-input':
    'The request about a fill, bulk copy or Across Worksheets action needs different input from the person (a choice, name or range).',
  'fill-conflict':
    'The change to a fill, bulk copy or Across Worksheets action would collide with content that is already there.',
  'fill-duplicate': 'Something that must be unique appears twice for a fill, bulk copy or Across Worksheets action.',
  'fill-limit': 'A size, count or time limit would be exceeded for a fill, bulk copy or Across Worksheets action.',
  'fill-malformed':
    'Stored content for a fill, bulk copy or Across Worksheets action is not shaped as expected and cannot be read safely.',
  'fill-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a fill, bulk copy or Across Worksheets action.',
  'fill-refused':
    'The editor declined a change to a fill, bulk copy or Across Worksheets action; the sentence says why. The cause fits no narrower kind.',
  'fill-unsupported':
    'The editor will not change a fill, bulk copy or Across Worksheets action this way: outside what it can change safely. The sentence says what.',

  // filter
  'filter-ambiguous': 'Stored content for a filter can be read in more than one way; the editor will not guess.',
  'filter-bad-input': 'The request about a filter needs different input from the person (a choice, name or range).',
  'filter-conflict': 'The change to a filter would collide with content that is already there.',
  'filter-duplicate': 'Something that must be unique appears twice for a filter.',
  'filter-limit': 'A size, count or time limit would be exceeded for a filter.',
  'filter-malformed': 'Stored content for a filter is not shaped as expected and cannot be read safely.',
  'filter-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a filter.',
  'filter-stale': 'The request about a filter refers to a state that has changed or expired; refresh and retry.',
  'filter-unsupported':
    'The editor will not change a filter this way: outside what it can change safely. The sentence says what.',

  // formula
  'formula-ambiguous':
    'Stored content for a formula or calculation can be read in more than one way; the editor will not guess.',
  'formula-bad-input':
    'The request about a formula or calculation needs different input from the person (a choice, name or range).',
  'formula-conflict': 'The change to a formula or calculation would collide with content that is already there.',
  'formula-duplicate': 'Something that must be unique appears twice for a formula or calculation.',
  'formula-limit': 'A size, count or time limit would be exceeded for a formula or calculation.',
  'formula-malformed':
    'Stored content for a formula or calculation is not shaped as expected and cannot be read safely.',
  'formula-not-found':
    'Something the request about a formula or calculation names (a sheet, item or record) does not exist.',
  'formula-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a formula or calculation.',
  'formula-refused':
    'The editor declined a change to a formula or calculation; the sentence says why. The cause fits no narrower kind.',
  'formula-stale':
    'The request about a formula or calculation refers to a state that has changed or expired; refresh and retry.',
  'formula-unsupported':
    'The editor will not change a formula or calculation this way: outside what it can change safely. The sentence says what.',

  // hyperlink
  'hyperlink-ambiguous': 'Stored content for a hyperlink can be read in more than one way; the editor will not guess.',
  'hyperlink-bad-input':
    'The request about a hyperlink needs different input from the person (a choice, name or range).',
  'hyperlink-conflict': 'The change to a hyperlink would collide with content that is already there.',
  'hyperlink-duplicate': 'Something that must be unique appears twice for a hyperlink.',
  'hyperlink-limit': 'A size, count or time limit would be exceeded for a hyperlink.',
  'hyperlink-malformed': 'Stored content for a hyperlink is not shaped as expected and cannot be read safely.',
  'hyperlink-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a hyperlink.',
  'hyperlink-refused':
    'The editor declined a change to a hyperlink; the sentence says why. The cause fits no narrower kind.',
  'hyperlink-stale': 'The request about a hyperlink refers to a state that has changed or expired; refresh and retry.',
  'hyperlink-unsupported':
    'The editor will not change a hyperlink this way: outside what it can change safely. The sentence says what.',

  // layout
  'layout-ambiguous':
    'Stored content for row, column or used-range layout can be read in more than one way; the editor will not guess.',
  'layout-bad-input':
    'The request about row, column or used-range layout needs different input from the person (a choice, name or range).',
  'layout-conflict': 'The change to row, column or used-range layout would collide with content that is already there.',
  'layout-duplicate': 'Something that must be unique appears twice for row, column or used-range layout.',
  'layout-limit': 'A size, count or time limit would be exceeded for row, column or used-range layout.',
  'layout-malformed':
    'Stored content for row, column or used-range layout is not shaped as expected and cannot be read safely.',
  'layout-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to row, column or used-range layout.',
  'layout-refused':
    'The editor declined a change to row, column or used-range layout; the sentence says why. The cause fits no narrower kind.',
  'layout-stale':
    'The request about row, column or used-range layout refers to a state that has changed or expired; refresh and retry.',
  'layout-unsupported':
    'The editor will not change row, column or used-range layout this way: outside what it can change safely. The sentence says what.',

  // macro
  'macro-protected':
    "The change would need the workbook's macro project to change, which the editor keeps untouched. Excel adds code modules for it.",

  // merge
  'merge-ambiguous': 'Stored content for merged cells can be read in more than one way; the editor will not guess.',
  'merge-bad-input': 'The request about merged cells needs different input from the person (a choice, name or range).',
  'merge-conflict': 'The change to merged cells would collide with content that is already there.',
  'merge-limit': 'A size, count or time limit would be exceeded for merged cells.',
  'merge-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to merged cells.',
  'merge-refused':
    'The editor declined a change to merged cells; the sentence says why. The cause fits no narrower kind.',
  'merge-unsupported':
    'The editor will not change merged cells this way: outside what it can change safely. The sentence says what.',

  // name
  'name-ambiguous': 'Stored content for a defined name can be read in more than one way; the editor will not guess.',
  'name-bad-input': 'The request about a defined name needs different input from the person (a choice, name or range).',
  'name-conflict': 'The change to a defined name would collide with content that is already there.',
  'name-duplicate': 'Something that must be unique appears twice for a defined name.',
  'name-limit': 'A size, count or time limit would be exceeded for a defined name.',
  'name-malformed': 'Stored content for a defined name is not shaped as expected and cannot be read safely.',
  'name-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a defined name.',
  'name-refused':
    'The editor declined a change to a defined name; the sentence says why. The cause fits no narrower kind.',
  'name-stale': 'The request about a defined name refers to a state that has changed or expired; refresh and retry.',
  'name-unsupported':
    'The editor will not change a defined name this way: outside what it can change safely. The sentence says what.',

  // note
  'note-ambiguous': 'Stored content for a note or comment can be read in more than one way; the editor will not guess.',
  'note-bad-input':
    'The request about a note or comment needs different input from the person (a choice, name or range).',
  'note-duplicate': 'Something that must be unique appears twice for a note or comment.',
  'note-limit': 'A size, count or time limit would be exceeded for a note or comment.',
  'note-malformed': 'Stored content for a note or comment is not shaped as expected and cannot be read safely.',
  'note-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a note or comment.',
  'note-refused':
    'The editor declined a change to a note or comment; the sentence says why. The cause fits no narrower kind.',
  'note-stale': 'The request about a note or comment refers to a state that has changed or expired; refresh and retry.',
  'note-unsupported':
    'The editor will not change a note or comment this way: outside what it can change safely. The sentence says what.',

  // object
  'object-ambiguous':
    'Stored content for a drawing, picture, shape, SmartArt, 3D or ink object can be read in more than one way; the editor will not guess.',
  'object-bad-input':
    'The request about a drawing, picture, shape, SmartArt, 3D or ink object needs different input from the person (a choice, name or range).',
  'object-conflict':
    'The change to a drawing, picture, shape, SmartArt, 3D or ink object would collide with content that is already there.',
  'object-limit':
    'A size, count or time limit would be exceeded for a drawing, picture, shape, SmartArt, 3D or ink object.',
  'object-malformed':
    'Stored content for a drawing, picture, shape, SmartArt, 3D or ink object is not shaped as expected and cannot be read safely.',
  'object-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a drawing, picture, shape, SmartArt, 3D or ink object.',
  'object-refused':
    'The editor declined a change to a drawing, picture, shape, SmartArt, 3D or ink object; the sentence says why. The cause fits no narrower kind.',
  'object-stale':
    'The request about a drawing, picture, shape, SmartArt, 3D or ink object refers to a state that has changed or expired; refresh and retry.',
  'object-unsupported':
    'The editor will not change a drawing, picture, shape, SmartArt, 3D or ink object this way: outside what it can change safely. The sentence says what.',

  // package
  'package-ambiguous':
    'Stored content for the workbook file package (parts, relationships, XML) can be read in more than one way; the editor will not guess.',
  'package-bad-input':
    'The request about the workbook file package (parts, relationships, XML) needs different input from the person (a choice, name or range).',
  'package-conflict':
    'The change to the workbook file package (parts, relationships, XML) would collide with content that is already there.',
  'package-duplicate':
    'Something that must be unique appears twice for the workbook file package (parts, relationships, XML).',
  'package-limit':
    'A size, count or time limit would be exceeded for the workbook file package (parts, relationships, XML).',
  'package-malformed':
    'Stored content for the workbook file package (parts, relationships, XML) is not shaped as expected and cannot be read safely.',
  'package-not-found':
    'Something the request about the workbook file package (parts, relationships, XML) names (a sheet, item or record) does not exist.',
  'package-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to the workbook file package (parts, relationships, XML).',
  'package-refused':
    'The editor declined a change to the workbook file package (parts, relationships, XML); the sentence says why. The cause fits no narrower kind.',
  'package-stale':
    'The request about the workbook file package (parts, relationships, XML) refers to a state that has changed or expired; refresh and retry.',
  'package-unsupported':
    'The editor will not change the workbook file package (parts, relationships, XML) this way: outside what it can change safely. The sentence says what.',

  // paste
  'paste-bad-input': 'The request about a paste needs different input from the person (a choice, name or range).',
  'paste-duplicate': 'Something that must be unique appears twice for a paste.',
  'paste-limit': 'A size, count or time limit would be exceeded for a paste.',
  'paste-malformed': 'Stored content for a paste is not shaped as expected and cannot be read safely.',
  'paste-refused': 'The editor declined a change to a paste; the sentence says why. The cause fits no narrower kind.',
  'paste-stale': 'The request about a paste refers to a state that has changed or expired; refresh and retry.',
  'paste-unsupported':
    'The editor will not change a paste this way: outside what it can change safely. The sentence says what.',

  // pivot
  'pivot-ambiguous': 'Stored content for a pivot table can be read in more than one way; the editor will not guess.',
  'pivot-bad-input': 'The request about a pivot table needs different input from the person (a choice, name or range).',
  'pivot-conflict': 'The change to a pivot table would collide with content that is already there.',
  'pivot-duplicate': 'Something that must be unique appears twice for a pivot table.',
  'pivot-limit': 'A size, count or time limit would be exceeded for a pivot table.',
  'pivot-malformed': 'Stored content for a pivot table is not shaped as expected and cannot be read safely.',
  'pivot-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a pivot table.',
  'pivot-refused':
    'The editor declined a change to a pivot table; the sentence says why. The cause fits no narrower kind.',
  'pivot-stale': 'The request about a pivot table refers to a state that has changed or expired; refresh and retry.',
  'pivot-unsupported':
    'The editor will not change a pivot table this way: outside what it can change safely. The sentence says what.',

  // print
  'print-ambiguous':
    'Stored content for print or page settings can be read in more than one way; the editor will not guess.',
  'print-bad-input':
    'The request about print or page settings needs different input from the person (a choice, name or range).',
  'print-conflict': 'The change to print or page settings would collide with content that is already there.',
  'print-duplicate': 'Something that must be unique appears twice for print or page settings.',
  'print-limit': 'A size, count or time limit would be exceeded for print or page settings.',
  'print-malformed': 'Stored content for print or page settings is not shaped as expected and cannot be read safely.',
  'print-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to print or page settings.',
  'print-refused':
    'The editor declined a change to print or page settings; the sentence says why. The cause fits no narrower kind.',
  'print-stale':
    'The request about print or page settings refers to a state that has changed or expired; refresh and retry.',
  'print-unsupported':
    'The editor will not change print or page settings this way: outside what it can change safely. The sentence says what.',

  // query
  'query-bad-input':
    'The request about a query, connection or saved action needs different input from the person (a choice, name or range).',
  'query-conflict':
    'The change to a query, connection or saved action would collide with content that is already there.',
  'query-duplicate': 'Something that must be unique appears twice for a query, connection or saved action.',
  'query-limit': 'A size, count or time limit would be exceeded for a query, connection or saved action.',
  'query-malformed':
    'Stored content for a query, connection or saved action is not shaped as expected and cannot be read safely.',
  'query-not-found':
    'Something the request about a query, connection or saved action names (a sheet, item or record) does not exist.',
  'query-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a query, connection or saved action.',
  'query-refused':
    'The editor declined a change to a query, connection or saved action; the sentence says why. The cause fits no narrower kind.',
  'query-stale':
    'The request about a query, connection or saved action refers to a state that has changed or expired; refresh and retry.',
  'query-unsupported':
    'The editor will not change a query, connection or saved action this way: outside what it can change safely. The sentence says what.',

  // rules
  'rules-ambiguous':
    'Stored content for conditional formatting or data validation can be read in more than one way; the editor will not guess.',
  'rules-bad-input':
    'The request about conditional formatting or data validation needs different input from the person (a choice, name or range).',
  'rules-conflict':
    'The change to conditional formatting or data validation would collide with content that is already there.',
  'rules-duplicate': 'Something that must be unique appears twice for conditional formatting or data validation.',
  'rules-limit': 'A size, count or time limit would be exceeded for conditional formatting or data validation.',
  'rules-malformed':
    'Stored content for conditional formatting or data validation is not shaped as expected and cannot be read safely.',
  'rules-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to conditional formatting or data validation.',
  'rules-refused':
    'The editor declined a change to conditional formatting or data validation; the sentence says why. The cause fits no narrower kind.',
  'rules-stale':
    'The request about conditional formatting or data validation refers to a state that has changed or expired; refresh and retry.',
  'rules-unsupported':
    'The editor will not change conditional formatting or data validation this way: outside what it can change safely. The sentence says what.',

  // session
  'session-bad-input':
    'The request about the session store (recovery, cache, checkpoints) needs different input from the person (a choice, name or range).',
  'session-duplicate':
    'Something that must be unique appears twice for the session store (recovery, cache, checkpoints).',
  'session-limit':
    'A size, count or time limit would be exceeded for the session store (recovery, cache, checkpoints).',
  'session-malformed':
    'Stored content for the session store (recovery, cache, checkpoints) is not shaped as expected and cannot be read safely.',
  'session-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to the session store (recovery, cache, checkpoints).',
  'session-refused':
    'The editor declined a change to the session store (recovery, cache, checkpoints); the sentence says why. The cause fits no narrower kind.',
  'session-stale':
    'The request about the session store (recovery, cache, checkpoints) refers to a state that has changed or expired; refresh and retry.',
  'session-unsupported':
    'The editor will not change the session store (recovery, cache, checkpoints) this way: outside what it can change safely. The sentence says what.',

  // sort
  'sort-bad-input': 'The request about a sort needs different input from the person (a choice, name or range).',
  'sort-duplicate': 'Something that must be unique appears twice for a sort.',
  'sort-unsupported':
    'The editor will not change a sort this way: outside what it can change safely. The sentence says what.',

  // spelling
  'spelling-bad-input':
    'The request about spell checking needs different input from the person (a choice, name or range).',
  'spelling-limit': 'A size, count or time limit would be exceeded for spell checking.',
  'spelling-refused':
    'The editor declined a change to spell checking; the sentence says why. The cause fits no narrower kind.',

  // spill
  'spill-ambiguous':
    'Stored content for a dynamic-array (spill) formula can be read in more than one way; the editor will not guess.',
  'spill-bad-input':
    'The request about a dynamic-array (spill) formula needs different input from the person (a choice, name or range).',
  'spill-conflict': 'The change to a dynamic-array (spill) formula would collide with content that is already there.',
  'spill-duplicate': 'Something that must be unique appears twice for a dynamic-array (spill) formula.',
  'spill-limit': 'A size, count or time limit would be exceeded for a dynamic-array (spill) formula.',
  'spill-malformed':
    'Stored content for a dynamic-array (spill) formula is not shaped as expected and cannot be read safely.',
  'spill-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a dynamic-array (spill) formula.',
  'spill-refused':
    'The editor declined a change to a dynamic-array (spill) formula; the sentence says why. The cause fits no narrower kind.',
  'spill-stale':
    'The request about a dynamic-array (spill) formula refers to a state that has changed or expired; refresh and retry.',
  'spill-unsupported':
    'The editor will not change a dynamic-array (spill) formula this way: outside what it can change safely. The sentence says what.',

  // style
  'style-ambiguous':
    'Stored content for a style, theme or number format can be read in more than one way; the editor will not guess.',
  'style-bad-input':
    'The request about a style, theme or number format needs different input from the person (a choice, name or range).',
  'style-conflict': 'The change to a style, theme or number format would collide with content that is already there.',
  'style-duplicate': 'Something that must be unique appears twice for a style, theme or number format.',
  'style-limit': 'A size, count or time limit would be exceeded for a style, theme or number format.',
  'style-malformed':
    'Stored content for a style, theme or number format is not shaped as expected and cannot be read safely.',
  'style-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to a style, theme or number format.',
  'style-refused':
    'The editor declined a change to a style, theme or number format; the sentence says why. The cause fits no narrower kind.',
  'style-stale':
    'The request about a style, theme or number format refers to a state that has changed or expired; refresh and retry.',
  'style-unsupported':
    'The editor will not change a style, theme or number format this way: outside what it can change safely. The sentence says what.',

  // table
  'table-ambiguous': 'Stored content for a table can be read in more than one way; the editor will not guess.',
  'table-bad-input': 'The request about a table needs different input from the person (a choice, name or range).',
  'table-conflict': 'The change to a table would collide with content that is already there.',
  'table-duplicate': 'Something that must be unique appears twice for a table.',
  'table-limit': 'A size, count or time limit would be exceeded for a table.',
  'table-malformed': 'Stored content for a table is not shaped as expected and cannot be read safely.',
  'table-protected': 'Protection, or content kept read-only to avoid loss, blocks the change to a table.',
  'table-refused': 'The editor declined a change to a table; the sentence says why. The cause fits no narrower kind.',
  'table-stale': 'The request about a table refers to a state that has changed or expired; refresh and retry.',
  'table-unsupported':
    'The editor will not change a table this way: outside what it can change safely. The sentence says what.',

  // workbook
  'workbook-ambiguous':
    'Stored content for the workbook as a whole can be read in more than one way; the editor will not guess.',
  'workbook-bad-input':
    'The request about the workbook as a whole needs different input from the person (a choice, name or range).',
  'workbook-conflict': 'The change to the workbook as a whole would collide with content that is already there.',
  'workbook-duplicate': 'Something that must be unique appears twice for the workbook as a whole.',
  'workbook-limit': 'A size, count or time limit would be exceeded for the workbook as a whole.',
  'workbook-malformed':
    'Stored content for the workbook as a whole is not shaped as expected and cannot be read safely.',
  'workbook-not-found':
    'Something the request about the workbook as a whole names (a sheet, item or record) does not exist.',
  'workbook-protected':
    'Protection, or content kept read-only to avoid loss, blocks the change to the workbook as a whole.',
  'workbook-refused':
    'The editor declined a change to the workbook as a whole; the sentence says why. The cause fits no narrower kind.',
  'workbook-stale':
    'The request about the workbook as a whole refers to a state that has changed or expired; refresh and retry.',
  'workbook-unsupported':
    'The editor will not change the workbook as a whole this way: outside what it can change safely. The sentence says what.',
});

/** An Error with a stable code. */
export function refusal(code, message) {
  if (!Object.hasOwn(REFUSAL_CODES, code)) throw Error(`Unknown refusal code: ${code}`);
  return Object.assign(Error(message), { code });
}

/** True for one of our refusals (its code is in this registry), false for a system error such as ENOENT. */
export const isRefusal = (error) => Object.hasOwn(REFUSAL_CODES, error?.code);

/** Re-word an error while keeping its code, for layers that add context. */
export function reworded(error, message, code = error?.code) {
  return Object.assign(Error(message), code ? { code } : {});
}
