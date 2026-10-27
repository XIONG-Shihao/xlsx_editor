/** Runs in a calculation worker. Fetch only the formula/reference projection
 * selected by the existing evaluator; write results back to the native store. */
import { calculationInputRanges, calculationDependencies, calculateWorkbook } from './calculation-node.mjs';
import { nativeFormula } from './native-formula.mjs';
import { createHash } from 'node:crypto';
import { columnName, cellPosition, rangeBounds } from './workbook-commands.mjs';
import { indexedCalendarGuards } from './indexed-metadata.mjs';
import { hasArrayContext, hasUnqualifiedArrayResults } from './array-calculation.mjs';
import { refusal } from './refusal-codes.mjs';

// Merge only adjacent selected cells. A bounding box or fixed-size tile could
// pull unrelated formulas into the compatibility projection and exhaust its
// budget. Exact rectangles also avoid one worker/native round trip per formula.
function* selectedFormulaRanges(selected) {
  const sheets = new Map();
  for (const key of selected) {
    const [name, ref] = key.split('\0'),
      p = cellPosition(ref);
    if (!sheets.has(name)) sheets.set(name, []);
    sheets.get(name).push(p.row * 16384 + p.col);
  }
  for (const [name, positions] of sheets) {
    positions.sort((a, b) => a - b);
    let active = new Map(),
      previousRow = -2;
    const range = (r) => ({ name, range: `${columnName(r.first)}${r.top + 1}:${columnName(r.last)}${r.bottom + 1}` });
    for (let i = 0; i < positions.length;) {
      const row = Math.floor(positions[i] / 16384),
        runs = [];
      while (i < positions.length && Math.floor(positions[i] / 16384) === row) {
        const col = positions[i++] % 16384,
          last = runs.at(-1);
        if (last && col === last.last + 1) last.last = col;
        else runs.push({ first: col, last: col, top: row, bottom: row });
      }
      if (row !== previousRow + 1) {
        for (const r of active.values()) yield range(r);
        active.clear();
      }
      const next = new Map();
      for (const run of runs) {
        const key = `${run.first}:${run.last}`,
          prior = active.get(key);
        if (prior) {
          prior.bottom = row;
          next.set(key, prior);
          active.delete(key);
        } else next.set(key, run);
      }
      for (const r of active.values()) yield range(r);
      active = next;
      previousRow = row;
    }
    for (const r of active.values()) yield range(r);
  }
}
export async function calculateIndexedWorkbook(input, native) {
  indexedCalendarGuards(input.view);
  try {
    return await calculateIndexed(input, native);
  } catch (error) {
    if (
      input.changedRanges &&
      /Unreliable mixed-calculation predecessor|Unqualified array context requires complete calculation/.test(
        String(error)
      )
    )
      return calculateIndexed({ ...input, changedRanges: undefined }, native);
    throw error;
  }
}
async function calculateIndexed(input, native) {
  if (['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'].includes(input.command?.operation))
    return (await import('./indexed-insertion.mjs')).prepareIndexedInsertion(input, native);
  const { root, revision } = input;
  let { version } = input;
  let partial;
  let selected;
  const projected = new Set();
  let projectedBytes = 0,
    selectionBytes = 0;
  // Fail before publication rather than constructing an unbounded compatibility graph.
  const maxProjectionBytes = input.maxProjectionBytes ?? 64 * 1024 * 1024;
  const view = { ...input.view, revision, sheets: input.view.sheets.map((sheet) => ({ ...sheet, cells: {} })) };
  let inputCells = 0,
    formulaCells = 0,
    rangeReads = 0;
  const ownedSpills = new Map();
  const ownerKey = (sheet, ref) => `${sheet.path}\0${ref}`;
  function spillInput(sheet, ref, cell) {
    const owner = ownedSpills.get(ownerKey(sheet, cell.spill?.master));
    if (!owner || cell.spill?.schema !== 1 || cell.spill.range !== owner.range)
      throw refusal('formula-bad-input', 'Spill input has no complete native ownership proof');
    // Typed blocked masters keep their authored protection in the evaluator.
    // Numeric outputs alone become literal predecessors. The native group,
    // formula, and complete geometry stay authoritative in persistent storage.
    if (owner.status === 'error') return cell;
    if (
      cell.calculation?.status !== 'calculated' ||
      typeof cell.calculation.value !== 'number' ||
      !Number.isFinite(cell.calculation.value)
    )
      throw refusal('formula-unsupported', 'Spill input has no qualified finite numeric result');
    return { value: cell.calculation.value, dataType: 'number', financialInputBlocked: true };
  }
  async function fetch(sheet, range, formulas = false, include) {
    let offset = 0;
    do {
      rangeReads++;
      const result = await native({
        action: 'workbook',
        operation: 'range',
        calculationInput: true,
        root,
        version,
        sheet: sheet.path,
        range,
        offset,
        limit: 2048,
        formulas,
      });
      for (const [ref, cell] of result.cells) {
        if (include && !(typeof include === 'function' ? include(ref) : include.has(ref))) continue;
        for (const key of ['value', 'cachedValue'])
          if (cell[key] && typeof cell[key] === 'object' && '$workbookNumber' in cell[key])
            cell[key] = Number(cell[key].$workbookNumber);
        if (cell.spill) sheet.cells[ref] = spillInput(sheet, ref, cell);
        else if (partial && cell.calculation?.nativePartial && cell.calculation.revision === revision) {
          if (formulas) continue;
          {
            const value = cell.calculation.value;
            sheet.cells[ref] = {
              value,
              dataType: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text',
            };
          }
        } else if (selected && (cell.formula || cell.formulaType) && !selected.has(`${sheet.name}\0${ref}`)) {
          if (cell.calculation?.status !== 'calculated')
            throw refusal('formula-unsupported', 'Unreliable mixed-calculation predecessor');
          const value = cell.calculation.value ?? null;
          sheet.cells[ref] = {
            value,
            dataType: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'text',
          };
        } else sheet.cells[ref] = cell;
        const key = `${sheet.name}\0${ref}`;
        if (!projected.has(key)) {
          projected.add(key);
          projectedBytes += Buffer.byteLength(JSON.stringify(sheet.cells[ref])) + ref.length + 64;
          if (projectedBytes + selectionBytes > maxProjectionBytes)
            throw refusal(
              'formula-limit',
              'Compatibility calculation input exceeds the 64 MiB projection budget; accepted workbook is unchanged'
            );
        }
      }
      offset = result.nextOffset;
    } while (offset !== null);
  }
  formulaCells = view.sheets.reduce((n, s) => n + (s.storedFormulaCount ?? 0), 0);
  if (!formulaCells) return { version, inputCells: 0, formulaCells: 0 };
  const staged = input.calculationSheet
    ? { version, owners: [] }
    : await native({ action: 'workbook', operation: 'spill-stage', root, version, revision, view });
  version = staged.version;
  for (const owner of staged.owners ?? []) {
    const sheet = view.sheets.find((sheet) => sheet.path === owner.sheet);
    if (!sheet) throw refusal('formula-unsupported', 'Spill owner worksheet is unavailable');
    ownedSpills.set(ownerKey(sheet, owner.master), owner);
    const bounds = rangeBounds(owner.range);
    sheet.rows = Math.max(sheet.rows, bounds.last.row + 1);
    sheet.columns = Math.max(sheet.columns, bounds.last.col + 1);
  }
  for (const owner of ownedSpills.values())
    await fetch(
      view.sheets.find((sheet) => sheet.path === owner.sheet),
      owner.range
    );
  const metadata = createHash('sha256')
    .update(
      JSON.stringify({
        sheets: view.sheets.map((s) => ({ name: s.name, path: s.path, date1904: !!s.date1904 })),
        names: view.namedRanges,
        tables: view.tables,
      })
    )
    .digest('hex');
  const dependency = (method, extra = {}) =>
    native({ action: 'workbook', operation: 'dependencies', root, method, version, ...extra });
  // Spill constructor/output dependencies have no reusable static-plan model
  // yet. Recompute the bounded groups and scalar dependents on every edit.
  const removed =
    !ownedSpills.size && input.changedRanges && input.formulaChanged
      ? await dependency('update-formulas', { metadata, changedRanges: input.changedRanges, view })
      : null;
  const match = input.calculationPlan
    ? { hit: true, plan: input.calculationPlan, count: 1 }
    : !ownedSpills.size && input.changedRanges && (!input.formulaChanged || removed?.hit)
      ? await dependency('select', { metadata, changedRanges: input.changedRanges })
      : { hit: false };
  if (match.hit && !match.count)
    return {
      version,
      inputCells: 0,
      formulaCells,
      rangeReads: 0,
      evaluatedFormulaCells: 0,
      mode: 'unaffected-static-inputs',
      skipped: true,
      nativePlanHit: true,
    };
  const stored = ownedSpills.size
    ? { supported: false }
    : await native({
        action: 'workbook',
        operation: 'calculate',
        root,
        version,
        revision,
        view,
        metadata,
        ...(match.hit ? { plan: match.plan } : {}),
      });
  if (stored.supported) {
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root,
      name: 'dependency-plan.json',
      data: JSON.stringify({ scope: { mode: 'static-references' } }),
    });
    return { ...stored, formulaCells };
  }
  // Prove a complete graph independently of evaluator support. Page every
  // formula once, including native formulas that will not enter the JS evaluator.
  let completePlan = match.hit ? match.plan : null,
    arrayContext = false;
  if (!match.hit && !ownedSpills.size) {
    completePlan = (await dependency('create', { metadata })).plan;
    for (const sheet of view.sheets) {
      let after;
      do {
        const page = await native({
          action: 'workbook',
          operation: 'range',
          calculationInput: true,
          root,
          version,
          sheet: sheet.path,
          range: 'A1:XFD1048576',
          after,
          limit: 512,
          formulas: true,
        });
        const projection = {
          ...view,
          sheets: view.sheets.map((s) => ({
            ...s,
            cells: s.name === sheet.name ? Object.fromEntries(page.cells) : {},
          })),
        };
        arrayContext ||= hasArrayContext(projection);
        const parsed = await nativeFormula(native, root, 'analyze', { view: projection });
        const items =
          parsed.supported && !hasArrayContext(projection)
            ? parsed.items
            : calculationDependencies(projection, { conservative: true });
        if (!items) {
          completePlan = null;
          break;
        }
        for (let start = 0; start < items.length; start += 128)
          await dependency('append', { plan: completePlan, items: items.slice(start, start + 128) });
        after = page.nextCursor;
      } while (after);
      if (completePlan === null) break;
    }
  }
  // Partial native results cannot prove a read from an unowned spill follower:
  // a missing cell there is not a qualified blank input for compatibility.
  if (!input.skipNativePartial && !arrayContext && !ownedSpills.size) {
    const candidate = await native({
      action: 'workbook',
      operation: 'calculate',
      root,
      version,
      revision,
      view,
      metadata,
      partial: true,
      tracePlan: completePlan,
      ...(match.hit ? { plan: match.plan } : {}),
    });
    if (candidate.supported) {
      partial = candidate;
      version = partial.version;
    }
  }
  let scope,
    plan = partial?.runtimePlan ?? completePlan,
    nativePlanHit = false,
    nativeParser = false,
    nativeEvaluator = false;
  const intersects = (change, range) => {
    if (change.sheet !== range.sheet) return false;
    const a = rangeBounds(change.range),
      b = range.ref
        ? (() => {
            const p = cellPosition(range.ref);
            return { startRow: p.row, endRow: p.row, startColumn: p.col, endColumn: p.col };
          })()
        : range;
    return (
      a.first.row <= b.endRow && a.last.row >= b.startRow && a.first.col <= b.endColumn && a.last.col >= b.startColumn
    );
  };
  if (match.hit) {
    nativePlanHit = true;
    selected = new Set();
    let offset = 0;
    do {
      const page = await dependency('page', { plan: match.plan, kind: 'formulas', offset });
      for (const item of page.items) {
        const key = `${item.sheet}\0${item.ref}`;
        if (!selected.has(key)) {
          selectionBytes += Buffer.byteLength(key) + 64;
          if (selectionBytes > maxProjectionBytes)
            throw refusal(
              'formula-limit',
              'Compatibility formula selection exceeds the 64 MiB projection budget; accepted workbook is unchanged'
            );
          selected.add(key);
        }
      }
      offset = page.nextOffset;
    } while (offset !== null);
    for (const { name, range } of selectedFormulaRanges(selected))
      await fetch(
        view.sheets.find((s) => s.name === name),
        range,
        true
      );
    if (hasArrayContext(view))
      throw refusal('formula-unsupported', 'Unqualified array context requires complete calculation');
    const parsed = await nativeFormula(native, root, 'analyze', { view });
    nativeParser = parsed.supported;
    scope = parsed.supported && !hasArrayContext(view) ? parsed.scope : calculationInputRanges(view);
  }
  if (!scope) {
    for (const sheet of view.sheets) if (sheet.storedFormulaCount) await fetch(sheet, 'A1:XFD1048576', true);
    const parsed = await nativeFormula(native, root, 'analyze', { view });
    nativeParser = parsed.supported;
    scope = parsed.supported && !hasArrayContext(view) ? parsed.scope : calculationInputRanges(view);
  }
  // Only this tiny cache hint stays JSON; formulas, edges and closure are SQLite.
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root,
    name: 'dependency-plan.json',
    data: JSON.stringify({ scope: { mode: scope.mode } }),
  });
  const ranges =
    scope.mode === 'full'
      ? view.sheets.map((sheet) => ({
          sheet: sheet.name,
          startRow: 0,
          endRow: sheet.rows - 1,
          startColumn: 0,
          endColumn: sheet.columns - 1,
        }))
      : scope.ranges;
  const exactTiles = new Map();
  for (const range of ranges)
    if (range.ref) {
      const p = cellPosition(range.ref),
        row = Math.floor(p.row / 32) * 32,
        col = Math.floor(p.col / 32) * 32,
        key = `${range.sheet}\0${row}\0${col}`;
      if (!exactTiles.has(key)) exactTiles.set(key, { sheet: range.sheet, row, col, refs: new Set() });
      exactTiles.get(key).refs.add(range.ref);
    }
  for (const tile of exactTiles.values()) {
    const sheet = view.sheets.find((sheet) => sheet.name === tile.sheet);
    if (!sheet) throw refusal('formula-unsupported', 'Calculation input sheet is unavailable');
    await fetch(
      sheet,
      `${columnName(tile.col)}${tile.row + 1}:${columnName(Math.min(16383, tile.col + 31))}${Math.min(1048576, tile.row + 32)}`,
      false,
      tile.refs
    );
  }
  // Coalesce overlapping input ranges by column span, then stream row bands.
  // A prefix-SUM workbook must not fetch A1:A1000 again for every formula.
  for (const sheet of view.sheets) {
    const groups = new Map();
    for (const range of ranges)
      if (!range.ref && range.sheet === sheet.name) {
        const r = {
          ...range,
          endRow: Math.min(range.endRow, sheet.rows - 1),
          endColumn: Math.min(range.endColumn, sheet.columns - 1),
        };
        if (r.startRow > r.endRow || r.startColumn > r.endColumn) continue;
        const key = `${r.startColumn}:${r.endColumn}`;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(r);
      }
    const merged = [];
    for (const group of groups.values()) {
      group.sort((a, b) => a.startRow - b.startRow);
      let last;
      for (const range of group) {
        if (last && range.startRow <= last.endRow + 1) last.endRow = Math.max(last.endRow, range.endRow);
        else {
          last = { ...range };
          merged.push(last);
        }
      }
    }
    if (!merged.length) continue;
    const first = merged.reduce((a, r) => Math.min(a, r.startRow), Infinity),
      last = merged.reduce((a, r) => Math.max(a, r.endRow), 0);
    for (let row = Math.floor(first / 128) * 128; row <= last; row += 128) {
      const active = merged.filter((r) => r.startRow <= row + 127 && r.endRow >= row);
      const columns = [];
      for (const r of active.toSorted((a, b) => a.startColumn - b.startColumn)) {
        const previous = columns.at(-1);
        if (previous && r.startColumn <= previous.end + 1) previous.end = Math.max(previous.end, r.endColumn);
        else columns.push({ start: r.startColumn, end: r.endColumn });
      }
      const include = (ref) => {
        const p = cellPosition(ref);
        return active.some(
          (r) => p.row >= r.startRow && p.row <= r.endRow && p.col >= r.startColumn && p.col <= r.endColumn
        );
      };
      for (const interval of columns)
        for (let col = interval.start; col <= interval.end; col += 64)
          await fetch(
            sheet,
            `${columnName(col)}${row + 1}:${columnName(Math.min(interval.end, col + 63))}${Math.min(row + 128, sheet.rows)}`,
            false,
            include
          );
    }
  }
  if (selected)
    for (const sheet of view.sheets)
      for (const [ref, cell] of Object.entries(sheet.cells))
        if (cell.formula && !selected.has(`${sheet.name}\0${ref}`)) {
          if (!scope.ranges.some((range) => intersects({ sheet: sheet.name, range: ref }, range))) {
            delete sheet.cells[ref];
            continue;
          }
          sheet.cells[ref] = {
            value: cell.calculation.value ?? null,
            dataType:
              typeof cell.calculation.value === 'number'
                ? 'number'
                : typeof cell.calculation.value === 'boolean'
                  ? 'boolean'
                  : 'text',
          };
        }
  for (const sheet of view.sheets) inputCells += Object.keys(sheet.cells).length;
  // Only a native complete-output proof may turn these authored array contexts
  // into scalar predecessors. Keep original formula/source metadata for cache
  // publication; neither portable syntax acceptance nor one cached value suffices.
  const spillConsumers = new Map();
  if (ownedSpills.size) {
    const candidates = [];
    for (const sheet of view.sheets)
      for (const [ref, cell] of Object.entries(sheet.cells))
        if (cell.formula && /#|ANCHORARRAY/i.test(cell.formula) && !ownedSpills.has(ownerKey(sheet, ref)))
          candidates.push({ sheet: sheet.path, ref });
    for (let offset = 0; offset < candidates.length; offset += 64) {
      const proof = await native({
        action: 'workbook',
        operation: 'spill-consumers',
        root,
        version,
        revision,
        view: { ...view, sheets: view.sheets.map((sheet) => ({ ...sheet, cells: {} })) },
        items: candidates.slice(offset, offset + 64),
      });
      if (proof.version !== version)
        throw refusal('formula-unsupported', 'Scalar spill proof belongs to a different accepted version');
      for (const result of proof.results) {
        const sheet = view.sheets.find((sheet) => sheet.path === result.sheet),
          original = sheet?.cells[result.ref];
        if (
          !original ||
          original.formula !== result.formula ||
          result.calculation?.status !== 'calculated' ||
          typeof result.calculation.value !== 'number' ||
          !Number.isFinite(result.calculation.value) ||
          result.calculation.spillConsumer?.schema !== 1
        )
          throw refusal('internal-invalid-field', 'Invalid native scalar spill proof');
        projectedBytes += Buffer.byteLength(JSON.stringify(result.calculation)) + 64;
        if (projectedBytes + selectionBytes > maxProjectionBytes)
          throw refusal(
            'formula-limit',
            'Scalar spill calculation exceeds the 64 MiB projection budget; accepted workbook is unchanged'
          );
        spillConsumers.set(ownerKey(sheet, result.ref), {
          sheet: sheet.path,
          cell: result.ref,
          value: { ...original, calculation: result.calculation },
        });
        sheet.cells[result.ref] = { value: result.calculation.value, dataType: 'number' };
      }
    }
  }
  // These roots return one typed scalar from a complete bounded transformation.
  // The native proof reads accepted literal inputs; projection is private and
  // publication restores the original authored formula and singleton metadata.
  const scalarArrays = new Map(),
    scalarRefusals = new Map(),
    scalarCandidates = [];
  for (const sheet of view.sheets)
    for (const [ref, cell] of Object.entries(sheet.cells))
      if (cell.formula && /\b(?:_xlfn\.)?(?:XLOOKUP|SEQUENCE|UNIQUE)\s*\(/i.test(cell.formula))
        scalarCandidates.push({ sheet: sheet.path, ref });
  for (let offset = 0; offset < scalarCandidates.length; offset += 64) {
    const proof = await native({
      action: 'workbook',
      operation: 'scalar-proofs',
      root,
      version,
      revision,
      items: scalarCandidates.slice(offset, offset + 64),
    });
    if (proof.version !== version || !Array.isArray(proof.results))
      throw refusal('formula-unsupported', 'Scalar array proof belongs to a different accepted version');
    for (const refusal of proof.refused ?? []) {
      const sheet = view.sheets.find((sheet) => sheet.path === refusal.sheet),
        original = sheet?.cells[refusal.ref];
      if (!original?.formula || typeof refusal.reason !== 'string')
        throw refusal('internal-invalid-field', 'Invalid native scalar array refusal');
      scalarRefusals.set(ownerKey(sheet, refusal.ref), {
        sheet: sheet.path,
        cell: refusal.ref,
        value: original,
        reason: refusal.reason,
      });
      // This blocked producer supplies the existing complete follower/dependent
      // protection policy while retaining the authored source in publication.
      sheet.cells[refusal.ref] = { ...original, formula: 'SEQUENCE(-1)' };
    }
    for (const result of proof.results) {
      const sheet = view.sheets.find((sheet) => sheet.path === result.sheet),
        original = sheet?.cells[result.ref],
        calculation = result.calculation,
        type = calculation?.scalarArray?.valueType;
      const typed =
        type === 'error'
          ? calculation?.status === 'error' && calculation.error === '#N/A'
          : calculation?.status === 'calculated' &&
            ['number', 'text', 'boolean'].includes(type) &&
            typeof calculation.value === { number: 'number', text: 'string', boolean: 'boolean' }[type] &&
            (type !== 'number' || Number.isFinite(calculation.value));
      if (!original || original.formula !== result.formula || calculation?.scalarArray?.schema !== 1 || !typed)
        throw refusal('internal-invalid-field', 'Invalid native scalar array proof');
      projectedBytes += Buffer.byteLength(JSON.stringify(calculation)) + 64;
      if (projectedBytes + selectionBytes > maxProjectionBytes)
        throw refusal(
          'formula-limit',
          'Scalar array calculation exceeds the 64 MiB projection budget; accepted workbook is unchanged'
        );
      scalarArrays.set(ownerKey(sheet, result.ref), {
        sheet: sheet.path,
        cell: result.ref,
        value: { ...original, calculation },
      });
      sheet.cells[result.ref] =
        type === 'error'
          ? { value: calculation.error, dataType: 'error', error: calculation.error }
          : { value: calculation.value, dataType: type };
    }
  }
  const evaluated = await nativeFormula(native, root, 'evaluate', { view });
  nativeEvaluator = evaluated.supported;
  // Native-independent input admission also governs compatibility: a source
  // profile miss must never become a guessed fallback cache (including IFERROR).
  const financialRefusals = new Map(),
    financialCandidates = [];
  for (const sheet of view.sheets)
    for (const [ref, cell] of Object.entries(sheet.cells))
      if (cell.formula && /\b(?:_xlfn\.)?(?:PMT|PV|FV)\s*\(/i.test(cell.formula))
        financialCandidates.push({ sheet: sheet.path, ref });
  for (let offset = 0; offset < financialCandidates.length; offset += 128) {
    const proof = await native({
      action: 'workbook',
      operation: 'financial-profiles',
      root,
      version,
      view: { ...view, sheets: view.sheets.map((sheet) => ({ ...sheet, cells: {} })) },
      items: financialCandidates.slice(offset, offset + 128),
    });
    if (proof.version !== version || proof.results.length !== Math.min(128, financialCandidates.length - offset))
      throw refusal('internal-invalid-field', 'Invalid financial profile receipt');
    for (const [index, result] of proof.results.entries()) {
      const expected = financialCandidates[offset + index];
      const sheet = view.sheets.find((s) => s.path === result.sheet),
        cell = sheet?.cells[result.ref];
      if (
        result.sheet !== expected.sheet ||
        result.ref !== expected.ref ||
        !cell ||
        cell.formula !== result.formula ||
        typeof result.supported !== 'boolean' ||
        (!result.supported && (typeof result.reason !== 'string' || !result.reason))
      )
        throw refusal('formula-stale', 'Financial input proof belongs to another formula');
      if (!result.supported) {
        financialRefusals.set(`${sheet.name}\0${result.ref}`, { formula: cell.formula, reason: result.reason });
        cell.formula = 'PLATFORM_UNSUPPORTED_FINANCIAL_PROFILE()';
      }
    }
  }
  let calculated;
  if (evaluated.supported && !financialRefusals.size) {
    calculated = view;
    for (const item of evaluated.results) {
      if (
        item.error !== undefined &&
        !['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!'].includes(item.error)
      )
        throw refusal('internal-invalid-field', 'Invalid native typed formula error');
      view.sheets.find((s) => s.name === item.sheet).cells[item.ref].calculation =
        item.error !== undefined
          ? { status: 'error', error: item.error, revision }
          : { status: 'calculated', value: item.value, revision };
    }
  } else calculated = await calculateWorkbook(view, { takeOwnership: true });
  for (const refusal of scalarRefusals.values()) {
    const sheet = calculated.sheets.find((sheet) => sheet.path === refusal.sheet);
    sheet.cells[refusal.cell] = {
      ...refusal.value,
      calculation: {
        status: 'unresolved',
        error: 'Array/spill calculation requires qualified output ownership: ' + refusal.reason,
        revision,
      },
    };
  }
  // A result-shape proof can discover an array producer outside the authored
  // catalog. Rebuild with every formula visible before any partial scalar
  // result can certify a blank or stale follower read.
  if (partial && hasUnqualifiedArrayResults(calculated))
    return calculateIndexedWorkbook({ ...input, changedRanges: undefined, skipNativePartial: true }, native);
  for (const sheet of calculated.sheets)
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const refusal = financialRefusals.get(`${sheet.name}\0${ref}`);
      if (refusal) {
        cell.formula = refusal.formula;
        cell.calculation = { status: 'unresolved', error: refusal.reason, revision };
      }
    }
  // A workbook-wide engine initialization failure is not a per-formula miss.
  // Preserve the established unresolved state instead of hiding it behind a
  // partially calculated workbook (for example duplicate local-name loading).
  if (
    partial &&
    calculated.sheets.some((s) =>
      Object.values(s.cells).some((c) => c.calculation?.error?.startsWith('Calculation unavailable:'))
    )
  )
    return calculateIndexedWorkbook({ ...input, skipNativePartial: true }, native);
  let accepted = version,
    changes = [];
  // Results are calculation updates, not authored formula edits. Empty intent
  // preserves any real edit inherited from the accepted version in native SQL.
  const flush = async () => {
    if (!changes.length) return;
    const result = await native({
      action: 'workbook',
      operation: 'patch',
      root,
      version: accepted,
      changes: JSON.parse(
        JSON.stringify(
          changes.map((change) => ({ ...change, edits: [] })),
          (_key, item) =>
            typeof item === 'number' && (!Number.isFinite(item) || Object.is(item, -0))
              ? { $workbookNumber: Object.is(item, -0) ? '-0' : String(item) }
              : item
        )
      ),
    });
    accepted = result.version;
    changes = [];
  };
  for (const sheet of calculated.sheets)
    for (const [ref, cell] of Object.entries(sheet.cells))
      if ((cell.formula || cell.formulaType) && !ownedSpills.has(ownerKey(sheet, ref))) {
        changes.push({ sheet: sheet.path, cell: ref, value: cell });
        if (changes.length === 1024) await flush();
      }
  for (const change of [...spillConsumers.values(), ...scalarArrays.values()]) {
    changes.push(change);
    if (changes.length === 1024) await flush();
  }
  await flush();
  // Dependency completeness does not depend on evaluation success. Unknown
  // contexts carry conservative edges; errors never certify an input value.
  if (hasUnqualifiedArrayResults(calculated)) plan = null;
  await dependency(plan ? 'bind' : 'invalidate', { plan, version: accepted });
  if (
    calculated.sheets.some((s) =>
      Object.values(s.cells).some((c) => (c.formula || c.formulaType) && c.calculation?.status !== 'calculated')
    )
  )
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root,
      name: 'dependency-plan.json',
      data: JSON.stringify({ scope: { mode: 'full' } }),
    });
  return {
    nativePartialCells: partial?.evaluatedFormulaCells ?? 0,
    spillConsumerCells: spillConsumers.size,
    scalarArrayCells: scalarArrays.size,
    projectedBytes,
    nativePlanHit,
    nativeParser,
    nativeEvaluator,
    nativeFallbackReason: evaluated.reason,
    version: accepted,
    inputCells,
    formulaCells,
    rangeReads,
    evaluatedFormulaCells: selected?.size ?? formulaCells,
    mode: selected ? 'affected-static-formulas' : scope.mode,
    reason: scope.reason,
    ...(ownedSpills.size
      ? {
          spillDimensions: view.sheets.map((sheet) => ({ path: sheet.path, rows: sheet.rows, columns: sheet.columns })),
        }
      : {}),
  };
}
