import './support/scratch-dir.mjs';
/** Destination review source tests. Run
 * after the replacement API lands. Controlled owner, no native/app claim. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash, randomUUID } from 'node:crypto';
import { rangeBounds } from '../core/workbook-commands.mjs';
import { createControlledQueryOwner } from './support/local-query-controlled-owner.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  output = await fs.mkdtemp(path.join(component, '.build/tests/query-replacement-checks-'));
const files = [
    'core/local-query.mjs',
    'core/local-query.d.mts',
    'adapters/local-query-host.mjs',
    'adapters/local-query-connections.mjs',
    'adapters/local-query-inspection.mjs',
    'tests/local-query-replacement.mjs',
    'tests/support/local-query-controlled-owner.mjs',
  ],
  pins = async () =>
    Promise.all(
      files.map(async (relative) => {
        const p = path.join(component, relative),
          b = await fs.readFile(p);
        return { relative, path: p, sha256: createHash('sha256').update(b).digest('hex'), bytes: b.length };
      })
    ),
  before = await pins();
const check = (name, fn) => {
  test(name, { timeout: 30000 }, async (t) => {
    await fn(t);
  });
};
const preview = async (f, revision = 1, options = {}) =>
  (await f.host().preview(f.record.id, revision, options)).queryPreview;
const reviewed = (p) => ({
  expectedConversionRevision: p.conversionInspection.revision,
  expectedReplacementRevision: p.replacementInspection.revision,
});
const valueCell = (value, extra = {}) => ({
  value,
  dataType:
    typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : value === null ? 'blank' : 'text',
  rawValue: value === null ? null : String(value),
  ...extra,
});
check(
  'first occupied target requires exact acknowledgment, without writes on opening or missing/stale approval',
  async (t) => {
    const f = await createControlledQueryOwner(t, output);
    f.cells.B3 = valueCell(5, { format: 2, bold: true, styleId: 8 });
    const before = structuredClone(f.cells),
      p = await preview(f),
      r = p.replacementInspection;
    assert.equal(r.occupied, 1);
    assert.equal(r.blocked, 0);
    assert.equal(r.rows[0].sourceRef, 'A1');
    assert.equal(r.rows[0].destinationRef, 'B3');
    assert.equal(r.rows[0].sourceText, '1');
    assert.equal(r.rows[0].current.display, '5');
    assert.deepEqual(f.cells, before);
    assert.equal(f.accepted(), 0);
    await assert.rejects(
      f.host().refresh(f.record.id, 1, { expectedConversionRevision: p.conversionInspection.revision }),
      /replace|acknowledg|review/i
    );
    await assert.rejects(
      f.host().refresh(f.record.id, 1, { ...reviewed(p), expectedReplacementRevision: randomUUID() }),
      /replace|acknowledg|review/i
    );
    assert.equal(f.accepted(), 0);
    assert.deepEqual(f.cells, before);
    assert.equal(f.requests.filter((r) => r.operation === 'importBegin').length, 0);
    await f.host().refresh(f.record.id, 1, reviewed(p));
    assert.equal(f.accepted(), 1);
    assert.equal(f.cells.B3.value, 1);
    assert.equal(f.cells.B3.bold, true);
  }
);
check(
  'all original values and source/destination refs beyond20 rows remain accessible through exact replacement pages',
  async (t) => {
    const rows = 205,
      f = await createControlledQueryOwner(t, output, {
        text: Array.from({ length: rows }, (_, i) => String(i + 1)).join('\n'),
      });
    for (let i = 0; i < rows; i++) f.cells['B' + (i + 3)] = valueCell('Original ' + i);
    const p = await preview(f),
      r = p.replacementInspection;
    assert.equal(r.occupied, rows);
    assert.equal(r.rows.length, 100);
    assert.equal(r.nextOffset, 100);
    let next = r.nextOffset,
      all = [...r.rows];
    while (next !== null) {
      const page = (
        await preview(f, 1, { expectedConversionRevision: p.conversionInspection.revision, replacementOffset: next })
      ).replacementInspection;
      assert.equal(page.revision, r.revision);
      all.push(...page.rows);
      next = page.nextOffset;
    }
    assert.equal(all.length, rows);
    assert.equal(new Set(all.map((row) => row.index)).size, rows);
    assert.equal(all.at(-1).sourceRef, 'A205');
    assert.equal(all.at(-1).destinationRef, 'B207');
    assert.equal(all.at(-1).current.display, JSON.stringify('Original 204'));
    assert(!JSON.stringify(p).includes(f.root));
    assert.equal(f.accepted(), 0);
  }
);
check('growth claims union padding beyond true output and still protects the prior managed rectangle', async (t) => {
  const query = { columns: [], filters: [], keepHeader: false },
    f = await createControlledQueryOwner(t, output, {
      text: Array.from({ length: 9 }, (_, i) => 'L' + i + ',R' + i).join('\n'),
      query,
      target: { sheet: 'S', start: 'B2' },
    });
  await f.host().refresh(f.record.id, 1);
  assert.equal(f.accepted(), 1);
  await fs.writeFile(f.filename, Array.from({ length: 7 }, (_, i) => 'N' + i + ',M' + i + ',D' + i).join('\n'));
  f.cells.D9 = valueCell('Keep tail');
  f.cells.D10 = valueCell(9, {
    formula: 'SUM(Z1)',
    calculation: { status: 'stored', pending: true },
    cachedResult: { authority: 'stored', status: 'calculated', value: 9 },
    cachedValue: '9',
  });
  const p = await preview(f, 2),
    r = p.replacementInspection;
  assert.equal(p.outputRows, 7);
  assert.equal(r.occupied, 2);
  assert.equal(r.totalNewPositions, 9);
  assert.deepEqual(
    r.rows.map((row) => row.destinationRef),
    ['D9', 'D10']
  );
  assert(r.rows.every((row) => row.sourceRef === null && row.sourceText === '' && row.outputKind === 'text'));
  assert.equal(r.rows[1].current.formula, 'SUM(Z1)');
  assert.equal(r.rows[1].current.pending, true);
  const prior = f.cells.B2;
  f.cells.B2 = valueCell('Manual changed managed output');
  await assert.rejects(f.host().refresh(f.record.id, 2, reviewed(p)), {
    code: /^(layout-stale|pivot-stale|query-stale)$/,
  });
  f.cells.B2 = prior;
  await f.host().refresh(f.record.id, 2, reviewed(p));
  assert.equal(f.accepted(), 2);
  assert.equal(f.cells.D9.value, '');
  assert.equal(f.cells.D10.value, '');
  assert.equal(f.cells.D10.formula, undefined);
});
check(
  'target relocation starts a fresh review and old definition/current-head/source/destination changes invalidate acknowledgment',
  async (t) => {
    const f = await createControlledQueryOwner(t, output);
    f.cells.B3 = valueCell('First');
    const first = await preview(f);
    await f.host().refresh(f.record.id, 1, reviewed(first));
    f.cells.G8 = valueCell('New target');
    await f
      .host()
      .update(f.record.id, 2, { name: 'Moved', target: { sheet: 'S', start: 'G8' }, query: f.record.query });
    await assert.rejects(f.host().refresh(f.record.id, 3, reviewed(first)), /review|changed|expired/i);
    const p = await preview(f, 3);
    assert.equal(p.replacementInspection.rows[0].destinationRef, 'G8');
    f.cells.G8 = valueCell('Changed after review');
    await assert.rejects(f.host().refresh(f.record.id, 3, reviewed(p)), /destination|review|changed/i);
    assert.equal(f.accepted(), 1);
    const fresh = await preview(f, 3);
    f.scope.ownerId = 'replacement';
    await assert.rejects(f.host().refresh(f.record.id, 3, reviewed(fresh)), /review|changed|expired/i);
    assert.equal(f.accepted(), 1);
  }
);
check(
  'ordinary formulas show authored input and stored/pending provenance separately, not a fresh cached scalar',
  async (t) => {
    const f = await createControlledQueryOwner(t, output);
    f.cells.B3 = valueCell(913, {
      formula: 'SUM(Z1:Z3)',
      calculation: { status: 'stored', pending: true },
      cachedValue: '913',
      cachedResult: { authority: 'stored', status: 'calculated', value: 912, revision: 'old' },
    });
    const p = await preview(f),
      row = p.replacementInspection.rows[0];
    assert.equal(row.current.kind, 'formula');
    assert.equal(row.current.formula, 'SUM(Z1:Z3)');
    assert.match(row.current.display, /SUM\(Z1:Z3\)/);
    assert.notEqual(row.current.display, '913');
    assert.equal(row.current.pending, true);
    assert.equal(p.replacementInspection.blocked, 0);
    await f.host().refresh(f.record.id, 1, reviewed(p));
    assert.equal(f.accepted(), 1);
    assert.equal(f.cells.B3.formula, undefined);
  }
);
check(
  'exact raw precision and signed zero are reviewed; special context is blocked without staging or partial owner writes',
  async (t) => {
    const f = await createControlledQueryOwner(t, output);
    f.cells.B3 = valueCell(9007199254740992, { rawValue: '9007199254740993' });
    let p = await preview(f),
      row = p.replacementInspection.rows[0];
    assert.equal(row.current.rawValue, '9007199254740993');
    assert(row.current.display.includes('9007199254740993'));
    assert(!row.current.display.includes('9007199254740992'));
    f.cells.B3 = valueCell(-0, { rawValue: '0' });
    await assert.rejects(f.host().refresh(f.record.id, 1, reviewed(p)), /destination|review|changed/i);
    p = await preview(f);
    assert.equal(p.replacementInspection.rows[0].current.display, '-0');
    f.cells.B3 = valueCell(2, { formulaType: 'array', formula: 'SUM(Z1)', formulaReference: 'B3:B4' });
    p = await preview(f);
    assert.equal(p.replacementInspection.blocked, 1);
    await assert.rejects(f.host().refresh(f.record.id, 1, reviewed(p)), /protected|special|replace|blocked/i);
    assert.equal(f.accepted(), 0);
    assert.equal(f.requests.filter((r) => r.operation === 'importBegin').length, 0);
  }
);
check(
  'empty new destination preserves legacy text-only refresh; stale replacement tokens cannot be ignored on a text definition',
  async (t) => {
    const f = await createControlledQueryOwner(t, output, {
      text: '00123\n',
      query: { columns: [], filters: [], keepHeader: false },
    });
    await f.host().refresh(f.record.id, 1);
    assert.equal(f.cells.B3.value, '00123');
    const p = await preview(f, 2);
    await assert.rejects(
      f.host().refresh(f.record.id, 2, { expectedReplacementRevision: randomUUID() }),
      /review|expired|replacement/i
    );
    assert.equal(f.accepted(), 1);
    assert.equal(p.replacementInspection.occupied, 0);
  }
);

check('filtered far destination cells cannot block or enter reads beyond the exact output union', async (t) => {
  const f = await createControlledQueryOwner(t, output, {
    text: '1,yes\n2,no\n3,no\n',
    query: {
      columns: [1],
      filters: [{ column: 2, op: 'equals', value: 'yes', caseSensitive: true }],
      keepHeader: false,
      policies: ['number'],
    },
  });
  f.cells.B4 = valueCell('x'.repeat(2 * 1024 * 1024));
  const p = await preview(f);
  assert.equal(p.outputRows, 1);
  assert.equal(p.replacementInspection.occupied, 0);
  for (const request of f.requests.filter((r) => r.operation === 'read')) assert.equal(request.input.range, 'B3:B3');
  assert.equal(f.accepted(), 0);
});

check(
  'combined long escaped review rows stay below1MiB and both lists continue without truncation or starvation',
  async (t) => {
    const count = 30,
      raw = ' '.repeat(30000) + '1.00',
      f = await createControlledQueryOwner(t, output, { text: Array.from({ length: count }, () => raw).join('\n') });
    for (let i = 0; i < count; i++) f.cells['B' + (i + 3)] = valueCell('"'.repeat(30000));
    const p = await preview(f);
    assert(Buffer.byteLength(JSON.stringify({ queryPreview: p })) <= 1024 * 1024);
    assert.equal(p.conversionInspection.totalIssues, count);
    const original = [];
    let next = p.replacementInspection.nextOffset;
    original.push(...p.replacementInspection.rows);
    while (next !== null) {
      const page = await preview(f, 1, {
        expectedConversionRevision: p.conversionInspection.revision,
        replacementOffset: next,
      });
      assert(Buffer.byteLength(JSON.stringify({ queryPreview: page })) <= 1024 * 1024);
      assert(page.replacementInspection.rows.length > 0);
      original.push(...page.replacementInspection.rows);
      next = page.replacementInspection.nextOffset;
    }
    assert.equal(original.length, count);
    assert(original.every((row) => row.sourceText === raw && row.current.rawValue === '"'.repeat(30000)));
    const issues = [...p.conversionInspection.issues];
    next = p.conversionInspection.nextOffset;
    while (next !== null) {
      const page = await preview(f, 1, { expectedConversionRevision: p.conversionInspection.revision, offset: next });
      assert(Buffer.byteLength(JSON.stringify({ queryPreview: page })) <= 1024 * 1024);
      issues.push(...page.conversionInspection.issues);
      next = page.conversionInspection.nextOffset;
    }
    assert.equal(issues.length, count);
    assert(issues.every((issue) => issue.raw === raw));
    assert.equal(f.accepted(), 0);
  }
);

check(
  'initial destination reads never exceed16positions before framed reply sizing and still expose large originals',
  async (t) => {
    const count = 200,
      f = await createControlledQueryOwner(t, output, {
        text: Array.from({ length: count }, (_, i) => String(i + 1)).join('\n'),
      });
    for (let i = 0; i < count; i++) f.cells['B' + (i + 3)] = valueCell('x'.repeat(30000));
    const p = await preview(f);
    assert.equal(p.replacementInspection.occupied, count);
    for (const request of f.requests.filter((r) => r.operation === 'read')) {
      const b = rangeBounds(request.input.range);
      assert((b.last.row - b.first.row + 1) * (b.last.col - b.first.col + 1) <= 16);
    }
    assert(p.replacementInspection.rows.length > 0);
    assert(p.replacementInspection.rows.every((row) => row.current.rawValue === 'x'.repeat(30000)));
    assert.equal(f.accepted(), 0);
  }
);

check(
  'only exact recoverable transport errors shrink reads; singlecelloversize and other errors refuse without staging',
  async (t) => {
    const f = await createControlledQueryOwner(t, output, {
        text: Array.from({ length: 20 }, (_, i) => String(i + 1)).join('\n'),
      }),
      original = f.context.request,
      attempts = [];
    f.context.request = async (operation, input) => {
      if (operation === 'read') {
        const b = rangeBounds(input.range),
          positions = (b.last.row - b.first.row + 1) * (b.last.col - b.first.col + 1);
        attempts.push(positions);
        if (positions > 4) throw Error('Native workbook reply exceeds transport bound');
      }
      return original(operation, input);
    };
    const p = await preview(f);
    assert.equal(p.outputRows, 20);
    assert.equal(attempts[0], 16);
    assert(attempts.includes(8));
    assert(attempts.includes(4));
    assert.equal(f.accepted(), 0);
    const single = await createControlledQueryOwner(t, output);
    single.context.request = async (operation) => {
      if (operation === 'read') throw Error('Native workbook reply exceeds transport bound');
      return {};
    };
    await assert.rejects(preview(single), { code: 'query-limit' });
    assert.equal(single.accepted(), 0);
    assert.equal(single.requests.filter((r) => r.operation === 'importBegin').length, 0);
    const other = await createControlledQueryOwner(t, output);
    let reads = 0;
    other.context.request = async (operation) => {
      if (operation === 'read') {
        reads++;
        throw Error('Unrelated native read failure');
      }
      return {};
    };
    await assert.rejects(preview(other), /Unrelated native read failure/);
    assert.equal(reads, 1);
    assert.equal(other.accepted(), 0);
  }
);
