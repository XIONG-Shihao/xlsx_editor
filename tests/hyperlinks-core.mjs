import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  hyperlinkCell,
  hyperlinkCommand,
  hyperlinkDestination,
  normalizeHyperlinkCommands,
  hyperlinkFields,
  bindHyperlink,
  validateHyperlink,
} from '../core/hyperlinks.mjs';
import { resolveWorkbookEdits } from '../core/workbook-commands.mjs';
const url = () => ({
  type: 'setHyperlink',
  sheet: 'Main',
  cell: 'A1',
  expectedHyperlinkRevision: null,
  link: { kind: 'url', target: 'https://example.test/a?x=1&y=2#here' },
  tooltip: 'Visit & inspect',
});
const binding = { documentId: 'document-A', ownerId: 'owner-A', state: 'head-A', revision: 'revision-A' },
  expected = { sheet: 'Main', cell: 'A1' };
const raw = {
  schema: 'workbook.xlsx.hyperlink-definition.v1',
  ...expected,
  hyperlinkRevision: null,
  link: null,
  tooltip: '',
  canEdit: true,
  protectedReason: null,
};
test('closed URL/internal commands preserve literal destinations and reject metadata/value injection', () => {
  assert.deepEqual(hyperlinkCommand(url()), url());
  const inside = { ...url(), link: { kind: 'internal', sheet: "Dest Ω's", cell: 'XFD1048576' } };
  assert.deepEqual(hyperlinkCommand(inside), inside);
  assert.deepEqual(
    hyperlinkCommand({ type: 'removeHyperlink', ...expected, expectedHyperlinkRevision: 'a'.repeat(64) }),
    { type: 'removeHyperlink', ...expected, expectedHyperlinkRevision: 'a'.repeat(64) }
  );
  for (const patch of [
    { value: 'changed text' },
    { relationshipId: 'rId1' },
    { operation: 'clearRange' },
    { expectedHyperlinkRevision: 'wrong' },
    { cell: 'A1:B2' },
    { link: { kind: 'url', target: 'https://example.test/', extra: true } },
  ])
    assert.throws(() => hyperlinkCommand({ ...url(), ...patch }));
});
test('ordinary addresses/URLs have explicit finite bounds and refuse scripts credentials and invalid XML', () => {
  for (const cell of ['A1', 'XFD1048576', 'Z80']) assert.equal(hyperlinkCell(cell), cell);
  for (const cell of ['a1', '$A$1', 'A01', 'XFE1', 'A1048577', 'A1:A1', 'Sheet!A1'])
    assert.throws(() => hyperlinkCell(cell));
  for (const target of [
    'javascript:alert(1)',
    'file:///tmp/source',
    'mailto:person@example.test',
    'https://name:password@example.test/',
    'https://example.test/ bad',
    'https://example.test/%QQ',
    'https://example.test\\path',
  ])
    assert.throws(() => hyperlinkDestination({ kind: 'url', target }));
  assert.throws(() => hyperlinkCommand({ ...url(), tooltip: 'a'.repeat(257) }));
  assert.throws(() => hyperlinkCommand({ ...url(), tooltip: '\ud800' }));
  assert.throws(() => hyperlinkDestination({ kind: 'internal', sheet: 'Bad:Sheet', cell: 'A1' }));
});
test('queued normalized batches own destinations and check source/target worksheet authority', () => {
  const original = [url(), { sheet: 'Main', cell: 'B1', value: 2 }],
    captured = normalizeHyperlinkCommands(original, { sheets: [{ name: 'Main' }] });
  original[0].link.target = 'https://changed.test/';
  original[1].value = 9;
  assert.equal(captured[0].link.target, 'https://example.test/a?x=1&y=2#here');
  assert.equal(captured[1].value, 2);
  assert.throws(() => normalizeHyperlinkCommands([url()], { sheets: [] }));
  assert.throws(() =>
    normalizeHyperlinkCommands([{ ...url(), link: { kind: 'internal', sheet: 'Missing', cell: 'A1' } }], {
      sheets: [{ name: 'Main' }],
    })
  );
  assert.throws(() => normalizeHyperlinkCommands([url(), { sheet: 'Main', cell: 'B1', value: 'x'.repeat(2097152) }]));
});
test('inspection requests and replies remain closed and original-owner/head bound', () => {
  const request = { operation: 'hyperlink', documentId: binding.documentId, state: binding.state, ...expected };
  assert.equal(hyperlinkFields(request), request);
  hyperlinkFields({ ...request, kind: 'indexed', sessionId: 'session-A' }, { transport: true });
  for (const extra of [{ source: 'source.xlsx' }, { expectedHyperlinkRevision: null }, { kind: 'indexed' }])
    assert.throws(() => hyperlinkFields({ ...request, ...extra }));
  const value = bindHyperlink(raw, binding, expected);
  assert.equal(validateHyperlink(value, binding, expected), value);
  for (const patch of [
    { state: 'head-B' },
    { ownerId: 'owner-B' },
    { documentId: 'document-B' },
    { revision: 'revision-B' },
    { sheet: 'Elsewhere' },
    { cell: 'B1' },
    { privateRoot: '/private/path' },
  ])
    assert.throws(() => validateHyperlink({ ...value, ...patch }, binding, expected));
});
test('protected and editable read definitions never borrow authority from an internal prerequisite', () => {
  const protectedValue = {
    ...raw,
    hyperlinkRevision: 'b'.repeat(64),
    canEdit: false,
    protectedReason: 'Imported range link remains protected.',
  };
  validateHyperlink(bindHyperlink(protectedValue, binding, expected), binding, expected);
  for (const patch of [
    { canEdit: false },
    { link: { kind: 'url', target: 'https://example.test/' } },
    { hyperlinkRevision: 'bad' },
    { protectedReason: '' },
    { protectedReason: 'Blocked' },
    { hyperlinkRevision: 'a'.repeat(64) },
  ])
    assert.throws(() => bindHyperlink({ ...raw, ...patch }, binding, expected));
  const editable = {
    ...raw,
    hyperlinkRevision: 'a'.repeat(64),
    link: { kind: 'url', target: 'https://example.test/' },
  };
  bindHyperlink(editable, binding, expected);
});
test('portable cell projection refuses hyperlink commands without changing cells', () => {
  const book = { sheets: [{ name: 'Main', cells: { A1: { value: 'keep text', styleId: 1 } } }] },
    before = structuredClone(book);
  assert.throws(() => resolveWorkbookEdits(book, [url()]), /native XLSX/);
  assert.deepEqual(book, before);
});
