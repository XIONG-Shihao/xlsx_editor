/** Controlled owner/stager for source host tests. This is not native replay,
 * durable owner, Save, or cache authority. Real owner tests are separate. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { LocalQueryConnections } from '../../adapters/local-query-connections.mjs';
import { LocalQueryHost } from '../../adapters/local-query-host.mjs';
import { convertTextField } from '../../core/text-conversion.mjs';
import { cellPosition, columnName, rangeBounds } from '../../core/workbook-commands.mjs';
export async function createControlledQueryOwner(
  t,
  directory,
  {
    text = '1\n',
    query = { columns: [], filters: [], keepHeader: false, policies: ['number'] },
    target = { sheet: 'S', start: 'B3' },
  } = {}
) {
  const root = await fs.mkdtemp(path.join(directory, 'owner-')),
    filename = path.join(root, 'source.csv'),
    staging = path.join(root, 'private');
  await fs.writeFile(filename, text);
  const store = await LocalQueryConnections.open(path.join(root, 'connections.json')),
    record = await store.create('doc', { name: 'Imported data', target, query }, filename),
    scope = { documentId: 'doc', ownerId: 'owner', state: 'head' },
    cells = {},
    requests = [],
    uploads = new Map(),
    receipts = new Map();
  let accepted = 0,
    applyFailure = null;
  const context = {
    current: () => ({ ...scope }),
    calendar: () => false,
    request: async (operation, input) => {
      requests.push({ operation, input: structuredClone(input) });
      if (operation === 'read') {
        const b = rangeBounds(input.range);
        return {
          cells: Object.fromEntries(
            Object.entries(cells)
              .filter(([ref]) => {
                const p = cellPosition(ref);
                return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
              })
              .map(([ref, cell]) => [ref, structuredClone(cell)])
          ),
        };
      }
      if (operation === 'importBegin') {
        const id = randomUUID();
        uploads.set(id, {
          plan: {
            shape: input.shape,
            count: input.count,
            policy: input.policy,
            ...(input.conversion ? { conversion: input.conversion } : {}),
          },
          cells: [],
          sequence: 0,
        });
        return { id, sequence: 0 };
      }
      if (operation === 'importUpload') {
        const upload = uploads.get(input.uploadId);
        upload.cells.push(...input.cells);
        return { id: input.uploadId, sequence: ++upload.sequence, count: upload.cells.length };
      }
      if (operation === 'importSeal') {
        const upload = uploads.get(input.uploadId),
          descriptor = {
            type: 'importCells',
            id: input.uploadId,
            sha256: 'a'.repeat(64),
            bytes: 100,
            ...upload.plan,
            ...(upload.plan.conversion ? { schema: 'workbook.xlsx.staged-paste.v2' } : {}),
          };
        upload.descriptor = descriptor;
        return { descriptor };
      }
      if (operation === 'importReceipt') return receipts.get(input.descriptor.id) ?? { found: false };
      return {};
    },
    apply: async (parent, descriptor) => {
      if (applyFailure) throw applyFailure;
      if (parent !== scope.state) throw Error('Controlled old head refused');
      const upload = uploads.get(descriptor.id),
        start = cellPosition(descriptor.shape.start),
        conversion = descriptor.conversion;
      for (const [index, text] of upload.cells) {
        const row = Math.floor(index / descriptor.shape.columns),
          col = index % descriptor.shape.columns,
          policy =
            conversion &&
            row < conversion.outputRows &&
            col < conversion.outputColumns &&
            !(conversion.keepHeader && row === 0)
              ? conversion.policies[col]
              : 'text',
          result = convertTextField(text, policy, { date1904: conversion?.date1904 });
        if (result.status !== 'converted') throw Error(result.reason);
        const ref = columnName(start.col + col) + (start.row + row + 1),
          prior = cells[ref] ?? {};
        cells[ref] = {
          ...prior,
          value: result.value,
          dataType: result.kind === 'date' ? 'number' : result.kind,
          rawValue: typeof result.value === 'number' ? String(result.value) : String(result.value),
          ...(result.kind === 'date' ? { format: 14 } : {}),
        };
        for (const key of ['formula', 'formulaType', 'calculation', 'cachedResult', 'error', 'spill'])
          delete cells[ref][key];
      }
      scope.state = 'accepted-' + ++accepted;
      receipts.set(descriptor.id, { found: true, head: scope.state, parentState: parent, descriptor });
      return { state: scope.state };
    },
  };
  const host = () => new LocalQueryHost({ store, document: 'doc', directory: staging, context });
  t.after(async () => {
    await host().disposeInspections('owner');
    await host().disposeInspections('replacement');
  });
  return {
    root,
    filename,
    staging,
    store,
    record,
    scope,
    cells,
    requests,
    context,
    host,
    accepted: () => accepted,
    refuse: (error) => {
      applyFailure = error;
    },
  };
}
