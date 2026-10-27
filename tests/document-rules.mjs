import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
test('rule collections are one durable undo frame and replay exact guarded settings after restart', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'document-rules-')),
    recovery = path.join(directory, 'recovery'),
    owners = [],
    create = () => {
      const owner = new SpreadsheetDocument({
        recoveryDirectory: recovery,
        indexDirectory: path.join(directory, 'indexes'),
      });
      owners.push(owner);
      return owner;
    };
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  let owner = create();
  await owner.create('Rules');
  const id = owner.status().documentId,
    apply = (commands) => owner.request({ operation: 'apply', state: owner.status().state, commands }),
    cf = {
      operation: 'setConditionalRules',
      sheet: 'Sheet1',
      expectedRulesRevision: null,
      rules: [{ range: 'A1:A5', kind: 'between', values: [1, 5], fillColor: '#FFCC00', stopIfTrue: true }],
    },
    dv = {
      operation: 'setValidationRules',
      sheet: 'Sheet1',
      expectedRulesRevision: null,
      rules: [
        {
          range: 'B1:B5',
          kind: 'list',
          values: ['Yes', 'No'],
          allowBlank: true,
          showDropdown: true,
          promptTitle: 'Choose',
          prompt: 'Select Yes or No',
          errorTitle: 'Invalid',
          error: 'Choose a listed value',
        },
      ],
    };
  await apply([cf]);
  const cfState = owner.status().state;
  await apply([dv]);
  const bothState = owner.status().state;
  assert.equal(owner.status().view.sheets[0].validationRules.length, 1);
  await owner.request({ operation: 'undo' });
  assert.equal(owner.status().state, cfState);
  assert.equal(owner.status().view.sheets[0].validationRules.length, 0);
  await owner.request({ operation: 'redo' });
  assert.equal(owner.status().state, bothState);
  await apply([{ sheet: 'Sheet1', cell: 'B1', value: 'Yes' }]);
  const accepted = owner.status().state;
  await assert.rejects(apply([{ sheet: 'Sheet1', cell: 'B1', value: 'Invalid' }]), { code: 'rules-refused' });
  assert.equal(owner.status().state, accepted);
  await owner.close();
  owner = create();
  await owner.recover(id);
  assert.equal(owner.status().view.sheets[0].conditionalRules[0].rule, 'between');
  assert.equal(owner.status().view.sheets[0].validationRules[0].values[0], 'Yes');
  assert.equal(
    (await owner.request({ operation: 'read', state: owner.status().state, sheet: 'Sheet1', range: 'B1' })).cells.B1
      .value,
    'Yes'
  );
  await owner.request({ operation: 'undo' });
  assert.equal(
    (await owner.request({ operation: 'read', state: owner.status().state, sheet: 'Sheet1', range: 'B1' })).cells.B1
      ?.value ?? null,
    null
  );
  await owner.request({ operation: 'redo' });
  const destination = path.join(directory, 'rules-copy.xlsx');
  await owner.request({ operation: 'copy', destination, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(destination)),
    xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
  assert.match(xml, /operator="between"/);
  assert.match(xml, /stopIfTrue="1"/);
  assert.match(xml, /type="list"/);
  assert.match(xml, /&quot;Yes,No&quot;|"Yes,No"/);
  const draft = path.join(recovery, 'draft-' + id),
    nodes = (await fs.readdir(draft)).filter(
      (name) => name.endsWith('.json') && name !== 'current.json' && name !== 'paste.json'
    );
  let journals = 0;
  for (const file of nodes) {
    const record = JSON.parse(await fs.readFile(path.join(draft, file), 'utf8'));
    if (record.schema === 1 && typeof record.commands === 'string') {
      journals++;
      assert.ok(Buffer.byteLength(record.commands) < 4096);
    }
  }
  assert.equal(journals, 3);
});
