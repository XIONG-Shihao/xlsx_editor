import './support/scratch-dir.mjs';
/** Current production model + actual Document/native literal batches.
 * No library/native build, GUI, package or general Excel parity claim. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { build } from 'esbuild';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookBinary, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
const component = fileURLToPath(new URL('../', import.meta.url)),
  folder = path.join(component, '.build/tests'),
  run = await fs.mkdtemp(path.join(folder, 'text-columns-owner-')),
  module = path.join(run, 'model.mjs'),
  entry = path.join(run, 'entry.ts'),
  sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  helper = nativeWorkbookBinary(),
  helperHash = sha(await fs.readFile(helper));
await fs.writeFile(
  entry,
  `export * from ${JSON.stringify(path.join(component, 'ui/textColumns.ts'))};export * from ${JSON.stringify(path.join(component, 'ui/textColumnsProtection.ts'))};`
);
await build({
  entryPoints: [entry],
  bundle: true,
  platform: 'node',
  format: 'esm',
  outfile: module,
  logLevel: 'silent',
});
const { previewTextColumns, applyTextColumns, textColumnProtection } = await import(pathToFileURL(module));
after(async () => {
  closeNativeWorkbookWorker();
});
test('actual native literal split recalculates outside SUM/LEN chain, has one Undo/Redo, preserves formula-looking/zero text and saves typed caches', async (t) => {
  const original = unzipSync(createWorkbook());
  original['custom/opaque.bin'] = new Uint8Array([0, 255, 19, 10]);
  original['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>00123,alpha</t></is></c><c r="B1"><v>5</v></c><c r="E1"><f>SUM(B1:B2)</f><v>12</v></c><c r="F1"><f>SUM(E1:E2)</f><v>23</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>00456,=SUM(Z1:Z2)</t></is></c><c r="B2"><v>7</v></c><c r="E2"><f>LEN(A1)</f><v>11</v></c></row></sheetData></worksheet>'
  );
  const source = path.join(run, 'source.xlsx');
  await fs.writeFile(source, zipSync(original));
  const sourceHash = sha(await fs.readFile(source)),
    doc = new SpreadsheetDocument({ recoveryDirectory: path.join(run, 'recovery'), cacheIdentity: null });
  await doc.open(source);
  t.after(() => doc.close());
  const first = doc.status(),
    capturedState = first.state,
    capturedDocument = first.documentId,
    capturedRuntime = first.view.storage.sessionId;
  const request = (operation, fields = {}) =>
      doc.request({ operation, documentId: doc.status().documentId, state: doc.status().state, ...fields }),
    values = async () => (await request('read', { sheet: 'Sheet1', range: 'A1:F2' })).cells;
  assert.equal((await values()).E1.calculation.value, 12);
  assert.equal((await values()).E2.calculation.value, 11);
  assert.equal((await values()).F1.calculation.value, 23);
  let accepts = 0;
  const context = {
    binding: {
      documentId: 'doc',
      grantDocumentId: capturedDocument,
      ownerIdentity: capturedRuntime,
      ownerKey: capturedRuntime,
      state: capturedRuntime + ':' + capturedState,
      sheet: 'Sheet1',
      selection: 'A1:A2',
      navigationGeneration: 0,
      activationEpoch: 0,
    },
    isCurrent: () => doc.status().state === capturedState && doc.status().view.storage.sessionId === capturedRuntime,
    isRuntimeCurrent: () => doc.status().view.storage.sessionId === capturedRuntime,
    availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
    read: async (range) =>
      (
        await doc.request({
          operation: 'read',
          documentId: capturedDocument,
          state: capturedState,
          sheet: 'Sheet1',
          range,
        })
      ).cells,
    inspect: async (source, destination) => {
      const conditional = await doc.request({
          operation: 'conditionalRules',
          documentId: capturedDocument,
          state: capturedState,
          sheet: 'Sheet1',
        }),
        validation = await doc.request({
          operation: 'validationRules',
          documentId: capturedDocument,
          state: capturedState,
          sheet: 'Sheet1',
        });
      return textColumnProtection(
        {
          sheet: first.view.sheets[0],
          tables: first.view.tables ?? [],
          pivots: first.view.pivots ?? [],
          conditional,
          validation,
        },
        source,
        destination
      );
    },
    accept: async (commands) => {
      accepts++;
      await doc.request({ operation: 'apply', documentId: capturedDocument, state: capturedState, commands });
    },
  };
  const preview = await previewTextColumns(context, { delimiter: ',', quotes: 'double', destination: 'A1' });
  assert.equal(preview.collisions, 4);
  assert.equal(preview.commands.length, 1);
  assert.equal(preview.commands[0].type, 'setCells');
  assert(preview.commands[0].cells.every((cell) => typeof cell.value === 'string' && !Object.hasOwn(cell, 'formula')));
  await assert.rejects(applyTextColumns(context, preview), /occupied destination/);
  assert.equal(accepts, 0);
  assert.equal(doc.status().state, capturedState);
  const result = await applyTextColumns(context, preview, { replaceFingerprint: preview.fingerprint });
  assert.equal(result.outcome, 'accepted');
  assert.equal(result.runtimeCurrent, true);
  assert.equal(result.current, false);
  assert.equal(accepts, 1);
  assert(doc.status().canUndo && !doc.status().canRedo);
  let cells = await values();
  assert.equal(cells.A1.value, '00123');
  assert.equal(cells.A2.value, '00456');
  assert.equal(cells.B1.value, 'alpha');
  assert.equal(cells.B2.value, '=SUM(Z1:Z2)');
  for (const ref of ['A1', 'A2', 'B1', 'B2']) {
    assert.equal(cells[ref].dataType, 'text');
    assert.equal(cells[ref].formula, undefined);
  }
  assert.equal(cells.E1.calculation.value, 0);
  assert.equal(cells.E2.calculation.value, 5);
  assert.equal(cells.F1.calculation.value, 5);
  await request('undo');
  cells = await values();
  assert.equal(cells.A1.value, '00123,alpha');
  assert.equal(cells.B1.value, 5);
  assert.equal(cells.E1.calculation.value, 12);
  assert.equal(cells.F1.calculation.value, 23);
  assert(!doc.status().canUndo && doc.status().canRedo);
  await request('redo');
  cells = await values();
  assert.equal(cells.B2.value, '=SUM(Z1:Z2)');
  assert.equal(cells.E1.calculation.value, 0);
  assert.equal(cells.F1.calculation.value, 5);
  const saved = path.join(run, 'saved-copy.xlsx');
  await request('copy', { destination: saved, expectedRevision: null });
  assert.equal(sha(await fs.readFile(source)), sourceHash);
  const oracle = spawnSync(
    process.env.XLSX_EDITOR_PYTHON ?? 'python3',
    [
      '-c',
      `import sys,zipfile,xml.etree.ElementTree as E,json\nn='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'\nwith zipfile.ZipFile(sys.argv[1]) as a,zipfile.ZipFile(sys.argv[2]) as b:\n assert a.read('custom/opaque.bin')==b.read('custom/opaque.bin')\n cells={c.attrib['r']:c for c in E.fromstring(b.read('xl/worksheets/sheet1.xml')).iter(n+'c')}\n expected={'A1':'00123','A2':'00456','B1':'alpha','B2':'=SUM(Z1:Z2)'}\n for ref,value in expected.items():\n  c=cells[ref];assert c.attrib['t']=='inlineStr' and c.find(n+'f') is None\n  assert ''.join(t.text or '' for t in c.iter(n+'t'))==value\n for ref,value in {'E1':0,'E2':5,'F1':5}.items():\n  c=cells[ref];assert c.find(n+'f') is not None and c.attrib.get('t','n')=='n' and float(c.find(n+'v').text)==value\n print(json.dumps({'typedText':expected,'outsideFreshCaches':{'E1':0,'E2':5,'F1':5},'opaquePreserved':True}))`,
      source,
      saved,
    ],
    { encoding: 'utf8' }
  );
  assert.equal(oracle.status, 0, oracle.stderr);
  await fs.writeFile(path.join(run, 'independent-readback.json'), oracle.stdout);
  const reopened = new SpreadsheetDocument({ recoveryDirectory: path.join(run, 'cold'), cacheIdentity: null });
  await reopened.open(saved);
  t.after(() => reopened.close());
  const cold = (
    await reopened.request({
      operation: 'read',
      documentId: reopened.status().documentId,
      state: reopened.status().state,
      sheet: 'Sheet1',
      range: 'A1:F2',
    })
  ).cells;
  assert.equal(cold.A1.value, '00123');
  assert.equal(cold.B2.formula, undefined);
  assert.equal(cold.F1.calculation.value, 5);
});

test('actual native permits a split disjoint from protected custom validation and unevaluable known duplicate rules, preserving both XML definitions', async (t) => {
  const original = unzipSync(createWorkbook()),
    validation =
      '<dataValidations count="1"><dataValidation type="custom" sqref="Z10:Z11" showErrorMessage="1"><formula1>LEN(Z10)&gt;0</formula1></dataValidation></dataValidations>',
    conditional =
      '<conditionalFormatting sqref="Z10:Z11"><cfRule type="duplicateValues" dxfId="0" priority="1" stopIfTrue="0"/></conditionalFormatting>';
  original['custom/opaque.bin'] = new Uint8Array([42, 0, 255]);
  original['xl/styles.xml'] = strToU8(
    new TextDecoder()
      .decode(original['xl/styles.xml'])
      .replace(
        '</styleSheet>',
        '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs></styleSheet>'
      )
  );
  original['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>00123,=1+2</t></is></c><c r="E1"><f>LEN(A1)</f><v>10</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>00456,tail</t></is></c></row><row r="10"><c r="Z10"><v>9007199254740993</v></c></row><row r="11"><c r="Z11"><v>2</v></c></row></sheetData>' +
      conditional +
      validation +
      '</worksheet>'
  );
  const source = path.join(run, 'protected-disjoint-source.xlsx');
  await fs.writeFile(source, zipSync(original));
  const sourceHash = sha(await fs.readFile(source)),
    doc = new SpreadsheetDocument({
      recoveryDirectory: path.join(run, 'protected-disjoint-recovery'),
      cacheIdentity: null,
    });
  await doc.open(source);
  t.after(() => doc.close());
  const first = doc.status(),
    state = first.state,
    id = first.documentId,
    owner = first.view.storage.sessionId;
  const request = (operation, fields = {}) =>
      doc.request({ operation, documentId: id, state: doc.status().state, ...fields }),
    conditionalDefinition = await request('conditionalRules', { sheet: 'Sheet1' }),
    validationDefinition = await request('validationRules', { sheet: 'Sheet1' });
  assert.equal(conditionalDefinition.canEdit, true);
  assert(conditionalDefinition.evaluationReasons.length > 0);
  assert.equal(validationDefinition.canEdit, false);
  assert.deepEqual(validationDefinition.protectedRanges, ['Z10:Z11']);
  assert.equal((await request('read', { sheet: 'Sheet1', range: 'E1:E1' })).cells.E1.calculation.value, 10);
  const inventory = {
    sheet: first.view.sheets[0],
    tables: first.view.tables ?? [],
    pivots: first.view.pivots ?? [],
    conditional: conditionalDefinition,
    validation: validationDefinition,
  };
  assert.equal(textColumnProtection(inventory, 'A1:A2', 'C1:D2').reason, null);
  assert(textColumnProtection(inventory, 'A1:A2', 'Z10:AA11').reason);
  let accepts = 0;
  const context = {
    binding: {
      documentId: 'protected-doc',
      grantDocumentId: id,
      ownerIdentity: owner,
      ownerKey: owner,
      state: owner + ':' + state,
      sheet: 'Sheet1',
      selection: 'A1:A2',
      navigationGeneration: 0,
      activationEpoch: 0,
    },
    isCurrent: () => doc.status().state === state,
    isRuntimeCurrent: () => doc.status().view.storage.sessionId === owner,
    availability: () => ({ active: true, readOnly: false, busy: false, literal: false, specialSelection: false }),
    read: async (range) =>
      (await doc.request({ operation: 'read', documentId: id, state, sheet: 'Sheet1', range })).cells,
    inspect: async (source, destination) => textColumnProtection(inventory, source, destination),
    accept: async (commands) => {
      accepts++;
      await doc.request({ operation: 'apply', documentId: id, state, commands });
    },
  };
  const preview = await previewTextColumns(context, { delimiter: ',', quotes: 'double', destination: 'C1' });
  assert.equal(preview.collisions, 0);
  await assert.rejects(
    previewTextColumns(context, { delimiter: ',', quotes: 'double', destination: 'Z10' }),
    /intersects/
  );
  assert.equal(accepts, 0);
  await applyTextColumns(context, preview);
  assert.equal(accepts, 1);
  let cells = (await request('read', { sheet: 'Sheet1', range: 'A1:E2' })).cells;
  assert.equal(cells.C1.value, '00123');
  assert.equal(cells.D1.value, '=1+2');
  assert.equal(cells.D1.formula, undefined);
  assert.equal(cells.E1.calculation.value, 10);
  await request('undo');
  cells = (await request('read', { sheet: 'Sheet1', range: 'C1:D2' })).cells;
  assert.equal(cells.C1, undefined);
  assert.equal(cells.D1, undefined);
  assert(!doc.status().canUndo && doc.status().canRedo);
  await request('redo');
  const saved = path.join(run, 'protected-disjoint-saved.xlsx');
  await request('copy', { destination: saved, expectedRevision: null });
  assert.equal(sha(await fs.readFile(source)), sourceHash);
  const zip = unzipSync(await fs.readFile(saved)),
    xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml']);
  assert(xml.includes(validation));
  assert(xml.includes(conditional));
  assert(xml.includes('<v>9007199254740993</v>'));
  assert.deepEqual(zip['xl/styles.xml'], original['xl/styles.xml']);
  assert.deepEqual(zip['custom/opaque.bin'], original['custom/opaque.bin']);
  const proof = {
    validation: { canEdit: false, protectedRanges: validationDefinition.protectedRanges, exactXmlPreserved: true },
    conditional: {
      canEdit: true,
      evaluationReasons: conditionalDefinition.evaluationReasons,
      knownRange: 'Z10:Z11',
      exactXmlPreserved: true,
    },
    acceptedBatches: accepts,
    oneUndoRedo: true,
    unsafeUnrelatedRawPreserved: true,
    stylesAndOpaquePreserved: true,
    sourceUnchanged: true,
  };
  await fs.writeFile(path.join(run, 'protected-disjoint-readback.json'), JSON.stringify(proof, null, 2) + '\n');
});
