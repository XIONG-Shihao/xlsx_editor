import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbookMetadata } from '../core/workbook-package.mjs';
import { indexedNamePlan } from '../core/indexed-names.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p = 'http://schemas.openxmlformats.org/package/2006/relationships';
const prior = "'Sheet1'!$A$1:$A$2",
  next = "'Sheet1'!$C$3:$E$9",
  hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
function fixture(kind) {
  const parts = unzipSync(createWorkbook());
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml']).replace(
      '</sheets>',
      `</sheets><definedNames><definedName name="Inputs">${prior}</definedName></definedNames>`
    )
  );
  if (kind === 'renamed-table') {
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${ns}" xmlns:r="${r}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="table"/></tableParts></worksheet>`
    );
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      `<Relationships xmlns="${p}"><Relationship Id="table" Type="${r}/table" Target="/custom/table.xml"/></Relationships>`
    );
    parts['custom/table.xml'] = strToU8(
      `<table xmlns="${ns}" id="1" name="Other" displayName="Other" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Amount"><calculatedColumnFormula>SUM(Inputs)</calculatedColumnFormula></tableColumn></tableColumns></table>`
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        '<Override PartName="/custom/table.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'
      )
    );
  } else if (kind === 'renamed-worksheet') {
    parts['custom/sheet.xml'] = strToU8(
      `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><f>SUM(Inputs)</f><v>999</v></c></row></sheetData></worksheet>`
    );
    delete parts['xl/worksheets/sheet1.xml'];
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace('worksheets/sheet1.xml', '/custom/sheet.xml')
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace('/xl/worksheets/sheet1.xml', '/custom/sheet.xml')
    );
  } else if (kind === 'unknown-reference') {
    parts['custom/_rels/item.xml.rels'] = strToU8(
      `<Relationships xmlns="${p}"><Relationship Id="dependency" Type="urn:vendor:formula-engine" Target="linked.xml"/></Relationships>`
    );
    parts['custom/item.xml'] = strToU8('<opaque/>');
    parts['custom/linked.xml'] = strToU8('<dependency name="Inputs"/>');
  } else if (kind === 'linked-properties') {
    parts['_rels/.rels'] = strToU8(
      strFromU8(parts['_rels/.rels']).replace(
        '</Relationships>',
        `<Relationship Id="linked" Type="${r}/custom-properties" Target="docProps/custom.xml"/></Relationships>`
      )
    );
    parts['docProps/custom.xml'] = strToU8(
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property name="Linked" pid="2" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" linkTarget="Inputs"/></Properties>'
    );
  }
  return parts;
}
for (const kind of ['renamed-table', 'renamed-worksheet', 'unknown-reference', 'linked-properties'])
  test(`${kind} cannot bypass either public or independent native name dependency protection`, async (t) => {
    const parts = fixture(kind),
      bytes = zipSync(parts),
      view = readWorkbookMetadata(bytes).view;
    assert.equal(view.readOnly, false);
    if (kind === 'renamed-table') assert.equal(view.tables[0].path, 'custom/table.xml');
    if (kind === 'renamed-worksheet') assert.equal(view.sheets[0].path, 'custom/sheet.xml');
    if (kind !== 'renamed-worksheet')
      assert.throws(
        () =>
          indexedNamePlan(bytes, view, {
            operation: 'updateName',
            sheet: 'Sheet1',
            name: 'Inputs',
            range: 'C3:E9',
            expectedNameFormula: prior,
          }),
        /referenced|reference/
      );
    else
      assert.equal(
        indexedNamePlan(bytes, view, {
          operation: 'updateName',
          sheet: 'Sheet1',
          name: 'Inputs',
          range: 'C3:E9',
          expectedNameFormula: prior,
        }).kind,
        'updateName'
      );
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), kind + '-'));
    t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const source = path.join(directory, 'source.xlsx');
    await fs.writeFile(source, bytes);
    await native({ action: 'workbook', operation: 'index', root: directory });
    const plan = {
      kind: 'updateName',
      name: 'Inputs',
      scope: null,
      priorFormula: prior,
      formula: next,
      replacements: { 'xl/workbook.xml': strFromU8(parts['xl/workbook.xml']).replace(prior, next) },
      removed: [],
      copies: {},
      fragments: {},
    };
    if (kind === 'renamed-worksheet') {
      plan.kind = 'removeName';
      delete plan.formula;
      plan.replacements['xl/workbook.xml'] = strFromU8(parts['xl/workbook.xml']).replace(
        '<definedName name="Inputs">' + prior + '</definedName>',
        ''
      );
    }
    await fs.writeFile(path.join(directory, 'package-metadata.json'), JSON.stringify(plan));
    const sentinel = Buffer.from('previous complete output');
    await fs.writeFile(path.join(directory, 'package.xlsx'), sentinel);
    await assert.rejects(
      native({ action: 'workbook', operation: 'package-edit', root: directory, version: 0, source: 'source.xlsx' }),
      /reference|Referenced|relationship/
    );
    assert.equal(hash(await fs.readFile(source)), hash(bytes));
    assert.deepEqual(await fs.readFile(path.join(directory, 'package.xlsx')), sentinel);
  });
