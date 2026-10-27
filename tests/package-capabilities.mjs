import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook, readWorkbookMetadata, editWorkbook } from '../core/workbook-package.mjs';
import { packageReadOnlyReasons } from '../core/package-capabilities.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
test('empty and explicitly inactive workbook protection admit ordinary edits without dropping the original record; active or unknown records stay protected', async (t) => {
  for (const markup of [
    '<workbookProtection/>',
    '<workbookProtection lockStructure="false" lockWindows="0" lockRevision="false"/>',
    '<workbookProtection lockStructure="1"/>',
    '<workbookProtection workbookPassword="ABCD"/>',
    '<workbookProtection lockStructure="unknown"/>',
    '<workbookProtection extra="false"/>',
    '<workbookProtection/><workbookProtection/>',
  ]) {
    const parts = unzipSync(createWorkbook());
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']).replace('<sheets>', markup + '<sheets>'));
    const bytes = Buffer.from(zipSync(parts)),
      active = ![
        '<workbookProtection/>',
        '<workbookProtection lockStructure="false" lockWindows="0" lockRevision="false"/>',
      ].includes(markup);
    assert.equal(readWorkbook(bytes).readOnly, active);
    assert.equal(readWorkbookMetadata(bytes).view.readOnly, active);
    const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-empty-protection-')),
      source = path.join(root, 'source.xlsx');
    await fs.writeFile(source, bytes);
    const book = new IndexedWorkbook(native, () => assert.fail('No fallback'));
    try {
      await book.open(source);
      assert.equal(book.view('0').readOnly, active);
      if (active) await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]), /read-only/);
      else {
        const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 7 }]),
          saved = await book.export(next.view.storage.state);
        assert(strFromU8(unzipSync(await fs.readFile(saved.path))['xl/workbook.xml']).includes(markup));
        assert.equal((await book.read(next.view.storage.state, 'Sheet1', 'A1')).cells.A1.value, 7);
      }
      assert.deepEqual(await fs.readFile(source), bytes);
    } finally {
      await book.close();
      await fs.rm(root, { recursive: true, force: true });
    }
  }
});
const types = 'http://schemas.openxmlformats.org/package/2006/content-types',
  rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
function packageFor({ mime, relation, namespace = rels, target = 'opaque.dat', targetMode } = {}) {
  const parts = unzipSync(createWorkbook());
  parts['custom/opaque.dat'] = new Uint8Array([0, 255, 7, 23]);
  if (mime) {
    const before = strFromU8(parts['[Content_Types].xml']);
    parts['[Content_Types].xml'] = strToU8(
      /\.macroEnabled\./.test(mime)
        ? before.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', mime)
        : before.replace('</Types>', `<Override PartName="/custom/opaque.dat" ContentType="${mime}"/></Types>`)
    );
  }
  if (relation)
    parts['custom/_rels/source.xml.rels'] = strToU8(
      `<r:Relationships xmlns:r="${namespace}"><r:Relationship Id="opaque" Type="${relation}" Target="${target}"${targetMode ? ' TargetMode="' + targetMode + '"' : ''}/></r:Relationships>`
    );
  return parts;
}
// Macros are carried untouched and do not guard; controls and package signatures do.
const guards = [
  { mime: 'application/vnd.ms-office.activeX+xml' },
  { mime: 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml' },
  { relation: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/control' },
  { relation: 'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary' },
  { relation: 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature' },
  {
    relation: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/control',
    target: 'https://example.invalid/control',
    targetMode: 'External',
  },
];
const macroEntries = [
  { mime: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' },
  { mime: 'application/vnd.ms-excel.template.macroEnabled.main+xml' },
  { mime: 'application/vnd.ms-excel.addin.macroEnabled.main+xml' },
  { mime: 'application/vnd.ms-office.vbaProject' },
  { relation: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject' },
  { relation: 'http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile' },
  { relation: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', target: '../../renamed.dat' },
];
test('macro types and relationships do not make a workbook read-only', () => {
  for (const fixture of macroEntries) {
    const parts = packageFor(fixture);
    assert.deepEqual(packageReadOnlyReasons(parts), []);
    assert.equal(readWorkbook(zipSync(parts)).readOnly, false);
    assert.equal(readWorkbookMetadata(zipSync(parts)).view.readOnly, false);
  }
});
test('renamed ActiveX/signature graph forces both readers read-only without rewriting package bytes', () => {
  for (const fixture of guards) {
    const parts = packageFor(fixture),
      bytes = zipSync(parts),
      before = hash(bytes);
    assert(!Object.keys(parts).some((name) => /activeX|_xmlsignatures/.test(name)));
    assert(packageReadOnlyReasons(parts).some((reason) => /ActiveX/.test(reason)));
    assert.equal(readWorkbook(bytes).readOnly, true);
    assert.equal(readWorkbookMetadata(bytes).view.readOnly, true);
    assert.throws(() => editWorkbook(bytes, [{ sheet: 'Sheet1', cell: 'A1', value: 'refused' }]), /Read-only/);
    assert.equal(hash(bytes), before);
    assert.deepEqual(unzipSync(bytes)['custom/opaque.dat'], parts['custom/opaque.dat']);
  }
});
test('malformed/ambiguous graph refuses editing while conventional ordinary OOXML remains editable', () => {
  const ordinary = packageFor({
    relation: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  });
  assert.deepEqual(packageReadOnlyReasons(ordinary), []);
  assert.equal(readWorkbook(zipSync(ordinary)).readOnly, false);
  const malformed = [];
  malformed.push(
    packageFor({
      relation: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
      namespace: 'urn:wrong',
    })
  );
  for (const mime of [
    ' application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    'application/vnd.ms-office.vbaProject ',
    'application/vnd.ms-office. vbaProject',
  ])
    malformed.push(packageFor({ mime }));
  for (const relation of [
    ' http://schemas.microsoft.com/office/2006/relationships/vbaProject',
    'http://schemas.microsoft.com/office/2006/relationships/vbaProject ',
  ])
    malformed.push(packageFor({ relation }));
  for (const xml of [
    `<Types xmlns="urn:wrong"/>`,
    `<Types xmlns="${types}"><Override PartName="/xl/workbook.xml" ContentType="a"/><Override PartName="/xl/workbook.xml" ContentType="b"/></Types>`,
    `<Types xmlns="${types}"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"><extra/></Override></Types>`,
    `<!DOCTYPE Types [<!ENTITY x 'unsafe'>]><Types xmlns="${types}"/>`,
  ]) {
    const parts = packageFor();
    parts['[Content_Types].xml'] = strToU8(xml);
    malformed.push(parts);
  }
  for (const xml of [
    `<Relationships xmlns="${rels}"><Relationship Id="same" Type="ordinary" Target="one"/><Relationship Id="same" Type="ordinary" Target="two"/></Relationships>`,
    `<Relationships xmlns="${rels}"><Relationship Id="x" Target="renamed.dat"/></Relationships>`,
    `<Relationships xmlns="${rels}">unexpected text</Relationships>`,
    `<Relationships xmlns="${rels}"><![CDATA[unexpected text]]></Relationships>`,
  ]) {
    const parts = packageFor();
    parts['custom/_rels/item.xml.rels'] = strToU8(xml);
    malformed.push(parts);
  }
  for (const parts of malformed) {
    assert(packageReadOnlyReasons(parts).some((reason) => /graph is invalid/.test(reason)));
    assert.equal(readWorkbook(zipSync(parts)).readOnly, true);
    assert.throws(() => editWorkbook(zipSync(parts), [{ sheet: 'Sheet1', cell: 'A1', value: 'refused' }]), /Read-only/);
  }
});
test('real indexed metadata reader propagates renamed ActiveX and signature graph admission and leaves source intact', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-macro-admission-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  for (const fixture of [guards[0], guards[2], guards[4]]) {
    const bytes = zipSync(packageFor(fixture)),
      source = path.join(directory, 'source-' + Math.random() + '.xlsx');
    await fs.writeFile(source, bytes);
    const book = new IndexedWorkbook(
      native,
      () => assert.fail('Readonly adapter must not write'),
      () => assert.fail('No formula calculation required')
    );
    try {
      const view = await book.open(source);
      assert.equal(view.readOnly, true);
      await assert.rejects(
        book.apply(view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', value: 'refused' }]),
        /read-only/
      );
      assert.equal(hash(await fs.readFile(source)), hash(bytes));
    } finally {
      await book.close();
    }
  }
});
