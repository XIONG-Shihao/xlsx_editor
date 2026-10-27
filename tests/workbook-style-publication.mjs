import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,readFile,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {resolveWorkbookStyles} from '../core/style-publication.mjs';
import {createWorkbook,readWorkbook} from '../core/workbook-package.mjs';
import {IndexedWorkbook} from '../core/indexed-session.mjs';
import {nativeWorkbookOperationAsync,closeNativeWorkbookWorker} from '../core/native-runtime.mjs';

const xml='<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><name val="Example"/></font></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>';
test('pure style planner keeps source XML, resolves unique records and rejects invalid authority',()=>{
  const before=xml;
  const result=resolveWorkbookStyles(xml,[{id:2,format:{styleId:0,bold:true}},{id:4,format:{styleId:0,bold:true}}]);
  assert.equal(xml,before);
  assert.deepEqual(result.styleIds,[{id:2,styleId:1},{id:4,styleId:1}]);
  assert.match(result.stylesXml,/<b\/>/);
  for(const signatures of [[{id:1,format:{styleId:1}}],[{id:1,format:{styleId:0}},{id:1,format:{styleId:0}}],[{id:1,format:{styleId:0,unknown:true}}]])
    assert.throws(()=>resolveWorkbookStyles(xml,signatures),/validation/);
  for(const invalid of ['<styleSheet/>',xml.replace('<styleSheet','<!DOCTYPE styleSheet><styleSheet'),xml.replace('</styleSheet>','')])
    assert.throws(()=>resolveWorkbookStyles(invalid,[]),/validation/);
});

test('shipping indexed adapter uses the shared resolver for edit, formatting, export and historical states',async()=>{
  const directory=await mkdtemp(path.join(tmpdir(),'platform-style-publication-'));
  const source=path.join(directory,'source.xlsx'),original=createWorkbook();
  let book;
  try {
    await writeFile(source,original);
    book=new IndexedWorkbook(nativeWorkbookOperationAsync,()=>{throw Error('Unexpected fallback')});
    let view=await book.open(source);
    const sheet=view.sheets[0].name;
    let result=await book.apply(view.storage.state,[{sheet,cell:'A1',value:42}]);
    const previousState=result.view.storage.state;
    result=await book.apply(result.view.storage.state,[{type:'formatRange',sheet,range:'A1:B1',format:{bold:true,fillColor:'AABBCC',numberFormat:'decimal1'}}]);
    const exported=await book.export(result.view.storage.state);
    const readback=readWorkbook(await readFile(exported.path));
    assert.equal(readback.sheets[0].cells.A1.value,42);
    assert.equal(readback.sheets[0].cells.A1.bold,true);
    assert.equal(readback.sheets[0].cells.A1.fillColor,'AABBCC');
    assert.equal(readback.sheets[0].cells.A1.format,'0.0');
    assert.equal(readback.sheets[0].cells.B1.bold,true);
    const undone=readWorkbook(await readFile((await book.export(previousState)).path));
    assert.equal(undone.sheets[0].cells.A1.value,42);
    assert.notEqual(undone.sheets[0].cells.A1.bold,true);
    assert.deepEqual(await readFile(source),Buffer.from(original));
  } finally {
    await book?.close();closeNativeWorkbookWorker();await rm(directory,{recursive:true,force:true});
  }
});
