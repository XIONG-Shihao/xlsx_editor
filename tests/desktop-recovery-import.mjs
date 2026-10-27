import {test} from 'node:test'
import assert from 'node:assert/strict'
import {mkdtemp,mkdir,writeFile,readFile,realpath,readdir,cp,rm} from 'node:fs/promises'
import {constants} from 'node:fs'
import {tmpdir} from 'node:os'
import {join,resolve} from 'node:path'
import {createHash} from 'node:crypto'
import {unzipSync,strFromU8} from 'fflate'
import {DOMParser} from '@xmldom/xmldom'
import {SpreadsheetWorker} from '../adapters/worker-client.mjs'
import {KeptRecoveryBindings} from '../adapters/kept-recovery.mjs'
import {workbook} from '../fixtures/standalone.mjs'
import {standaloneClient} from './support/desktopClient.mjs'
const hash=bytes=>createHash('sha256').update(bytes).digest('hex')
async function witness(root){const result=[];async function visit(directory){for(const item of (await readdir(directory,{withFileTypes:true})).sort((a,b)=>a.name.localeCompare(b.name))){const path=join(directory,item.name);if(item.isDirectory())await visit(path);else if(item.isFile())result.push([path.slice(root.length),hash(await readFile(path))])}}await visit(root);return result}
test('standalone native package opens a separate copy of frozen saved-file recovery',{timeout:60000,skip:process.env.PLATFORM_TEST_STANDALONE&&process.env.PLATFORM_TEST_FROZEN_WORKBOOK?false:'opt-in: PLATFORM_TEST_STANDALONE and PLATFORM_TEST_FROZEN_WORKBOOK select a desktop package and a frozen workbook binary'},async()=>{
 const root=await mkdtemp(join(tmpdir(),'tauri-standalone-import-')),source=join(root,'electron'),target=join(root,'native'),moved=join(root,'Relocated XLSX.app'),copied=join(root,'Recovered copy.xlsx'),original=workbook()
 const prior=process.env.XLSX_EDITOR_NATIVE_BIN;let worker,client
 try{
  await mkdir(source);await writeFile(join(root,'External.xlsx'),original);const filename=await realpath(join(root,'External.xlsx'))
  process.env.XLSX_EDITOR_NATIVE_BIN=resolve(process.env.PLATFORM_TEST_FROZEN_WORKBOOK)
  worker=new SpreadsheetWorker({filename,recoveryDirectory:join(source,'recovery')});let accepted=await worker.open()
  accepted=await worker.request({operation:'apply',documentId:accepted.documentId,state:accepted.state,commands:[{type:'setCells',sheet:'Sheet1',cells:[{cell:'A1',value:42}]}]})
  assert.equal(accepted.dirty,true)
  const binding=await KeptRecoveryBindings.open(join(source,'kept-recovery.json'));await binding.pin(filename,accepted.documentId)
  await worker.close();worker=null
  const before=await witness(source)
  // This is an explicit offline copy into a fresh profile, not a second writer
  // pointed at Electron's original directory. No browser cache is required.
  await cp(source,target,{recursive:true,verbatimSymlinks:true,mode:constants.COPYFILE_FICLONE})
  await cp(resolve(process.env.PLATFORM_TEST_STANDALONE),moved,{recursive:true,verbatimSymlinks:true,mode:constants.COPYFILE_FICLONE})
  client=await standaloneClient(moved,target,{copyPath:copied})
  const list=await client.invoke('xlsx:recovery-files');assert.deepEqual(list.warnings,[])
  assert.ok(list.entries.some(item=>item.id===accepted.documentId&&item.dirty))
  const pending=await client.invoke('xlsx:open-recovery',accepted.documentId);let native=await client.invoke('xlsx:activate-open',pending.id)
  assert.equal(native.documentId,accepted.documentId);assert.equal(native.dirty,true)
  const request=(operation,fields={})=>client.invoke('xlsx:request',native.id,{operation,documentId:native.documentId,state:native.state,...fields})
  assert.equal((await request('read',{sheet:'Sheet1',range:'A1:C1'})).cells.A1.value,42)
  native=await request('undo');assert.equal((await request('read',{sheet:'Sheet1',range:'A1:C1'})).cells.A1.value,10)
  native=await request('redo');assert.equal((await request('read',{sheet:'Sheet1',range:'A1:C1'})).cells.A1.value,42)
  assert.equal(await client.invoke('xlsx:copy',native.id),true)
  const parts=unzipSync(await readFile(copied)),ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',xml=new DOMParser().parseFromString(strFromU8(parts['xl/worksheets/sheet1.xml']),'application/xml'),cells=new Map(Array.from(xml.getElementsByTagNameNS(ns,'c')).map(node=>[node.getAttribute('r'),node]))
  assert.equal(Number(cells.get('A1').getElementsByTagNameNS(ns,'v')[0].textContent),42);assert.equal(Number(cells.get('C1').getElementsByTagNameNS(ns,'v')[0].textContent),62)
  assert.equal(cells.get('C1').getElementsByTagNameNS(ns,'f')[0].textContent,'A1+B1');assert.deepEqual(parts['custom/preserved.bin'],new Uint8Array([0,1,2,255]))
  await client.stop();client=null
  assert.deepEqual(await witness(source),before,'Using native recovery never writes back into the frozen source profile')
  assert.deepEqual(new Uint8Array(await readFile(filename)),original,'Conflict-copy publication leaves the external workbook exact')
 }finally{try{await worker?.close();await client?.stop()}finally{if(prior===undefined)delete process.env.XLSX_EDITOR_NATIVE_BIN;else process.env.XLSX_EDITOR_NATIVE_BIN=prior;await rm(root,{recursive:true,force:true})}}
})
