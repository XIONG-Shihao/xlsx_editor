import {bunVersion,sourceBun} from './tooling/bun.mjs'
import {productRoot} from './tooling/stage.mjs'
import {spawn} from 'node:child_process'
import {mkdtemp,rm,readFile,mkdir,writeFile} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join,resolve} from 'node:path'
import assert from 'node:assert/strict'
import {build} from 'esbuild'
import {readWorkbook} from '../../core/workbook-package.mjs'
import {processOwner} from './support/processOwner.mjs'
await build({entryPoints:[join(productRoot,'desktop/backend/host/framing.ts')],bundle:true,platform:'node',format:'esm',outfile:join(productRoot,'desktop/.build/framing.mjs')})
const {frame,Frames}=await import('../.build/framing.mjs')
const profile=await mkdtemp(join(tmpdir(),'tauri-standalone-coordinator-'))
const delivery=join(profile,'published');await mkdir(delivery)
const savedFile=join(delivery,'verified.xlsx'),copyFile=join(delivery,'copy.xlsx');let saves=0,savedBytes
const resources=process.argv[2]?join(resolve(process.argv[2]),'Contents/Resources'):null
const manifest=resources?JSON.parse(await readFile(join(resources,'standalone-package.json'),'utf8')):null
if(resources)assert.deepEqual(manifest?.javascriptRuntime,{kind:'bun',version:bunVersion,executable:'backend/bun'})
const runtime='bun',executable=resources?join(resources,'backend/bun'):sourceBun
let sequence=0,logs='',approved=false,decision=1,child,exited,processes;const memory=[];const pending=new Map(),events=[]
function start(){
 child=spawn(executable,[...(runtime==='bun'?['--no-install']:[]),resources?join(resources,'app/desktop/out/main/index.js'):join(productRoot,'desktop/out/main/index.js')],{cwd:profile,stdio:['pipe','pipe','pipe'],env:{...process.env,PATH:'/usr/bin:/bin',XLSX_EDITOR_USER_DATA:profile,...(resources?{XLSX_EDITOR_RESOURCES:resources,XLSX_EDITOR_APP_ROOT:join(resources,'app'),XLSX_EDITOR_PACKAGED:'1'}:{}),XLSX_EDITOR_NATIVE_BIN:resources?join(resources,'native/workbook-engine'):join(productRoot,'native/bin/darwin-arm64/workbook-engine')}})
 const owner=child;processes=processOwner(child.pid,{role:(_row,isRoot)=>isRoot?'xlsx-coordinator':'helper'});approved=false
 exited=new Promise(resolve=>owner.once('exit',(code,signal)=>resolve({code,signal})))
 owner.stderr.on('data',data=>logs+=data)
 owner.stdout.pipe(new Frames()).on('data',message=>{
  if(message.kind==='native')owner.stdin.write(frame({kind:'native-result',id:message.id,result:message.action==='messageDialog'?{response:decision}:message.action==='saveDialog'?{canceled:false,filePath:++saves===1?savedFile:copyFile}:null}))
  if(message.kind==='event'){events.push(message);if(message.channel==='standalone:exit-approved')approved=true}
  const request=pending.get(message.id);if(request){pending.delete(message.id);message.error?request.reject(Error(message.error)):request.resolve(message.result)}
 })
 owner.on('exit',()=>{for(const request of pending.values())request.reject(Error('Coordinator exited: '+logs));pending.clear()})
}
const invoke=(channel,...args)=>new Promise((resolve,reject)=>{const id=String(++sequence),timer=setTimeout(()=>{pending.delete(id);reject(Error('Coordinator test request deadline: '+channel))},30000);pending.set(id,{resolve:value=>{clearTimeout(timer);resolve(value)},reject:error=>{clearTimeout(timer);reject(error)}});child.stdin.write(frame({kind:'invoke',id,channel,args}))})
start()
try{
 const info=await invoke('standalone:initialize');assert.equal(info.product,'standalone-xlsx-tauri');assert.deepEqual(info.javascriptRuntime,{kind:'bun',version:bunVersion})
 const bridge=await readFile(join(productRoot,'standalone/bridge.mjs'),'utf8')
 const endpoints=[...bridge.matchAll(/invoke\('([^']+)'/g)].map(match=>'xlsx:'+match[1])
 assert.deepEqual(info.methods,endpoints.toSorted(),'Every component bridge operation is present; no substitute reduced API')
 assert.equal(await invoke('xlsx:current'),null)
 const opened=await invoke('xlsx:new');assert.ok(opened.id)
 const active=await invoke('xlsx:activate-open',opened.id,false);assert.equal(active.unsaved,true)
 assert.equal((await invoke('xlsx:documents')).documents.length,1)
 const changed=await invoke('xlsx:request',active.id,{operation:'apply',documentId:active.documentId,state:active.state,commands:[{type:'setCells',sheet:'Sheet1',cells:[{cell:'A1',value:15},{cell:'B1',formula:'A1*3'}]}]})
 assert.equal(changed.dirty,true)
 const cells=await invoke('xlsx:request',active.id,{operation:'read',state:changed.state,sheet:'Sheet1',range:'A1:B1'})
 assert.equal(cells.cells.A1.value,15);assert.equal(cells.cells.B1.calculation.value,45)
 const saved=await invoke('xlsx:save',active.id);assert.equal(saved.dirty,false);assert.equal(saved.unsaved,false)
 savedBytes=await readFile(savedFile);const book=readWorkbook(savedBytes);assert.equal(book.sheets[0].cells.A1.value,15);assert.equal(book.sheets[0].cells.B1.cachedValue,45)
 assert.equal(await invoke('xlsx:copy',active.id),true);assert.equal(readWorkbook(await readFile(copyFile)).sheets[0].cells.B1.cachedValue,45);assert.deepEqual(await readFile(savedFile),savedBytes,'Save Copy leaves source exact')
 const discarded=await invoke('xlsx:request',active.id,{operation:'discard'})
 assert.equal(discarded.dirty,false)
 const second=await invoke('xlsx:new');await invoke('xlsx:activate-open',second.id)
 const selected=await invoke('xlsx:activate-tab',active.id);assert.equal(selected.id,active.id)
 const beforeEdit=(await invoke('xlsx:request',active.id,{operation:'read',documentId:selected.documentId,state:selected.state,sheet:'Sheet1',range:'A1:A1'})).cells.A1?.value
 const commands=[{type:'setCells',sheet:'Sheet1',cells:[{cell:'A1',value:42}]}]
 await assert.rejects(invoke('xlsx:request',active.id,{operation:'apply',commands}),/stale/)
 const dirty=await invoke('xlsx:request',active.id,{operation:'apply',documentId:selected.documentId,state:selected.state,commands})
 assert.equal(dirty.dirty,true)
 await invoke('standalone:close');const cancelled=await invoke('xlsx:close-all')
 assert.equal(cancelled.cancelled,true);assert.equal(approved,false)
 assert.equal((await invoke('xlsx:documents')).documents.length,2)
 assert.ok(events.some(event=>event.channel==='standalone:exit-cancelled'),'Cancel permits a subsequent native Quit')
 decision=3;const kept=await invoke('xlsx:close-tab',active.id)
 assert.equal(kept.closed,true);assert.equal((await invoke('xlsx:recovery-files')).entries.length,1)
 assert.equal((await invoke('xlsx:documents')).documents.length,1)
 await invoke('xlsx:close-all')
 assert.equal(approved,true)
 memory.push(processes.sample('before-cold-relaunch'));assert.ok(processes.live().every(row=>!/(?:^|\s|\/)node(?:\s|$)/.test(row.identity)),'No owned Node process');
 child.stdin.end();const exit=await exited;await processes.closed();assert.equal(exit.code,0,logs)
 // Reopen only after the preceding owner and all of its workers retired.
 // This is a full coordinator process lifetime, not a same-process refresh.
 start();await invoke('standalone:initialize');assert.equal(await invoke('xlsx:current'),null)
 const recovery=(await invoke('xlsx:recovery-files')).entries;assert.equal(recovery.length,1)
 const prepared=await invoke('xlsx:open-recovery',recovery[0].id)
 let recovered=await invoke('xlsx:activate-open',prepared.id)
 const read=()=>invoke('xlsx:request',recovered.id,{operation:'read',documentId:recovered.documentId,state:recovered.state,sheet:'Sheet1',range:'A1:A1'})
 assert.equal((await read()).cells.A1.value,42)
 recovered=await invoke('xlsx:request',recovered.id,{operation:'undo',documentId:recovered.documentId,state:recovered.state})
 assert.equal((await read()).cells.A1?.value,beforeEdit)
 recovered=await invoke('xlsx:request',recovered.id,{operation:'redo',documentId:recovered.documentId,state:recovered.state})
 assert.equal((await read()).cells.A1.value,42)
 await invoke('xlsx:request',recovered.id,{operation:'discard'});await invoke('xlsx:close-all')
 assert.equal(approved,true);child.stdin.end();assert.equal((await exited).code,0,logs);await processes.closed()
 assert.deepEqual(await readFile(savedFile),savedBytes,'Kept/recovered edits and final Discard leave the saved source exact')
 const report={passed:true,packaged:!!resources,runtime,memory,methods:info.methods.length,checks:['complete frozen API','empty standalone home','new workbook','native owner edit/read/formula','save-as and copy with independent archive readback (simulated dialog)','discard','independent tabs','native Quit cancellation reset','retained recovery','normal owner close and coordinator EOF','cold process recovery with undo/redo'],events:events.length};console.log(JSON.stringify(report))
}finally{
 if(child.exitCode===null&&child.signalCode===null){child.kill('SIGTERM');await Promise.race([exited,new Promise(resolve=>setTimeout(resolve,5000))]);if(child.exitCode===null&&child.signalCode===null){child.kill('SIGKILL');await exited}}
 await rm(profile,{recursive:true,force:true})
}
