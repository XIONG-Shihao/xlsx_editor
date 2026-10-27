import {spawn} from 'node:child_process'
import {join,resolve} from 'node:path'
import {readFile} from 'node:fs/promises'
import {build} from 'esbuild'
import {fileURLToPath} from 'node:url'
import {processOwner} from '../../desktop/scripts/support/processOwner.mjs'
const product=fileURLToPath(new URL('../..',import.meta.url)),framing=join(product,'.build/tests/desktop-framing.mjs')
await build({entryPoints:[join(product,'desktop/backend/host/framing.ts')],bundle:true,platform:'node',format:'esm',outfile:framing})
const {frame,Frames}=await import(framing)
export async function standaloneClient(bundle,profile,{copyPath}={}){
 const resources=join(resolve(bundle),'Contents/Resources')
 const manifest=JSON.parse(await readFile(join(resources,'standalone-package.json'),'utf8'))
 if(manifest.javascriptRuntime?.kind!=='bun'||manifest.javascriptRuntime.executable!=='backend/bun')throw Error('Standalone verification requires the contained Bun runtime')
 const child=spawn(join(resources,'backend/bun'),['--no-install',join(resources,'app/desktop/out/main/index.js')],{cwd:profile,stdio:['pipe','pipe','pipe'],env:{...process.env,PATH:'/usr/bin:/bin',XLSX_EDITOR_USER_DATA:profile,XLSX_EDITOR_RESOURCES:resources,XLSX_EDITOR_APP_ROOT:join(resources,'app'),XLSX_EDITOR_PACKAGED:'1'}})
 const owner=processOwner(child.pid),pending=new Map();let serial=0,logs=''
 const exited=new Promise(resolve=>child.once('exit',(code,signal)=>resolve({code,signal})))
 child.stderr.on('data',bytes=>{logs+=bytes})
 child.stdout.pipe(new Frames()).on('data',message=>{
  if(message.kind==='native'){
   const result=message.action==='messageDialog'?{response:3}:message.action==='saveDialog'?{canceled:!copyPath,filePath:copyPath??''}:null
   child.stdin.write(frame({kind:'native-result',id:message.id,result}))
  }
  const waiter=pending.get(message.id);if(waiter){pending.delete(message.id);message.error?waiter.reject(Error(message.error)):waiter.resolve(message.result)}
 })
 child.on('exit',code=>{for(const waiter of pending.values())waiter.reject(Error('Standalone coordinator exited ('+code+'): '+logs));pending.clear()})
 const invoke=(channel,...args)=>new Promise((resolve,reject)=>{const id=String(++serial);pending.set(id,{resolve,reject});child.stdin.write(frame({kind:'invoke',id,channel,args}))})
 await invoke('standalone:initialize')
 return{invoke,logs:()=>logs,async stop(){try{await invoke('xlsx:close-all');child.stdin.end();const result=await exited;if(result.code!==0)throw Error('Standalone EOF failed: '+logs)}finally{if(child.exitCode===null&&child.signalCode===null){child.kill('SIGTERM');await exited}await owner.closed()}}}
}
