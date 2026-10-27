/** Native print host for the same accepted page composer. Large HTML/PDF bytes
 * stay in private temporary files, never a base64 native-action frame. */
import fs from 'node:fs/promises'
import {constants} from 'node:fs'
import path from 'node:path'
import {tmpdir} from 'node:os'
import {randomUUID,createHash} from 'node:crypto'
import {hostAction} from './host'
import {validatePdfDocument} from '../../../adapters/print-document.mjs'
import {printOverflowInspectionSource} from '../../../core/print-layout.mjs'

const MAX_BYTES=16*1024*1024
export function createNativePrintHost({native=hostAction}={}){
 async function prepare(input,{signal,checkCurrent=()=>{},onProgress=()=>{}}={}){
  const document=validatePdfDocument(input),id=randomUUID(),root=await fs.mkdtemp(path.join(tmpdir(),`platform-tauri-print-${id}-`))
  await fs.chmod(root,0o700)
  const check=()=>{signal?.throwIfAborted();checkCurrent()}
  let closePromise,closed=false,handedOff=false
  const close=()=>closePromise??=(async()=>{
   // A timeout is not evidence that WebKit/printing stopped. A busy close
   // preserves staging instead of removing a live producer's files.
   const result=await native('workbookPrintClose',{id})
   if(!result?.closed)throw Error('unavailable: Native print did not release its lease; staging preserved')
   closed=true;await fs.rm(root,{recursive:true,force:true})
  })().catch(error=>{closePromise=undefined;throw error})
  const abort=()=>{if(!handedOff)void close().catch(()=>{})}
  signal?.addEventListener('abort',abort,{once:true})
  try{
   check();await fs.writeFile(path.join(root,'document.html'),document.html,{flag:'wx',mode:0o600});check()
   onProgress('rendering-pages');check()
   const prepared=await native('workbookPrintPrepare',{id,root,sha256:createHash('sha256').update(document.html).digest('hex'),inspection:printOverflowInspectionSource(),pageWidthPt:document.pageWidthPt,pageHeightPt:document.pageHeightPt,pageCount:document.pageCount})
   check();if(prepared?.prepared!==true||prepared.pageCount!==document.pageCount)throw Error('validation: Native print readiness differs from accepted pages')
   return {id,root,document,check,close,dispatch:()=>{handedOff=true;signal?.removeEventListener('abort',abort)},dispose:async()=>{signal?.removeEventListener('abort',abort);if(!closed)await close()}}
  }catch(error){
   signal?.removeEventListener('abort',abort)
   try{await close()}catch(cleanup){
    // An absent lease before admission permits removal; an ambiguous active
    // lease/timeout does not. Never restart a renderer to hide that failure.
    if(String(cleanup).includes('lease is missing or retired'))await fs.rm(root,{recursive:true,force:true})
   }
   throw error
  }
 }
 return {
  async pdf(input,options={}){
   const owner=await prepare(input,options)
   try{
    owner.check();options.onProgress?.('creating-pdf')
    const rendered=await native('workbookPrintPdf',{id:owner.id});owner.check()
    if(rendered?.rendered!==true)throw Error('unavailable: Native PDF did not finish')
    const result=await native('workbookPrintInspect',{id:owner.id});owner.check()
    if(result?.pageCount!==owner.document.pageCount||result.pageWidthPt!==owner.document.pageWidthPt||result.pageHeightPt!==owner.document.pageHeightPt)throw Error('validation: Native PDF differs from accepted layout')
    const sourcePath=path.join(owner.root,'pages.pdf'),file=await fs.open(sourcePath,constants.O_RDONLY|constants.O_NOFOLLOW)
    try{
     const metadata=await file.stat(),magic=Buffer.alloc(5);await file.read(magic,0,5,0)
     if(!metadata.isFile()||metadata.nlink!==1||metadata.size<8||metadata.size>MAX_BYTES||metadata.size!==result.bytes||magic.toString()!=='%PDF-')throw Error('validation: Invalid bounded native PDF file')
    }finally{await file.close()}
    owner.check();options.onProgress?.('pdf-ready')
    return {sourcePath,pageCount:result.pageCount,dispose:owner.dispose}
   }catch(error){await owner.dispose();throw error}
  },
  async printer(input,options={}){
   const owner=await prepare(input,options)
   try{
    options.onProgress?.('choosing-printer');owner.check();await options.authorizeDispatch?.();owner.check()
    options.beforeDispatch?.();owner.dispatch()
    const result=await native('workbookPrintDispatch',{id:owner.id})
    if(result?.pageCount!==owner.document.pageCount||!result.submitted&&!result.cancelled)throw Error('unavailable: Native printer returned no terminal result')
    return result
   }finally{await owner.dispose()}
  }
 }
}
