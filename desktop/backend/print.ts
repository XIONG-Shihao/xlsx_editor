/** Same accepted composer and bounded private native PDF/print host. */
import {readFile} from 'node:fs/promises'
import {createNativePrintHost} from './host/nativePrint.mjs'
interface PrintOptions {signal?:AbortSignal;checkCurrent?:()=>void;onProgress?:(phase:string)=>void;authorizeDispatch?:()=>Promise<void>;beforeDispatch?:()=>void;parentWindow?:unknown}
export function createWorkbookPdfRenderer(_ports:unknown){
  const host=createNativePrintHost()
  return async(input:unknown,options:PrintOptions)=>{
    const result=await host.pdf(input,options)
    try{return {bytes:await readFile(result.sourcePath),pageCount:result.pageCount}}
    finally{await result.dispose()}
  }
}
export function createWorkbookNativePrinter(_ports:unknown){
  const host=createNativePrintHost()
  return (input:unknown,options:PrintOptions)=>host.printer(input,options)
}
