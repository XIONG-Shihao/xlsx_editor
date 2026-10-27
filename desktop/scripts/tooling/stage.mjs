// Running build steps and replacing a finished bundle, for this product only.
import {cp,mkdir,rename,rm} from 'node:fs/promises'
import {constants} from 'node:fs'
import {basename,dirname,join,resolve} from 'node:path'
import {spawn} from 'node:child_process'
import {fileURLToPath} from 'node:url'

/** The product folder, wherever it is. */
export const productRoot=resolve(fileURLToPath(new URL('../../..',import.meta.url)))
export function run(command,args,options={}){
 return new Promise((resolve,reject)=>{
  const child=spawn(command,args,{cwd:productRoot,stdio:'inherit',...options})
  child.once('error',reject)
  child.once('exit',(code,signal)=>code===0?resolve():reject(Error(`${command} failed (${signal||code})`)))
 })
}
const copy=(source,target)=>cp(source,target,{recursive:true,dereference:false,verbatimSymlinks:true,mode:constants.COPYFILE_FICLONE})
// Copy and verify before replacing. The old bundle exists only during this
// transaction for rollback, never as a retained history/version directory.
export async function publishBundle(source,destination,verify,checkIdle=()=>{}){
 const directory=join(dirname(destination),`.${basename(destination)}.replacement-${process.pid}`)
 const candidate=join(directory,basename(destination)),previous=join(directory,'previous')
 await mkdir(directory,{recursive:false})
 let movedPrevious=false
 try{
  await copy(source,candidate)
  await verify(candidate)
  await checkIdle()
  try{await rename(destination,previous);movedPrevious=true}
  catch(error){if(error.code!=='ENOENT')throw error}
  try{await rename(candidate,destination)}
  catch(error){
   if(movedPrevious){
    try{await rename(previous,destination);movedPrevious=false}
    catch(restoreError){throw new AggregateError([error,restoreError],`Replacement failed; recover the prior bundle at ${previous}`)}
   }
   throw error
  }
  movedPrevious=false
 }finally{
  // Retain a prior bundle only if rollback itself failed, for recovery.
  if(!movedPrevious)await rm(directory,{recursive:true,force:true})
 }
}
