// This product's pinned Bun (desktop/bun-runtime.json): downloaded once into
// .build/runtimes/bun after its digest is checked, and verified when packaged.
import {createHash,randomUUID} from 'node:crypto'
import {createReadStream,createWriteStream,readFileSync,realpathSync} from 'node:fs'
import {chmod,copyFile,lstat,mkdir,readFile,rename,rm,writeFile} from 'node:fs/promises'
import {execFileSync} from 'node:child_process'
import {join} from 'node:path'
import {fileURLToPath} from 'node:url'
import {Readable,Transform} from 'node:stream'
import {pipeline} from 'node:stream/promises'
import assert from 'node:assert/strict'
import {productRoot} from './stage.mjs'

export const bunDefinition=JSON.parse(readFileSync(new URL('../../bun-runtime.json',import.meta.url),'utf8'))
export const bunVersion=bunDefinition.version
const installed=join(productRoot,'.build/runtimes/bun/darwin-arm64')
export const sourceBun=join(installed,'payload/bun')
export const bunManifest=executable=>({kind:'bun',version:bunVersion,executable})

export async function fileHash(path,algorithm='sha256',encoding='hex'){
 const hash=createHash(algorithm)
 for await(const chunk of createReadStream(path))hash.update(chunk)
 return hash.digest(encoding)
}
const exists=path=>lstat(path).then(()=>true,error=>{if(error.code==='ENOENT')return false;throw error})

// A dropped connection is tried again, up to four times in all; a wrong digest is not.
async function archive(){
 for(let attempt=1;;attempt++){
  try{return await download()}
  catch(error){
   if(attempt>=4||/digest|budget/.test(error.message))throw error
   console.error(`Bun download attempt ${attempt} failed (${error.cause?.code??error.message}); retrying`)
   await new Promise(resolve=>setTimeout(resolve,2000*attempt))
  }
 }
}
async function download(){
 const expected=bunDefinition.integrity.replace(/^sha512-/,''),directory=join(productRoot,'.build/cache/archives')
 const path=join(directory,`bun-${bunVersion}.tgz`)
 await mkdir(directory,{recursive:true})
 if(await exists(path)&&await fileHash(path,'sha512','base64')===expected)return path
 const temporary=path+'.partial-'+randomUUID()
 try{
  const response=await fetch(bunDefinition.url,{signal:AbortSignal.timeout(600000)})
  if(!response.ok||!response.body)throw Error('Bun download failed: '+response.status)
  let size=0;const hash=createHash('sha512')
  const counter=new Transform({transform(chunk,_encoding,callback){size+=chunk.length;if(size>512*1024*1024)callback(Error('Bun download exceeds its byte budget'));else{hash.update(chunk);callback(null,chunk)}}})
  await pipeline(Readable.fromWeb(response.body),counter,createWriteStream(temporary,{flags:'wx',mode:0o600}))
  if(hash.digest('base64')!==expected)throw Error('Bun archive digest mismatch')
  await rename(temporary,path);return path
 }finally{await rm(temporary,{force:true})}
}

/** Installs the pinned Bun unless the installed one matches its receipt. */
export async function prepareBun(){
 if(process.platform!=='darwin'||process.arch!=='arm64')throw Error('Bun preparation targets macOS arm64')
 const receipt=join(installed,'manifest.json')
 try{
  const saved=JSON.parse(await readFile(receipt,'utf8'))
  if(saved.definition===JSON.stringify(bunDefinition)&&await fileHash(sourceBun)===saved.files[0].sha256)return sourceBun
 }catch{}
 const downloaded=await archive(),staging=installed+'.staging-'+randomUUID()
 try{
  const entries=execFileSync('tar',['-tzf',downloaded],{encoding:'utf8'}).trim().split('\n')
  if(entries.some(path=>!path||path.startsWith('/')||path.split('/').includes('..')))throw Error('Unsafe Bun archive entry')
  await mkdir(join(staging,'extract'),{recursive:true});await mkdir(join(staging,'payload'))
  execFileSync('tar',['-xzf',downloaded,'-C',join(staging,'extract')],{stdio:'pipe'})
  const executable=join(staging,'payload/bun')
  await copyFile(join(staging,'extract',bunDefinition.executable),executable);await chmod(executable,0o755)
  if(execFileSync(executable,['--version'],{encoding:'utf8',timeout:10000}).trim()!==bunVersion)throw Error('Bun version mismatch')
  const size=(await lstat(executable)).size
  await writeFile(join(staging,'manifest.json'),JSON.stringify({definition:JSON.stringify(bunDefinition),version:bunVersion,files:[{path:'bun',sha256:await fileHash(executable),size}]},null,2)+'\n')
  await rm(join(staging,'extract'),{recursive:true,force:true})
  await rm(installed,{recursive:true,force:true});await mkdir(join(installed,'..'),{recursive:true});await rename(staging,installed)
  console.log('Prepared verified bun '+bunVersion)
  return sourceBun
 }finally{await rm(staging,{recursive:true,force:true})}
}

export async function verifyBun(executable,receipt){
 const provenance=JSON.parse(await readFile(receipt,'utf8'))
 assert.deepEqual(JSON.parse(provenance.definition),bunDefinition,'Bun provenance matches the reviewed download pin')
 assert.equal(provenance.version,bunVersion)
 assert.equal(provenance.files.length,1)
 assert.equal(provenance.files[0].path,'bun')
 assert.equal(await fileHash(executable),provenance.files[0].sha256,'Bun executable matches the verified archive')
 assert.equal(execFileSync(executable,['--version'],{encoding:'utf8',env:{PATH:'/usr/bin:/bin'}}).trim(),bunVersion)
}

if(process.argv[1]&&realpathSync(process.argv[1])===fileURLToPath(import.meta.url))await prepareBun()
