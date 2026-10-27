// Native code in a packaged app: found, signed if unsigned, and listed with
// its hashes, for packaging.
import {readdir,realpath,open} from 'node:fs/promises'
import {createReadStream} from 'node:fs'
import {createHash} from 'node:crypto'
import {join,relative} from 'node:path'
import {execFileSync} from 'node:child_process'
export async function files(root,bundle){const result=[];for(const entry of await readdir(root,{withFileTypes:true})){const path=join(root,entry.name);if(entry.isDirectory())result.push(...await files(path,bundle));else if(entry.isFile())result.push(path);else if(entry.isSymbolicLink()){const target=await realpath(path);if(relative(bundle,target).startsWith('..'))throw Error('Package link escapes: '+relative(bundle,path))}}return result}
export async function hash(path){const hash=createHash('sha256');for await(const chunk of createReadStream(path))hash.update(chunk);return hash.digest('hex')}
export async function natives(bundle){const root=join(bundle,'Contents/Resources'),result=[];for(const path of await files(root,bundle)){const fd=await open(path,'r'),bytes=Buffer.alloc(4);try{await fd.read(bytes,0,4,0)}finally{await fd.close()}if(['cffaedfe','cefaedfe','feedfacf','feedface','cafebabe','bebafeca'].includes(bytes.toString('hex')))result.push(path)}return result}
export async function signNativeResources(bundle){const root=join(bundle,'Contents/Resources'),paths=await natives(bundle);for(const path of paths){try{execFileSync('codesign',['--verify','--strict',path],{stdio:'pipe'})}catch{execFileSync('codesign',['--force','--sign','-',path],{stdio:'pipe'})}}return Promise.all(paths.map(async path=>({path:relative(root,path),sha256:await hash(path)})))}
