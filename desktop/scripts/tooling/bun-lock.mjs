import {readFileSync,realpathSync} from 'node:fs'
import {execFileSync} from 'node:child_process'
import {join,resolve} from 'node:path'
import assert from 'node:assert/strict'
import {sourceBun} from "./bun.mjs"

export function readBunLock(root,bun=sourceBun){
 return JSON.parse(execFileSync(bun,['--eval','console.log(JSON.stringify(Bun.JSONC.parse(await Bun.file(process.argv[1]).text())))',join(root,'bun.lock')],{encoding:'utf8'}))
}
export function verifyBunDependencies(root,bun=sourceBun){
 const manifest=JSON.parse(readFileSync(join(root,'package.json'),'utf8')),lock=readBunLock(root,bun)
 assert.deepEqual(manifest.dependencies,lock.workspaces[''].dependencies,'Production dependencies match the Bun lock')
 for(const name of Object.keys(manifest.dependencies??{})){
  const entry=lock.packages[name]
  assert(entry?.[0]?.startsWith(name+'@'),'Dependency is locked: '+name)
  const reference=entry[0].slice(name.length+1),installed=JSON.parse(readFileSync(join(root,'node_modules',name,'package.json'),'utf8'))
  let version=reference
  if(reference.startsWith('file:')){
   const local=resolve(root,reference.slice(5))
   assert.equal(realpathSync(join(root,'node_modules',name)),realpathSync(local),'Local component stays at its contained source')
   version=JSON.parse(readFileSync(join(local,'package.json'),'utf8')).version
  }else if(reference.startsWith('https:')){
   version=new URL(reference).pathname.match(/-(\d+\.\d+\.\d+(?:-[\w.-]+)?)\.tgz$/)?.[1]
   assert(version,'URL package must declare its version in the pinned archive name: '+name)
  }
  assert.equal(installed.version,version,'Installed production dependency matches Bun lock: '+name)
 }
 return Object.keys(manifest.dependencies??{}).length
}
