// The product carries no Electron: no imports of it or its packaging libraries,
// and none in the manifest, the lock or the scripts.
import {test} from 'node:test'
import assert from 'node:assert/strict'
import {existsSync,readdirSync,readFileSync,statSync} from 'node:fs'
import {join} from 'node:path'
import {fileURLToPath} from 'node:url'
import {readBunLock} from '../desktop/scripts/tooling/bun-lock.mjs'
import {prepareBun} from '../desktop/scripts/tooling/bun.mjs'

const product=fileURLToPath(new URL('..',import.meta.url))
const skip=new Set(['node_modules','dist','target','out','release','.build','corpus'])
function* sources(directory){
 for(const name of readdirSync(directory)){
  if(skip.has(name))continue
  const path=join(directory,name)
  if(statSync(path).isDirectory())yield* sources(path)
  else if(/\.(ts|tsx|mjs|cjs|js)$/.test(name))yield path
 }
}
test('nothing imports Electron or its packaging libraries',()=>{
 for(const file of sources(product)){
  const text=readFileSync(file,'utf8')
  assert.doesNotMatch(text,/(?:from\s*|import\s*\(|require\s*\()\s*['"](?:electron(?:\/[^'"]*)?|@electron\/[^'"]+|electron-builder|app-builder-lib[^'"]*|builder-util)['"]/,file)
  assert.doesNotMatch(text,/import\s*\{[^}]*\b_electron\b/,file)
 }
})
test('the manifest, lock and scripts hold no Electron',async()=>{
 const manifest=JSON.parse(readFileSync(join(product,'package.json'),'utf8')),lock=readBunLock(product,await prepareBun())
 for(const dependencies of [manifest.dependencies,manifest.devDependencies])
  for(const name of Object.keys(dependencies??{}))assert.doesNotMatch(name,/electron|^@electron\//i,name)
 for(const name of Object.keys(lock.packages))
  assert.doesNotMatch(name,/(?:^|\/)(?:electron|electron-builder|app-builder-lib|builder-util)(?:\/|$)|@electron\//,name)
 for(const command of Object.values(manifest.scripts))assert.doesNotMatch(command,/electron|stage-standalone|vite\.standalone/,command)
 assert.equal(manifest.main,undefined,'the editor resolves through its library exports')
 assert.equal(existsSync(join(product,'electron-builder.cjs')),false)
})
