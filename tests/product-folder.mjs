// The product builds, tests and packages from this folder alone. No file in it
// reaches outside it by a relative path (imports, file reads, fixture paths,
// new URL(...), Cargo paths, document links), and none names a place in the
// repository it came from.
import {test} from 'node:test'
import assert from 'node:assert/strict'
import {readdirSync,readFileSync,statSync} from 'node:fs'
import {dirname,join,relative,resolve,sep} from 'node:path'
import {fileURLToPath} from 'node:url'

const product=fileURLToPath(new URL('..',import.meta.url)).replace(/\/$/,'')
const skip=new Set(['node_modules','dist','target','out','release','.build','gen','__pycache__'])
const text=/\.(?:mjs|cjs|js|mts|ts|tsx|rs|toml|json|py|sh|md|html|css)$/
function* files(directory){
 for(const name of readdirSync(directory)){
  if(skip.has(name))continue
  const path=join(directory,name),info=statSync(path)
  if(info.isDirectory())yield* files(path)
  else if(text.test(name)&&info.size<2_000_000)yield path
 }
}
// Quoted relative paths and Markdown links: what code and documents reach by path.
const reached=source=>[...source.matchAll(/['"`](\.\.\/[^'"`\s]*)['"`]|\]\((\.\.\/[^)\s#]*)/g)].map(match=>match[1]??match[2])
// Escapes some tests feed the editor on purpose, to see it refuse them.
const attacks=new Set(['tests/document-objects.mjs: ../../private','tests/package-capabilities.mjs: ../../renamed.dat','tests/portable-package.test.mjs: ../../outside.xml'])
// The repository's other places, written as paths.
const elsewhere=/\b(?:editor\/(?:xlsx|docx|calc)|standalone-(?:docx|calc|host)|standalone-xlsx\/|OfficeMCP\/)/

test('nothing in the product reaches outside its folder',()=>{
 const problems=[]
 for(const file of files(product)){
  const source=readFileSync(file,'utf8'),name=relative(product,file)
  for(const path of reached(source)){
   const target=resolve(dirname(file),path.replace(/\$\{[^}]*\}.*$/,''))
   if(target!==product&&!target.startsWith(product+sep)&&!attacks.has(`${name}: ${path}`))problems.push(`${name}: ${path}`)
  }
  const named=source.match(elsewhere)
  if(named&&file!==fileURLToPath(import.meta.url))problems.push(`${name}: names ${named[0]}`)
 }
 assert.deepEqual(problems,[])
})
