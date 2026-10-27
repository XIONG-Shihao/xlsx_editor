// Copies this product by itself into an empty folder and installs, builds, tests
// and packages it there, to show it needs nothing beside it.
//
//   node scripts/test/folder-alone.mjs [--keep]
//
// The copy holds the files Git would keep (tracked and new, not ignored), or,
// outside a Git checkout, everything but build output. It starts with no
// node_modules, no build output and no Bun: all come from this folder's own pins.
// It is made a Git repository of its own, as the product would be on its own
// (some checks list the product's files with git ls-files).
import {execFileSync,spawnSync} from 'node:child_process'
import {cpSync,mkdirSync,mkdtempSync,readdirSync,rmSync,statSync} from 'node:fs'
import {tmpdir} from 'node:os'
import {dirname,join,relative} from 'node:path'
import {fileURLToPath} from 'node:url'

const product=fileURLToPath(new URL('../..',import.meta.url)).replace(/\/$/,'')
const output=new Set(['node_modules','dist','target','out','release','.build','gen','artifacts','__pycache__','bin','.DS_Store'])
function listed(){
 const git=spawnSync('git',['ls-files','-co','--exclude-standard','-z'],{cwd:product,encoding:'utf8'})
 if(git.status===0)return git.stdout.split('\0').filter(Boolean)
 const files=[]
 const walk=directory=>{for(const name of readdirSync(directory)){if(output.has(name))continue;const path=join(directory,name);statSync(path).isDirectory()?walk(path):files.push(relative(product,path))}}
 walk(product);return files
}
const copy=join(mkdtempSync(join(tmpdir(),'xlsx-alone-')),'xlsx')
for(const file of listed()){
 try{statSync(join(product,file))}catch{continue} // Deleted but not yet committed.
 mkdirSync(dirname(join(copy,file)),{recursive:true});cpSync(join(product,file),join(copy,file))
}
execFileSync('git',['init','-q'],{cwd:copy});execFileSync('git',['add','-A'],{cwd:copy})
console.log('Copied the product alone to '+copy)
const steps=[
 ['Editor and desktop build','npm',['run','desktop:build']],
 ['TypeScript','npm',['run','typecheck']],
 ['Nothing reaches outside','node',['--test','tests/product-folder.mjs']],
 ['Desktop shell tests','cargo',['test','--offline','--manifest-path','desktop/src-tauri/Cargo.toml']],
 ['Every test file','node',['scripts/test/run.mjs']],
 ['Desktop coordinator','node',['desktop/scripts/verify-coordinator.mjs']],
 ['Package','npm',['run','desktop:package']],
]
let failed=false
for(const [name,command,args] of steps){
 console.log(`\n== ${name}: ${command} ${args.join(' ')}`)
 const run=spawnSync(command,args,{cwd:copy,stdio:'inherit'})
 console.log(`== ${name}: ${run.status===0?'passed':'FAILED'}`)
 if(run.status!==0){failed=true;break}
}
if(failed||process.argv.includes('--keep'))console.log('\nThe copy is kept at '+copy)
else rmSync(dirname(copy),{recursive:true,force:true})
process.exitCode=failed?1:0
