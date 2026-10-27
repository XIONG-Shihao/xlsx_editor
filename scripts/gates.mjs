// Every gate of this product in order, run from its folder. Each step runs even
// when an earlier one failed, and the summary lists every result.
//
//   npm run gates
import {spawnSync} from 'node:child_process'
import {fileURLToPath} from 'node:url'

const root=fileURLToPath(new URL('..',import.meta.url))
const steps=[
 ['Editor and desktop build','npm',['run','desktop:build']],
 ['TypeScript','npm',['run','typecheck']],
 ['Desktop TypeScript','npx',['tsc','--noEmit','-p','desktop/tsconfig.json']],
 ['Desktop shell tests','cargo',['test','--offline','--manifest-path','desktop/src-tauri/Cargo.toml']],
 ['Every test file','node',['scripts/test/run.mjs']],
 ['Desktop coordinator','node',['desktop/scripts/verify-coordinator.mjs']],
]
const results=[]
for(const [name,command,args] of steps){
 console.log(`\n== ${name}: ${command} ${args.join(' ')}`)
 const started=Date.now(),run=spawnSync(command,args,{cwd:root,stdio:'inherit'})
 results.push([name,run.status===0?'passed':'FAILED',Date.now()-started])
}
console.log('\nGates:')
for(const [name,result,ms] of results)console.log(`  ${result.padEnd(8)} ${name} (${Math.round(ms/1000)} s)`)
process.exit(results.some(([,result])=>result==='FAILED')?1:0)
