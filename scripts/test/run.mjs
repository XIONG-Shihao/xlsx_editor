// Runs every test file of this product, several at a time, each in its own process group.
//
//   node scripts/test/run.mjs                 build, then run the suite
//   options: --jobs N (default: three quarters of the cores), --only text,
//            --max-minutes N (default 30), --list (print the tasks without running them)
//
// The package scripts group test files into commands chained with &&. Here every
// file named in them runs once, as its own task, with the flags its command used;
// every other file in tests/ runs too, unless tests/not-in-suite.json names it.
// Each file has a deadline of four times its last time (2 to 10 minutes); the
// slowest start first (times kept in .build/test-timings.json); a file that fails
// is run once more alone, unless many failed together.
//
// Python checks need Python 3.10+ with openpyxl (scripts/audit/python.mjs):
// XLSX_EDITOR_TEST_PYTHON, or uv on PATH (or XLSX_EDITOR_UV).
import {spawn,spawnSync} from 'node:child_process'
import {existsSync,readdirSync,readFileSync,writeFileSync} from 'node:fs'
import {availableParallelism} from 'node:os'
import {join} from 'node:path'
import {fileURLToPath} from 'node:url'
import {pythonOnPath} from '../audit/python.mjs'

const root=fileURLToPath(new URL('../..',import.meta.url))
const args=process.argv.slice(2),option=name=>{const i=args.indexOf('--'+name);return i<0?undefined:args[i+1]}
const jobs=Math.max(1,Number(option('jobs')??Math.floor(availableParallelism()*0.75))),only=option('only')
const present=new Set(readdirSync(join(root,'tests')).filter(file=>file.endsWith('.mjs')))

function suite(){
 const scripts=JSON.parse(readFileSync(join(root,'package.json'),'utf8')).scripts,out=[],taken=new Set()
 const expand=pattern=>pattern.includes('*')?[...present].filter(file=>new RegExp('^'+pattern.replace(/[.]/g,'\\.').replace(/\*/g,'.*')+'$').test(file)).sort():[pattern]
 for(const group of ['test:format','test:ordinary','test','test:document','test:features','test:compat','test:pivots','test:remaining-formats','test:tables','test:scalar-arrays','test:month-dates','test:search','test:tabs','test:delimited-import','test:objects:extensions'])
  for(const command of (scripts[group]??'').split(' && ').filter(Boolean)){
   if(/^npm run /.test(command))continue
   const words=command.split(/\s+/),named=words.filter(word=>/^tests\//.test(word))
   if(words[0]!=='node'||!named.length){if(!taken.has(command)){taken.add(command);out.push({id:command.slice(0,70),command:'sh',args:['-c',command]})}continue}
   const flags=words.slice(1).filter(word=>!/^tests\//.test(word)&&word)
   if(flags.includes('--test')&&!flags.includes('--test-force-exit'))flags.push('--test-force-exit')
   for(const file of named.flatMap(word=>expand(word.slice('tests/'.length)))){
    if(taken.has(file))continue
    taken.add(file)
    out.push({id:file,command:process.execPath,args:[...flags,'tests/'+file]})
   }
  }
 const excluded=JSON.parse(readFileSync(join(root,'tests/not-in-suite.json'),'utf8')),outside=new Set(Object.entries(excluded).filter(([key])=>key!=='_about').flatMap(([,value])=>Array.isArray(value)?value:Object.keys(value)))
 for(const file of outside)if(!present.has(file))throw Error('tests/not-in-suite.json names a file that does not exist: '+file)
 for(const file of [...present].sort())if(!taken.has(file)&&!outside.has(file))out.push({id:file,command:process.execPath,args:['--import','./scripts/ts-resolve-register.mjs','--test','--test-force-exit','--test-concurrency=1','tests/'+file]})
 return out
}
const tasks=suite().filter(task=>!only||task.id.includes(only))
if(args.includes('--list')){for(const task of tasks)console.log(task.id.padEnd(60)+' '+[task.command.split('/').pop(),...task.args].join(' '));process.exit(0)}

pythonOnPath()
// The tests import the built editor (dist/) and run the built engine: build both first.
if(spawnSync('npm',['run','build'],{cwd:root,stdio:['ignore','inherit','inherit']}).status!==0)throw Error('Building the editor failed')

const timingsFile=join(root,'.build/test-timings.json')
const lastSeconds=existsSync(timingsFile)?JSON.parse(readFileSync(timingsFile,'utf8')):{}
const limitFor=task=>{const last=lastSeconds[task.id];return last===undefined?600000:Math.min(600000,Math.max(120000,last*4000))}
const running=new Map()
function run(task,limitMs=limitFor(task)){
 return new Promise(done=>{
  const started=Date.now(),child=spawn(task.command,task.args,{cwd:root,detached:true,stdio:['ignore','pipe','pipe']})
  let output='';const keep=chunk=>{output=(output+chunk).slice(-400000)};child.stdout.on('data',keep);child.stderr.on('data',keep)
  const stop=()=>{try{process.kill(-child.pid,'SIGKILL')}catch{/* The group already ended. */}}
  let timedOut=false
  const timer=setTimeout(()=>{timedOut=true;output+='\nTimed out after '+limitMs+' ms';stop()},limitMs)
  running.set(task.id,{started,stop})
  child.on('close',code=>{clearTimeout(timer);stop();running.delete(task.id)
   const count=key=>[...output.matchAll(new RegExp('^(?:ℹ|#) '+key+' (\\d+)','gm'))].reduce((sum,match)=>sum+Number(match[1]),0)
   done({id:task.id,ok:code===0,timedOut,seconds:Math.round((Date.now()-started)/1000),pass:count('pass'),fail:count('fail'),skipped:count('skipped'),
    failures:[...new Set([...output.matchAll(/^(?:✖) (.*)$/gm)].map(match=>match[1].replace(/ \([\d.]+ms\)$/,'')).filter(line=>line!=='failing tests:'))].slice(0,20),tail:code===0?'':output.slice(-1500)})})
 })
}
// Files of 10 s or more last time take at most half the slots.
const heavy=task=>(lastSeconds[task.id]??600)>=10
async function pool(list,width){
 const results=[],waiting=[...list],heavyCap=Math.max(1,Math.floor(width/2));let heavyRunning=0
 const take=()=>{const index=heavyRunning<heavyCap?0:waiting.findIndex(task=>!heavy(task));return index<0?undefined:waiting.splice(index,1)[0]}
 const slot=async()=>{
  for(;;){
   const task=take()
   if(!task){if(!waiting.length)return;await new Promise(resolve=>setTimeout(resolve,200));continue}
   const counted=heavy(task);if(counted)heavyRunning++
   const result=await run(task);if(counted)heavyRunning--
   results.push(result);process.stdout.write((result.ok?'ok   ':'FAIL ')+String(result.seconds).padStart(4)+'s  '+result.id+'\n')
  }
 }
 await Promise.all(Array.from({length:Math.min(width,list.length)},slot))
 return results
}
const maxMinutes=Number(option('max-minutes')??30),runStarted=Date.now()
const watch=setInterval(()=>{
 const now=Date.now()
 for(const [id,{started}] of running){const usual=lastSeconds[id];if(usual!==undefined&&now-started>Math.max(60000,usual*2000))console.log(`still running: ${id} (${Math.round((now-started)/1000)} s, usually ${usual} s)`)}
 if(now-runStarted>maxMinutes*60000){
  console.log(`\nStopped: the run passed its ${maxMinutes}-minute limit. Still running: ${[...running.keys()].join(', ')||'nothing'}`)
  for(const {stop} of running.values())stop()
  process.exit(2)
 }
},60000)
watch.unref()
const longestFirst=list=>[...list].sort((a,b)=>(lastSeconds[b.id]??600)-(lastSeconds[a.id]??600))
console.log(`${tasks.length} files, ${jobs} at a time`)
const first=await pool(longestFirst(tasks),jobs)
const failedFirst=first.filter(result=>!result.ok),shared=failedFirst.length>Math.max(5,first.length*0.03)
if(shared)console.log(`\n${failedFirst.length} files failed together: a shared cause, so none are retried.`)
const retry=shared?[]:failedFirst.filter(result=>!result.timedOut).map(result=>tasks.find(task=>task.id===result.id))
if(retry.length)console.log(`\n${retry.length} retried alone`)
const second=await pool(retry,1)
const final=new Map(first.map(result=>[result.id,result]));for(const result of second)final.set(result.id,{...result,retried:true,firstPassFailures:final.get(result.id)?.failures??[]})
const all=[...final.values()],failed=all.filter(result=>!result.ok),flaky=all.filter(result=>result.ok&&result.retried)
try{writeFileSync(timingsFile,JSON.stringify({...lastSeconds,...Object.fromEntries(first.map(result=>[result.id,result.seconds]))},null,1)+'\n')}catch{/* Ordering only. */}
console.log(`\n${all.length} files in ${Math.round((Date.now()-runStarted)/1000)} s: ${all.length-failed.length} passed, ${failed.length} failed`+(flaky.length?`, ${flaky.length} passed only when run alone`:''))
console.log(`tests: ${all.reduce((sum,result)=>sum+result.pass,0)} passed, ${all.reduce((sum,result)=>sum+result.fail,0)} failed, ${all.reduce((sum,result)=>sum+result.skipped,0)} skipped`)
for(const result of flaky){console.log('load-sensitive: '+result.id);for(const line of result.firstPassFailures)console.log('  failed in the parallel pass: '+line)}
for(const result of failed){console.log('\nFAILED '+result.id);for(const line of result.failures)console.log('  '+line);if(!result.failures.length)console.log(result.tail.split('\n').slice(-12).join('\n'))}
process.exitCode=failed.length?1:0
