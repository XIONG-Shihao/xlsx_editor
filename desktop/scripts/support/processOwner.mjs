import {execFileSync} from 'node:child_process'
import assert from 'node:assert/strict'
/** PID plus creation/command identity; retain descendants after reparenting.
 * OS RSS may include shared pages. Never count unrelated processes. */
export function processOwner(pid,{role=(_row,isRoot)=>isRoot?'root':'helper'}={}){
 assert.ok(Number.isInteger(pid)&&pid>0,'A live owner PID is required')
 const identities=new Map()
 function live(){
  const rows=execFileSync('ps',['-axo','pid,ppid,lstart,command'],{encoding:'utf8'}).split('\n').flatMap(line=>{const match=line.match(/^\s*(\d+)\s+(\d+)\s+(.+)$/);return match?[{pid:Number(match[1]),parentPid:Number(match[2]),identity:match[3]}]:[]})
  const root=rows.find(row=>row.pid===pid);if(root&&!identities.has(pid))identities.set(pid,root.identity)
  const seen=new Set(rows.filter(row=>identities.get(row.pid)===row.identity).map(row=>row.pid));let changed=true
  while(changed){changed=false;for(const row of rows)if(seen.has(row.parentPid)&&!seen.has(row.pid)){seen.add(row.pid);identities.set(row.pid,row.identity);changed=true}}
  return rows.filter(row=>identities.get(row.pid)===row.identity)
 }
 function sample(label){const owned=live();assert.ok(owned.some(row=>row.pid===pid),'Memory owner has exited');const rss=new Map(execFileSync('ps',['-p',owned.map(row=>row.pid).join(','),'-o','pid=,rss='],{encoding:'utf8'}).trim().split('\n').map(line=>line.trim().split(/\s+/).map(Number)));assert.ok(rss.has(pid),'Owner RSS is unavailable');const processes=owned.filter(row=>rss.has(row.pid)).map(row=>({...row,role:role(row,row.pid===pid),rssBytes:rss.get(row.pid)*1024}));return {label,sampledAtMs:Date.now(),processes,skippedProcesses:owned.length-processes.length,rssTotalBytes:processes.reduce((sum,row)=>sum+row.rssBytes,0)}}
 async function closed(){const started=performance.now();while(live().length&&performance.now()-started<10000)await new Promise(resolve=>setTimeout(resolve,100));assert.deepEqual(live(),[],'Owned processes remain; preserve the fixture for diagnosis')}
 assert.ok(live().some(row=>row.pid===pid),'The owner is not live');return {sample,live,closed}
}
