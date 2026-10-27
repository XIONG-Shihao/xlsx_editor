import {test} from 'node:test'
import assert from 'node:assert/strict'
import {mkdtemp,writeFile,open,rm} from 'node:fs/promises'
import {join} from 'node:path'
import {fileURLToPath} from 'node:url'
import {tmpdir} from 'node:os'
import {spawnSync} from 'node:child_process'
import {build} from 'esbuild'
// The desktop coordinator writes the profile only through the lock its shell passes down.
const product=fileURLToPath(new URL('..',import.meta.url)),lease=join(product,'.build/tests/profile-lease.mjs')
await build({entryPoints:[join(product,'desktop/backend/host/profileLease.ts')],bundle:true,platform:'node',format:'esm',outfile:lease})
test('only the exact inherited profile lock inode is accepted for a writer',async()=>{
  const root=await mkdtemp(join(tmpdir(),'tauri-profile-fd-'))
  const other=join(root,'other-file');await writeFile(other,'');await writeFile(join(root,'.tauri-profile.lock'),'')
  const valid=await open(join(root,'.tauri-profile.lock'),'r'),invalid=await open(other,'r')
  try{
    const run=fd=>spawnSync(process.execPath,['--input-type=module','-e',`import {profileLockFd} from ${JSON.stringify(lease)};console.log(profileLockFd())`],{env:{...process.env,XLSX_EDITOR_USER_DATA:root,XLSX_EDITOR_PROFILE_LOCK_FD:'3'},stdio:['ignore','pipe','pipe',fd],encoding:'utf8'})
    const held=run(valid.fd);assert.equal(held.status,0,held.stderr);assert.equal(held.stdout.trim(),'3')
    const wrong=run(invalid.fd);assert.notEqual(wrong.status,0);assert.match(wrong.stderr,/does not match this profile/)
  }finally{await valid.close();await invalid.close();await rm(root,{recursive:true,force:true})}
})
