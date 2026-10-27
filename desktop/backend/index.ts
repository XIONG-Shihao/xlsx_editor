import {setupHost,send} from './host/host'
import {profileLockFd} from './host/profileLease'
import {setStandalonePlatform} from '../../standalone/platform.mjs'
import {platform,handlers,ready,connectEvents,invoke,dispatchMenu,incomingFiles,requestClose,disconnected} from './platform'

profileLockFd() // Verify the inherited writer lease before starting workers.
connectEvents((channel,...args)=>{void send({kind:'event',channel,args}).catch(console.error)})
setupHost(async(channel,args)=>{
  if(channel==='standalone:initialize'){await ready;return {product:'standalone-xlsx-tauri',javascriptRuntime:{kind:process.versions.bun?'bun':'node',version:process.versions.bun??process.versions.node},methods:[...handlers.keys()].sort(),profileLockHeld:profileLockFd()!==undefined}}
  if(channel==='standalone:close'){await ready;requestClose();return {pending:true}}
  if(channel==='standalone:menu'){await ready;dispatchMenu(args[0]);return null}
  if(channel==='standalone:files'){await ready;incomingFiles(args[0]);return null}
  return invoke(channel,args)
},error=>{console.error(error);disconnected()})
// The editor's standalone code runs on this shell's platform.
setStandalonePlatform(platform)
void import('../../standalone/main.mjs').catch(error=>{console.error(error);process.exitCode=1;disconnected()})
process.stdin.on('end',disconnected)
process.on('SIGTERM',disconnected)
