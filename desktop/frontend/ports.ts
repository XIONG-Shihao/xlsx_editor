import {invoke as nativeInvoke} from '@tauri-apps/api/core'
import {listen} from '@tauri-apps/api/event'
import {encode,decode} from './codec'
const listeners=new Map<string,Set<(...args:unknown[])=>void>>()
let token=''
export const invoke=async(channel:string,...args:unknown[]):Promise<any>=>decode(await nativeInvoke('standalone_invoke',{token,channel,args:encode(args)}))
export function subscribe(channel:string,callback:(value:unknown)=>void):()=>void{
  const listener=(_event:unknown,value:unknown):void=>callback(value)
  let callbacks=listeners.get(channel)
  if(!callbacks){callbacks=new Set();listeners.set(channel,callbacks)}
  callbacks.add(listener)
  return()=>{callbacks.delete(listener)}
}
export async function initialize():Promise<void>{
  token=crypto.randomUUID()
  await listen<{channel:string;args:unknown[];receipt?:number}>('backend:event',event=>{
    if(event.payload.channel==='standalone:unavailable'){
      const banner=document.getElementById('standalone-service-error')??document.body.appendChild(Object.assign(document.createElement('aside'),{id:'standalone-service-error',className:'desktop-error'}))
      banner.setAttribute('role','alert');banner.textContent='The spreadsheet service stopped. Restart the editor to reopen accepted recovery data.'
    }
    for(const callback of listeners.get(event.payload.channel)??[])try{callback({},...event.payload.args.map(decode))}catch(error){console.error(error)}
    if(event.payload.receipt!==undefined)void nativeInvoke('standalone_event_receipt',{receipt:event.payload.receipt}).catch(console.error)
  })
  await nativeInvoke('standalone_initialize',{token})
}
