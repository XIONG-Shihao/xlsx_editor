/** Native desktop ports for the spreadsheet coordinator. The Rust shell owns
 * the WebView, process confinement and native window operations. */
import {EventEmitter} from 'node:events'
import {dirname,join} from 'node:path'
import {homedir} from 'node:os'
import {hostAction,dialog,clipboard} from './host/host'
import {createWorkbookPdfRenderer,createWorkbookNativePrinter} from './print'
export {dialog,clipboard}
// The native shell validates the scheme again and opens the user's browser.
export const shell={openExternal:async(url:string):Promise<void>=>{await hostAction('openExternal',{url})}}

type Handler=(event:unknown,...args:unknown[])=>unknown
export const handlers=new Map<string,Handler>()
let resolveReady:()=>void,rejectReady:(error:unknown)=>void
export const ready=new Promise<void>((resolve,reject)=>{resolveReady=resolve;rejectReady=reject})
void ready.catch(()=>{})
let publish:(channel:string,...args:unknown[])=>void=()=>{}
export const connectEvents=(send:typeof publish):void=>{publish=send}
const paths=new Map([['userData',process.env.XLSX_EDITOR_USER_DATA??join(homedir(),'Library/Application Support/XLSX Editor Tauri')]])
let menuReady:Promise<unknown>=Promise.resolve()
const menuActions=new Map<string,()=>void>()
export const app=Object.assign(new EventEmitter(),{
  isPackaged:process.env.XLSX_EDITOR_PACKAGED==='1',
  setName:(_name:string)=>{},
  getPath:(name:string):string=>name==='appData'?dirname(paths.get('userData')!):paths.get(name)??(()=>{throw Error('validation: Unsupported standalone host path')})(),
  setPath:(name:string,value:string):void=>{if(name!=='userData')throw Error('validation: Unsupported standalone host path');paths.set(name,value)},
  // The native shell selected this isolated profile and admitted the process.
  commandLine:{hasSwitch:(name:string)=>name==='user-data-dir'},
  requestSingleInstanceLock:()=>true,
  whenReady:()=>Promise.resolve(),
  quit:()=>publish('standalone:exit-approved'),
  exit:(code:number)=>{rejectReady(Error('unavailable: Standalone editor failed to start'));publish('standalone:exit-failed',{code})}
})
class Contents extends EventEmitter {
  mainFrame={url:''}
  send(channel:string,...args:unknown[]):void {publish(channel,...args)}
  setWindowOpenHandler(_handler:unknown):void{/* Native navigation/window policy. */}
}
export class BrowserWindow extends EventEmitter {
  static instance:BrowserWindow|undefined
  readonly webContents=new Contents()
  private destroyed=false
  constructor(_options:unknown){super();if(BrowserWindow.instance)throw Error('conflict: Standalone editor owns one window');BrowserWindow.instance=this}
  static getAllWindows():BrowserWindow[]{return this.instance&&!this.instance.destroyed?[this.instance]:[]}
  isDestroyed():boolean{return this.destroyed}
  isMinimized():boolean{return false}
  async loadURL(page:string):Promise<void>{
    // This is a private coordinator identity, never an arbitrary native URL.
    if(page!=='workbook-standalone://editor')throw Error('permission_denied: Unexpected standalone page identity')
    this.webContents.mainFrame.url=page
    await menuReady;resolveReady()
  }
  show():void{void hostAction('reveal').catch(console.error)}
  focus():void{void hostAction('focus').catch(console.error)}
  restore():void{this.show()}
  reload():void{void hostAction('standaloneReload').catch(console.error)}
  close():void{
    if(this.destroyed)return
    let prevented=false
    this.emit('close',{preventDefault:()=>{prevented=true}})
    if(!prevented){this.destroyed=true;this.emit('closed');app.emit('window-all-closed')}
  }
}
export const ipcMain={handle:(name:string,handler:Handler):void=>{
  if(!name.startsWith('xlsx:')||handlers.has(name))throw Error('invalid_request: Invalid standalone registration')
  handlers.set(name,handler)
}}
export async function invoke(name:string,args:unknown[]):Promise<unknown>{
  await ready
  const window=BrowserWindow.instance,handler=handlers.get(name)
  if(!window||window.isDestroyed()||!handler)throw Error('invalid_request: Unknown standalone operation')
  const result=await handler({sender:window.webContents,senderFrame:window.webContents.mainFrame},...args)
  if((name==='xlsx:close-result'&&args[0]===false)||(name==='xlsx:close-all'&&result&&typeof result==='object'&&'cancelled' in result&&result.cancelled===true))publish('standalone:exit-cancelled')
  return result
}
interface MenuEntry {label?:string;role?:string;type?:string;accelerator?:string;click?:()=>void;submenu?:MenuEntry[]}
function menuEntry(entry:MenuEntry,identity:string):unknown{
  if(entry.click)menuActions.set(identity,entry.click)
  return {id:identity,...(entry.label?{label:entry.label}:{}),...(entry.role?{role:entry.role}:{}),...(entry.type?{type:entry.type}:{}),...(entry.accelerator?{accelerator:entry.accelerator}:{}),...(entry.submenu?{submenu:entry.submenu.map((child,index)=>menuEntry(child,identity+'.'+index))}:{})}
}
export const Menu={
  buildFromTemplate:(entries:MenuEntry[])=>entries.map((entry,index)=>menuEntry(entry,'xlsx-menu.'+index)),
  setApplicationMenu:(items:unknown)=>{menuReady=hostAction('standaloneMenu',{items});void menuReady.catch(rejectReady)}
}
const permissions={setPermissionCheckHandler:(_value:unknown)=>{},setPermissionRequestHandler:(_value:unknown)=>{},setDevicePermissionHandler:(_value:unknown)=>{}}
export const session={defaultSession:permissions}
export function dispatchMenu(id:unknown):void{
  if(typeof id!=='string'||!menuActions.has(id))throw Error('invalid_request: Unknown standalone menu')
  menuActions.get(id)!()
}
export function incomingFiles(files:unknown):void{
  if(!Array.isArray(files)||files.length>16||files.some(file=>typeof file!=='string'||file.length>4096||!/^\//.test(file)||!/\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(file)))throw Error('validation: Invalid native spreadsheet launch selection')
  for(const file of files)app.emit('open-file',{preventDefault:()=>{}},file)
}
export function requestClose():void{BrowserWindow.instance?.close()}
export function disconnected():void{app.emit('host-disconnected')}
/** Everything platform.mjs asks of a shell, ready for setStandalonePlatform. */
export const platform={app,BrowserWindow,clipboard,dialog,ipcMain,Menu,session,shell,createWorkbookPdfRenderer,createWorkbookNativePrinter}
