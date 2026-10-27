import './grid-host'
/** Test owner only: real renderer grid, history and WorkbookWorker. No evaluator. */
import React, { useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { WorkbookWorker } from '@office-editors/xlsx-editor/ui/workbookWorker'
import { SpreadsheetGrid } from '@office-editors/xlsx-editor/ui/SpreadsheetGrid'
import { SheetHistory } from '@office-editors/xlsx-editor/ui/sheetHistory'
import type { WorkbookCommand, WorkbookView } from '@office-editors/xlsx-editor/contracts'

function Harness() {
  const [view,setView]=useState<WorkbookView|null>(null),[busy,setBusy]=useState(false),[error,setError]=useState('')
  const [focus,setFocus]=useState<{sheet:string;cell:string}>()
  const worker=useRef<WorkbookWorker>(),current=useRef<WorkbookView>(),history=useRef<SheetHistory>(),lastCommands=useRef<WorkbookCommand[]>([])
  const snapshot=()=>JSON.parse(JSON.stringify({view:current.current,edits:history.current?.current.edits,frame:history.current?.current.id,lastCommands:lastCommands.current}))
  const apply=async(commands:WorkbookCommand[])=>{
    lastCommands.current=commands
    setBusy(true)
    try { const result=await worker.current!.apply(current.current!,commands);history.current!.batch(result.edits,result.view.sheets,result.view);current.current=result.view;setView(result.view);setError('');return {ok:true,...snapshot()} }
    catch(error){setError(String(error));return {ok:false,error:String(error),...snapshot()}}
    finally{setBusy(false)}
  }
  const restore = (direction: 'undo' | 'redo') => {
    const frame = direction === 'undo' ? history.current!.undo() : history.current!.redo()
    if (!frame.view) throw Error('History frame is missing its workbook view')
    current.current = frame.view
    setView(frame.view)
  }
  useEffect(()=>{
    worker.current=new WorkbookWorker((view,commands)=>(globalThis as any).__indexedCall({operation:'apply',state:view.storage!.state,commands}))
    ;(globalThis as any).__xlsxDifferential={async initialize(base:WorkbookView){const prepared=await worker.current!.prepare(base);current.current=prepared;history.current=new SheetHistory(prepared.sheets);history.current.current.view=prepared;setView(prepared);return snapshot()},apply,snapshot,restoreFocus:setFocus}
    return()=>worker.current?.dispose()
  },[])
  return <><p role="alert">{error}</p>{view&&<SpreadsheetGrid sheets={view.sheets} namedRanges={view.namedRanges} tables={view.tables} nameCommands={view.storage?.kind==='rust-indexed-v1'} focus={focus} readOnly={false} busy={busy}
    storageKey={view.storage?`${view.storage.sessionId}:${view.storage.state}`:undefined}
    readRange={view.storage?(sheet,range)=>(globalThis as any).__indexedCall({operation:'read',state:view.storage!.state,sheet,range}).then((result:any)=>result.cells):undefined}
    scan={view.storage?(sheet,after)=>(globalThis as any).__indexedCall({operation:'scan',state:view.storage!.state,sheet,after}):undefined}
    query={view.storage?(sheet,query,mode,after,options)=>(globalThis as any).__indexedCall({operation:'query',state:view.storage!.state,sheet,query,mode,after,range:options?.range,matchLimit:options?.matchLimit}):undefined}
    filterValues={view.storage?.kind==='rust-indexed-v1'?(sheet,range,column,query,offset)=>(globalThis as any).__indexedCall({operation:'filterValues',state:view.storage!.state,sheet,range,column,query,offset}):undefined}
    canUndo={history.current?.canUndo} canRedo={history.current?.canRedo} onUndo={() => restore('undo')} onRedo={() => restore('redo')}
    onEdit={edit=>{void apply([edit])}} onBatch={commands=>{void apply(commands)}}/>}</>
}
createRoot(document.getElementById('root')!).render(<Harness/>);
