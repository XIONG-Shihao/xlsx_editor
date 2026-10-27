import type {WorkbookView,WorkbookCommand,WorkbookEdit,WorkbookCell} from '../core/contracts.js';
import type {SelectionSummary} from '../core/indexed-session.mjs';
export class CsvFileWorker {
 readonly id:string;
 open(source:string,options?:{expectedRevision?:string;delimiter?:string;maxSessionBytes?:number;maxIndexBytes?:number;maxHistoryBytes?:number;maxStates?:number;maxProjectionBytes?:number;maxReplyBytes?:number}):Promise<WorkbookView>;
 read(state:string,sheet:string,range:string):Promise<{sheet:string;range:string;cells:Record<string,WorkbookCell>}>;
 scan(state:string,sheet:string,after?:string):Promise<{cells:Record<string,WorkbookCell>;nextCursor:string|null}>;
 query(state:string,sheet:string,query:string,mode:string,after?:string,options?:{range?:string;matchLimit?:number}):Promise<{cells:Record<string,WorkbookCell>;matches:number;scanned:number;nextCursor:string|null}>;
 selectionSummary(state:string,sheet:string,range:string):Promise<SelectionSummary>;
 apply(state:string,commands:WorkbookCommand[]):Promise<{view:WorkbookView;edits:WorkbookEdit[]}>;
 compact(retain:string[]):Promise<{compacted:boolean}>;
 export(state:string):Promise<{path:string;revision:string}>;
 diagnosticStats():Record<string,number>;
 close():Promise<void>;
}
