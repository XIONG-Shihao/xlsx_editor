import type {SpreadsheetStatus,SpreadsheetRequest,SpreadsheetProgress} from './document-session.mjs';
export interface WorkerProgressOptions {onProgress?:(event:SpreadsheetProgress)=>void;onStart?:(requestId:number)=>void}
export class SpreadsheetWorker {
  constructor(options:{filename?:string;newWorkbook?:{name?:string};recoveryId?:string;resume?:boolean;fresh?:boolean;prepared?:boolean;recoveryDirectory:string;maxRecoveryBytes?:number;maxIndexCacheBytes?:number;initialCalculationPolicy?:{mode:'auto'|'manual';recalculateBeforeSave:boolean}}&WorkerProgressOptions);
  open():Promise<SpreadsheetStatus>;
  request(request:SpreadsheetRequest,options?:WorkerProgressOptions):Promise<any>;
  authorizeCancel(requestId:number):Promise<{granted:boolean;reason?:'completed'|'finishing'}>;
  readonly closed:boolean;
  /** Cancels this owner; accepted recovery entries survive. Reopen before continuing. */
  cancel():Promise<void>;
  close():Promise<void>;
}
