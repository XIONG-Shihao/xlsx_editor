import type {WorkbookView, WorkbookCommand,WorkbookPrintSettings} from '../core/contracts.js';
import type {PasteShape,PasteDescriptor,TypedPasteConversion,PastePlan} from '../core/paste-protocol.mjs';
export interface SpreadsheetStatus {
  state:string; name:string; kind:'xlsx'|'xlsm'|'xls'|'xlsb'|'csv'; view:WorkbookView;
  dirty:boolean; canUndo:boolean; canRedo:boolean; recovered:boolean;
  documentId:string;viewPreferenceId:string;unsaved:boolean;
  warning?:string;
}
export interface SpreadsheetRequest {
  operation:'status'|'binding'|'apply'|'undo'|'redo'|'save'|'saveAs'|'activate'|'copy'|'discard'|'read'|'scan'|'trace'|'usedRange'|'noteMembership'|'notes'|'accessibility'|'duplicatePreview'|'previewFormula'|'cancelFormulaPreview'|'query'|'selection'|'filterValues'|'hyperlink'|'objects'|'object'|'vbaModules'|'vbaModule'|'vbaFunctions'|'evaluateVba'|'dataOnlyExport'|'conditionalRules'|'validationRules'|'importBegin'|'importUpload'|'importSeal'|'importCancel'|'importPrune'|'importReceipt'|'printDefinition'|'printPrepare'|'printPage'|'printDocument'|'printRelease';
  state?:string; commands?:(WorkbookCommand|PasteDescriptor)[]; sheet?:string; range?:string; query?:string;
  mode?:'display'|'raw'|'formula'; after?:string; destination?:string; expectedRevision?:string|null;
  matchLimit?:number;
  hasHeaders?:boolean;columns?:string[];
  column?:string;offset?:number;limit?:number;
  formula?:string;cell?:string;direction?:'precedents'|'dependents';
  documentId?:string;uploadId?:string;shape?:PasteShape;count?:number;policy?:PastePlan['policy'];conversion?:TypedPasteConversion;sequence?:number;cells?:[number,string][];
  descriptor?:PasteDescriptor;parentState?:string;
  objectId?:string;
  moduleId?:string;
  functionName?:string;args?:number[];format?:'xls'|'xlsb'|'xlsx';
  settings?:WorkbookPrintSettings;jobId?:string;pageNumber?:number;requestId?:string;
}
export interface SpreadsheetProgress {phase:string;cancellable:boolean;requestId?:number;recoveryId?:string}
export interface SpreadsheetOperationOptions {onProgress?:(event:SpreadsheetProgress)=>void;beforeCritical?:(phase:string)=>void;leaveCritical?:()=>void;checkCancelled?:()=>void}
export class SpreadsheetDocument {
  constructor(options:{recoveryDirectory:string;maxRecoveryBytes?:number;maxIndexCacheBytes?:number;indexDirectory?:string;native?:(input:any)=>Promise<any>;cacheIdentity?:string|null;initialCalculationPolicy?:{mode:'auto'|'manual';recalculateBeforeSave:boolean}}&SpreadsheetOperationOptions);
  open(filename:string,options?:{fresh?:boolean;prepared?:boolean}):Promise<SpreadsheetStatus>;
  create(name?:string,options?:{prepared?:boolean}):Promise<SpreadsheetStatus>;
  recover(documentId:string):Promise<SpreadsheetStatus>;
  resume(filename:string):Promise<SpreadsheetStatus>;
  status():SpreadsheetStatus;
  request(request:SpreadsheetRequest,options?:SpreadsheetOperationOptions):Promise<any>;
  close():Promise<void>;
}
export interface SpreadsheetRecoveryEntry {id:string;name:string;unsaved:boolean;dirty:boolean;createdAt:number;displayPath?:string}
export function listSpreadsheetRecovery(recoveryDirectory:string):Promise<{entries:SpreadsheetRecoveryEntry[];warnings:string[]}>;
