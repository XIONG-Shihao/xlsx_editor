import type {WorkbookView,WorkbookCell,WorkbookObjectPage} from '../core/contracts.js';
import type {SelectionSummary} from '../core/indexed-session.mjs';
export type PassiveFormat='xlsm'|'xls'|'xlsb';
export interface PassiveVbaModule {id:string;name:string;kind:'standard'|'document-or-class';readOnly:boolean;privateModule:boolean;sourceBytes:number;length:number;sha256:string}
export interface PassiveVbaPage {modules:PassiveVbaModule[];total:number;nextOffset:number|null;projectName?:string;codepage?:number;protectedReason?:string;execution:'disabled'}
export interface PassiveVbaSource {id:string;name:string;code:string;offset:number;total:number;nextOffset:number|null;sha256:string;execution:'disabled'}
export interface PureVbaFunctionPage {moduleId:string;sourceSha256:string;functions:Array<{name:string;parameters:string[];returnType:'Double'|'numeric constant'}>;limits:Record<string,number>;execution:'explicit-pure-numeric-only';protectedReason?:string}
export interface PureVbaEvaluation {moduleId:string;sourceSha256:string;sourceRevision:string;functionName:string;value:number;steps:number;policy:string;execution:'explicit-pure-numeric-only'}
export interface DataOnlyExport {path:string;format:'xls'|'xlsb'|'xlsx';revision:string;sourceRevision:string;bytes:number;sheets:number;positions:number;values:number;formulaValues:number;omissions:string[];policy:string}
/** Node/trusted document-owner API. Export paths are never renderer DTOs. */
export class PassiveWorkbook {
  readonly id:string;
  open(source:string,options:{format:PassiveFormat;storageDirectory?:string}):Promise<WorkbookView>;
  view(state:string):WorkbookView;
  has(state:string):boolean;
  diagnosticStats():{retainedStates:number;storageRoots:number;closed:number;calculationCalls:number;pendingCalls:number};
  retainedStates():string[];
  storageRoots():string[];
  read(state:string,sheet:string,range:string):Promise<{sheet:string;range:string;cells:Record<string,WorkbookCell>}>;
  scan(state:string,sheet:string,after?:string):Promise<{cells:Record<string,WorkbookCell>;nextCursor:string|null}>;
  query(state:string,sheet:string,query:string,mode:'display'|'raw'|'formula',after?:string,options?:{range?:string;matchLimit?:number}):Promise<{cells:Record<string,WorkbookCell>;matches:number;scanned:number;nextCursor:string|null}>;
  selectionSummary(state:string,sheet:string,range:string):Promise<SelectionSummary&{policy:string}>;
  calculationSummary(state:string):Promise<{revision:string;engine:string;calculated:0;stored:number;unresolved:number;error:0;circular:0}>;
  objects(state:string,sheet:string,options?:{after?:string;limit?:number}):Promise<WorkbookObjectPage&{protectedReason:string}>;
  vbaModules(state:string,options?:{offset?:number;limit?:number}):Promise<PassiveVbaPage>;
  vbaModule(state:string,moduleId:string,options?:{offset?:number;limit?:number}):Promise<PassiveVbaSource>;
  vbaFunctions(state:string,moduleId:string):Promise<PureVbaFunctionPage>;
  evaluateVba(state:string,moduleId:string,functionName:string,args:number[]):Promise<PureVbaEvaluation>;
  dataOnlyExport(state:string,options:{format:'xls'|'xlsb'|'xlsx'}):Promise<DataOnlyExport>;
  object(state:string):Promise<never>;
  apply(state:string,commands?:unknown[]):Promise<never>;
  conditionalRules(state:string):Promise<never>;
  validationRules(state:string):Promise<never>;
  printDefinition(state:string):Promise<never>;
  printSnapshot(state:string):Promise<never>;
  filterValues(state:string):Promise<never>;
  keys(state:string):Promise<never>;
  compact(keep?:string[]):Promise<{retainedStates:1;storageRoots:1;unchanged:true}>;
  export(state:string):Promise<{path:string;revision:string}>;
  close():Promise<void>;
}
