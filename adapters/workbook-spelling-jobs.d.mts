import type {OfflineSpellingAvailability,OfflineSpellingReply,OfflineSpellingProvider} from './offline-spelling.mjs';
export interface SpellingOwnerReceipt {readonly documentId:string;readonly ownerId:string;readonly state:string;readonly revision:string}
interface SpellingRequestBase extends SpellingOwnerReceipt {readonly requestId:string;readonly kind?:'indexed';readonly sessionId?:string}
export type WorkbookSpellingRequest=(SpellingRequestBase&{readonly operation:'spellingAvailability'|'spellingCancel'})|(SpellingRequestBase&{readonly operation:'spellingCheck';readonly locale:string;readonly words:readonly string[]});
export interface WorkbookSpellingAvailability {schema:1;requestId:string;receipt:SpellingOwnerReceipt;availability:OfflineSpellingAvailability}
export interface WorkbookSpellingCheck {schema:1;receipt:SpellingOwnerReceipt;result:OfflineSpellingReply}
export interface WorkbookSpellingCancellation {schema:1;requestId:string;cancelled:boolean}
export interface WorkbookSpellingResult {spellingAvailability?:WorkbookSpellingAvailability;spellingCheck?:WorkbookSpellingCheck;spellingCancellation?:WorkbookSpellingCancellation}
export interface WorkbookSpellingAccess {
  capture(request:WorkbookSpellingRequest):SpellingOwnerReceipt;
  verify(request:WorkbookSpellingRequest,receipt:SpellingOwnerReceipt):Promise<void>|void;
  authorize(request:WorkbookSpellingRequest,receipt:SpellingOwnerReceipt):Promise<void>|void;
}
export function spellingRequest(input:unknown):WorkbookSpellingRequest;
export class WorkbookSpellingJobs {
  constructor(options?:Partial<WorkbookSpellingAccess>&{providerFactory?:()=>OfflineSpellingProvider});
  run(request:unknown,access?:WorkbookSpellingAccess):Promise<WorkbookSpellingResult>;
  cancel(request:unknown,access?:WorkbookSpellingAccess):Promise<WorkbookSpellingResult>;
  retire():void;
  close():Promise<void>;
}
