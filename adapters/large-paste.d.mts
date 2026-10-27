import type {PastePlan,PasteDescriptor,PastePage} from '../core/paste-protocol.mjs';
export interface PasteUpload {id:string;state:string}
export interface PasteStager {
  begin(request:PastePlan&{state:string}):Promise<{id:string;sequence:number}>;
  upload(request:PasteUpload&PastePage):Promise<{id:string;sequence:number;count:number;bytes:number}>;
  seal(request:PasteUpload):Promise<PasteDescriptor>;
  cancel(request:PasteUpload):Promise<{cancelled:true}>;
  pages(descriptor:PasteDescriptor,context:{state:string}):AsyncIterable<PastePage>;
  prune(referencedIds:string[]):Promise<{removed:number}>;
  close():Promise<void>;
  stats():{activeUploads:number;attachmentBytes:number};
}
export function createPasteStager(options:{directory:string;documentId:string;maxBytes?:number;checkCancelled?:()=>void;accountBytes?:(change:{id:string;delta:number;kind:'reserve'|'release'})=>Promise<void>}):PasteStager;
