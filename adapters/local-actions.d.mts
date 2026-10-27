import type {LocalActionDraft,LocalActionList} from '../core/local-actions.mjs';
/** Main-only. Filename is a trusted host private-profile path, never a descriptor field. */
export class LocalActions {
 private constructor(filename:string);
 static open(filename:string):Promise<LocalActions>;
 list():Promise<LocalActionList>;
 create(draft:LocalActionDraft,expectedRegistryRevision:number):Promise<LocalActionList>;
 update(id:string,draft:LocalActionDraft,expectedRegistryRevision:number):Promise<LocalActionList>;
 remove(id:string,expectedRegistryRevision:number):Promise<LocalActionList>;
}
