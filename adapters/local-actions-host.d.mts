import type {LocalActionDraft,LocalActionList} from '../core/local-actions.mjs';
export const LOCAL_ACTION_REQUEST_BYTES:16384;
interface LocalActionBinding {readonly documentId:string;readonly state:string}
export type LocalActionRequest=
 | (LocalActionBinding&{readonly operation:'list'})
 | (LocalActionBinding&{readonly operation:'create';readonly expectedRevision:number;readonly draft:LocalActionDraft})
 | (LocalActionBinding&{readonly operation:'update';readonly expectedRevision:number;readonly actionId:string;readonly draft:LocalActionDraft})
 | (LocalActionBinding&{readonly operation:'remove';readonly expectedRevision:number;readonly actionId:string});
export interface LocalActionStore {
 list():Promise<LocalActionList>;
 create(draft:LocalActionDraft,expectedRegistryRevision:number):Promise<LocalActionList>;
 update(actionId:string,draft:LocalActionDraft,expectedRegistryRevision:number):Promise<LocalActionList>;
 remove(actionId:string,expectedRegistryRevision:number):Promise<LocalActionList>;
}
/** Own data-only envelope, no paths/actors/grants/Run. State is base, UUIDv4 or
 * SHA256; registry revision is a nonnegative safe integer. The returned request
 * and every nested draft object are copied and frozen. */
export function localActionRequest(body:unknown):LocalActionRequest;
/** Store selection and original owner authorization must happen in main first.
 * This adapter invokes only the selected private store's CRUD/list method. */
export function invokeLocalActionRequest(store:LocalActionStore,request:unknown):Promise<LocalActionList>;
