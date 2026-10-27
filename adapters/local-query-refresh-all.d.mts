import type {LocalQueryConnections} from './local-query-connections.mjs';
import type {LocalQueryRefreshAllScope,LocalQueryRefreshAllPreview,LocalQueryRefreshAllResult} from '../core/local-query.mjs';
import type {DelimitedProgress} from '../core/delimited-import.mjs';
export interface RefreshAllContext {
 current:()=>LocalQueryRefreshAllScope;
 calendar?:(sheet:string)=>boolean|undefined;
 request:(operation:string,input:any,fenced?:boolean)=>Promise<any>;
 apply:(parent:string,descriptor:any)=>Promise<{state:string}>;
}
export interface RefreshAllOptions {signal?:AbortSignal;onProgress?:(event:DelimitedProgress)=>void;authorizeApply?:()=>Promise<void>;beforeApply?:()=>void;afterApply?:(accepted:{parentState:string;state:string})=>void|Promise<void>}
export const REFRESH_ALL_LIMITS:Readonly<{plans:4;connections:32;lifetimeMs:number;responseBytes:number}>;
export class LocalQueryRefreshAll {
 constructor(args:{store:LocalQueryConnections;document:string;directory:string;context:RefreshAllContext;now?:()=>number});
 preview(options?:RefreshAllOptions):Promise<{refreshAllPreview:LocalQueryRefreshAllPreview}>;
 refresh(id:string,options?:RefreshAllOptions):Promise<{refreshAllResult:LocalQueryRefreshAllResult}>;
 dispose():void;
}
