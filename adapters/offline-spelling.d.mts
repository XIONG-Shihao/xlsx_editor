export type OfflineSpellingLocale='en-US'|'en-GB';
export interface OfflineSpellingEngineIdentity {readonly package:'nspell';readonly version:'2.1.5';readonly codeSha256:string;readonly license:'MIT';readonly licenseSha256:string;readonly dependency:{readonly package:'is-buffer';readonly version:'2.0.5';readonly codeSha256:string;readonly license:'MIT';readonly licenseSha256:string}}
export interface OfflineSpellingDictionaryIdentity {readonly locale:OfflineSpellingLocale;readonly package:'dictionary-en'|'dictionary-en-gb';readonly version:'4.0.0'|'3.0.0';readonly affSha256:string;readonly affBytes:number;readonly dicSha256:string;readonly dicBytes:number;readonly license:'(MIT AND BSD)';readonly licenseSha256:string;readonly licenseBytes:number;readonly dictionaryId:string}
export interface OfflineSpellingAvailability {schema:1;provider:'nspell';providerVersion:'2.1.5';providerId:string;expectedProvider:OfflineSpellingEngineIdentity;verifiedProvider?:OfflineSpellingEngineIdentity;locales:Array<{locale:OfflineSpellingLocale;status:'available'|'unavailable';expected:OfflineSpellingDictionaryIdentity;verified?:OfflineSpellingDictionaryIdentity;reason?:string}>}
export interface OfflineSpellingRequest {requestId:string;locale:string;words:readonly string[]}
export interface OfflineSpellingWord {word:string;correct:boolean;suggestions:string[]}
interface OfflineSpellingReplyBase {schema:1;requestId:string;locale:string;provider:'nspell';providerVersion:'2.1.5';providerId:string}
export type OfflineSpellingReply=(OfflineSpellingReplyBase&{status:'completed';dictionaryId:string;dictionary:OfflineSpellingDictionaryIdentity;results:OfflineSpellingWord[]})|(OfflineSpellingReplyBase&{status:'unavailable'|'cancelled';results:[];reason:string});
export interface OfflineSpellingProvider {availability():Promise<OfflineSpellingAvailability>;check(request:OfflineSpellingRequest):Promise<OfflineSpellingReply>;cancel(requestId:string):Promise<void>;dispose():Promise<void>}
export function createOfflineSpellingProvider():OfflineSpellingProvider;
export const OFFLINE_SPELLING_ENGINE:OfflineSpellingEngineIdentity;
export const OFFLINE_SPELLING_DICTIONARIES:readonly OfflineSpellingDictionaryIdentity[];
export const OFFLINE_SPELLING_PROVIDER_ID:string;
export const OFFLINE_SPELLING_LIMITS:Readonly<{words:100;wordCodePoints:256;requestBytes:32768;suggestions:10;replyBytes:1048576;queuedRequests:4;requestIds:4096;startupMs:15000;workMs:5000;workerOldMb:128;workerYoungMb:16}>;
