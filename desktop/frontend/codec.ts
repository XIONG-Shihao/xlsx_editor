/** Preserve the original desktop argument/value semantics across JSON IPC. */
export function encode(value:unknown):unknown{
 if(value===undefined)return {$platformUndefined:1}
 if(value instanceof Uint8Array){let text='';for(let offset=0;offset<value.length;offset+=8192)text+=String.fromCharCode(...value.subarray(offset,offset+8192));return {$platformBytes:btoa(text)}}
 if(Array.isArray(value))return value.map(encode)
 if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).map(([key,item])=>[key,encode(item)]))
 return value
}
export function decode(value:unknown):unknown{
 if(value&&typeof value==='object'&&!Array.isArray(value)){
  const object=value as Record<string,unknown>
  if(Object.keys(object).length===1){if(object.$platformUndefined===1)return undefined;if(typeof object.$platformBytes==='string')return Uint8Array.from(atob(object.$platformBytes),character=>character.charCodeAt(0))}
  return Object.fromEntries(Object.entries(object).map(([key,item])=>[key,decode(item)]))
 }
 return Array.isArray(value)?value.map(decode):value
}
