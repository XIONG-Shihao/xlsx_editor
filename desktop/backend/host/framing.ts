import { Transform } from 'node:stream'
const MAX_FRAME = 16 * 1024 * 1024
const RAW_TAG = 0x80000000, RAW_CHUNK_MAX = 1024 * 1024, RAW_META_MAX = 4096
export function frame(value: unknown): Buffer {
  const data = Buffer.from(JSON.stringify(value, (_key, item) => item instanceof Uint8Array ? { $platformBytes: Buffer.from(item).toString('base64') } : item))
  if (data.length > MAX_FRAME) throw Error('invalid_request: Transport frame exceeds limit')
  const header = Buffer.alloc(4); header.writeUInt32LE(data.length)
  return Buffer.concat([header, data])
}
export function decodeBytes(_key: string, value: unknown): unknown {
  if (value && typeof value === 'object' && Object.keys(value).length === 1 && '$platformUndefined' in value && value.$platformUndefined === 1) return undefined
  if (value && typeof value === 'object' && Object.keys(value).length === 1 && '$platformBytes' in value && typeof value.$platformBytes === 'string') return new Uint8Array(Buffer.from(value.$platformBytes, 'base64'))
  return value
}
/** Accumulate slices once per frame. Fragmented large transfers must not
 * repeatedly copy the whole partially received payload. */
export class Frames extends Transform {
  private header=Buffer.alloc(4)
  private headerBytes=0
  private length=0
  private raw=false
  private slices:Buffer[]=[]
  private received=0
  constructor() { super({ readableObjectMode: true }) }
  _transform(chunk: Buffer, _encoding: BufferEncoding, done: (error?: Error) => void): void {
    try {
      let offset=0
      while(offset<chunk.length){
        if(!this.length){
          const count=Math.min(4-this.headerBytes,chunk.length-offset)
          chunk.copy(this.header,this.headerBytes,offset,offset+count);this.headerBytes+=count;offset+=count
          if(this.headerBytes<4)break
          const tagged=this.header.readUInt32LE();this.raw=Boolean(tagged&RAW_TAG);this.length=tagged&0x7fffffff;this.headerBytes=0
          if(!this.length||this.length>(this.raw?4+RAW_META_MAX+RAW_CHUNK_MAX:MAX_FRAME))throw Error('invalid_request: Invalid transport frame length')
        }
        const count=Math.min(this.length-this.received,chunk.length-offset)
        this.slices.push(chunk.subarray(offset,offset+count));this.received+=count;offset+=count
        if(this.received<this.length)continue
        const body=this.slices.length===1?this.slices[0]:Buffer.concat(this.slices,this.length)
        const message=this.raw?decodeRawUpload(body):JSON.parse(body.toString(),decodeBytes)
        // Object-mode streams reserve null for EOF. Invalid envelopes fail.
        if(!message||typeof message!=='object'||Array.isArray(message)||message instanceof Uint8Array)throw Error('invalid_request: Invalid desktop envelope')
        this.length=0;this.received=0;this.slices=[]
        this.push(message)
      }
      done()
    }catch(error){done(error as Error)}
  }
  _flush(done: (error?: Error) => void): void {done(this.length||this.headerBytes?Error('Incomplete transport frame'):undefined)}
}
function decodeRawUpload(body:Buffer):unknown{
  if(body.length<5)throw Error('invalid_request: Invalid binary document frame')
  const metadataBytes=body.readUInt32LE()
  if(!metadataBytes||metadataBytes>RAW_META_MAX||metadataBytes+4>=body.length)throw Error('invalid_request: Invalid binary document metadata')
  const message=JSON.parse(body.subarray(4,4+metadataBytes).toString()),request=message.args?.[0],bytes=body.subarray(4+metadataBytes)
  if(message.kind!=='invoke'||typeof message.id!=='string'||message.channel!=='document:upload'||!Array.isArray(message.args)||message.args.length!==1||request?.action!=='chunk'||typeof request.token!=='string'||!/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(request.token)||!Number.isSafeInteger(request.offset)||request.offset<0||request.offset+bytes.length>250*1024*1024||!request.bytes||Object.keys(request.bytes).length!==1||request.bytes.$platformRaw!==1||bytes.length>RAW_CHUNK_MAX)throw Error('invalid_request: Invalid binary document upload')
  // View only: ownership stays with this complete frame while the async upload
  // writes the bytes. No Base64 string or numeric-array expansion is created.
  request.bytes=new Uint8Array(bytes.buffer,bytes.byteOffset,bytes.byteLength)
  return message
}
