/** Observe real journal filesystem boundaries; production code has no test hooks. */
const fs=require('node:fs/promises');
const path=require('node:path');
const {syncBuiltinESMExports}=require('node:module');
const {createHash}=require('node:crypto');
const config=JSON.parse(process.argv[2]);
const original={open:fs.open,rename:fs.rename,readFile:fs.readFile};
const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
let armed=false,hit=false,manifestPublished=false,nodeId,descriptor,before;
const send=message=>new Promise((resolve,reject)=>process.send(message,error=>error?reject(error):resolve()));
async function checkpoint(boundary){
  if(!armed||hit||config.boundary!==boundary)return;hit=true;
  await send({type:'checkpoint',boundary,nodeId,descriptor,before,sourceHash:hash(await original.readFile(config.filename))});
  await new Promise(()=>{});
}
fs.open=async function(filename,...options){
  const handle=await original.open(filename,...options),sync=handle.sync.bind(handle);
  handle.sync=async()=>{
    const result=await sync();if(!armed)return result;
    const node=/^([a-f0-9-]{36})\.json\.[a-f0-9-]{36}\.tmp$/.exec(path.basename(String(filename)));
    if(node){nodeId=node[1];await checkpoint('node-file-sync');}
    if(path.basename(String(filename)).startsWith('current.json.')&&String(filename).endsWith('.tmp'))await checkpoint('manifest-file-sync');
    if(filename===config.ownerDirectory&&manifestPublished)await checkpoint('manifest-directory-sync');
    return result;
  };return handle;
};
fs.rename=async function(source,destination){
  const result=await original.rename(source,destination);if(!armed)return result;
  if(/^[a-f0-9-]{36}\.json$/.test(path.basename(String(destination)))){nodeId=path.basename(destination).slice(0,-5);await checkpoint('node-rename');}
  if(path.basename(String(destination))==='current.json'){manifestPublished=true;await checkpoint('manifest-rename');}
  return result;
};
syncBuiltinESMExports();
(async()=>{
  const {SpreadsheetDocument}=await import('../../adapters/document-session.mjs');
  const {nativeWorkbookOperationAsync,closeNativeWorkbookWorker}=await import('../../core/native-runtime.mjs');
  const owner=new SpreadsheetDocument({recoveryDirectory:config.recoveryDirectory,indexDirectory:config.indexDirectory,native:input=>nativeWorkbookOperationAsync(input)});
  try{
    await owner.open(config.filename);before=owner.status();const bound={state:before.state,documentId:before.documentId};
    const handle=await owner.request({operation:'importBegin',...bound,shape:{sheet:'Sheet1',start:'A1',rows:1001,columns:1},count:1001,policy:'grid'});
    const {inspectPaste,pastePages}=await import('../../ui/largePaste.ts'),text='1\n'.repeat(1001),plan=inspectPaste(text,{sheet:'Sheet1',start:'A1'});
    for(const page of pastePages(text,plan))await owner.request({operation:'importUpload',...bound,uploadId:handle.id,...page});
    descriptor=await owner.request({operation:'importSeal',...bound,uploadId:handle.id});armed=true;
    await owner.request({operation:'apply',state:bound.state,commands:[descriptor]});
    await send({type:'complete',status:owner.status()});
  }finally{armed=false;await owner.close();closeNativeWorkbookWorker();if(process.connected)process.disconnect();}
})().catch(async error=>{if(process.connected)await send({type:'fatal',error:String(error.stack??error)}).catch(()=>{});process.exitCode=1;if(process.connected)process.disconnect();});
