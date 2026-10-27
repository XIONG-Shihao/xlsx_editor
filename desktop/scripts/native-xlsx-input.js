// Injected only by the existing debug verification runners. Observe genuine
// OS composition events; never synthesize input or supply a model/UI result.
window.__PLATFORM_VERIFY_XLSX_INPUT__=async({cell,history,report})=>{
 const checks=[],trace=[],sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms))
 const assert=(value,message)=>{if(!value)throw Error(message);checks.push(message)}
 async function wait(predicate,name,timeout=90000){const start=performance.now();while(!predicate()){if(performance.now()-start>timeout)throw Error('Native input deadline: '+name);await sleep(50)}}
 const types=['compositionstart','compositionupdate','compositionend','input','keydown']
 const observe=event=>{
  const label=event.target?.getAttribute?.('aria-label')
  if(!['Sheet1 E1','Sheet1 E2'].includes(label))return
  trace.push({type:event.type,cell:label.slice(7),data:event.data??null,key:event.key??null,isComposing:!!event.isComposing,trusted:event.isTrusted})
 }
 for(const type of types)document.addEventListener(type,observe,true)
 try{
  cell('E1').focus()
  await report('type-native-chinese-ime',{instruction:'Switch ABC to Simplified Chinese Pinyin, type zhongwen, select 中文 with Space, then Enter to commit.'})
  await wait(()=>cell('E1')?.value==='中文'&&cell('E1')?.dataset.editing==='false','physical Pinyin Chinese commit')
  assert(trace.some(event=>event.cell==='E1'&&event.type==='compositionstart'&&event.trusted)&&trace.some(event=>event.cell==='E1'&&event.type==='compositionend'&&event.trusted),'Actual trusted OS composition starts and finishes in the XLSX cell')
  await wait(()=>history('Undo')&&!history('Undo').disabled,'IME Undo ready');history('Undo').click()
  await wait(()=>cell('E1')?.value==='','one Undo removes whole Chinese commit')
  await wait(()=>history('Redo')&&!history('Redo').disabled,'IME Redo ready');history('Redo').click()
  await wait(()=>cell('E1')?.value==='中文','Redo restores whole Chinese commit')
  assert(true,'One Undo/Redo restores the complete native Chinese composition')
  cell('E2').focus()
  await report('type-native-ascii-after-ime',{instruction:'Switch back to ABC, type ok, then Enter to commit.'})
  await wait(()=>cell('E2')?.value==='ok'&&cell('E2')?.dataset.editing==='false','physical ABC commit')
  assert(!trace.some(event=>event.cell==='E2'&&event.type==='compositionstart'),'Restored ABC input commits without a Chinese composition')
  assert(cell('E1').value==='中文','ABC input preserves the accepted Chinese cell')
  return {checks,trace}
 }finally{for(const type of types)document.removeEventListener(type,observe,true)}
}
