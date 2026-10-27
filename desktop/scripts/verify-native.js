(async()=>{
 const native=(command,args)=>window.__TAURI_INTERNALS__.invoke(command,args),checks=[],samples=[],sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms))
 const assert=(value,message)=>{if(!value)throw Error(message);checks.push(message)}
 async function wait(predicate,name,timeout=30000){const start=performance.now();while(!await predicate()){if(performance.now()-start>timeout)throw Error('Timed out: '+name);await sleep(50)}}
 const cell=ref=>document.querySelector(`.desktop-workbook:not([hidden]) input[aria-label="Sheet1 ${ref}"]`)
 const click=text=>{const button=[...document.querySelectorAll('.desktop-header button')].find(node=>node.textContent.trim()===text);assert(button&&!button.disabled,'Enabled standalone action '+text);button.click()}
 async function edit(ref,text){const input=cell(ref);input.focus();Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(input,text);input.dispatchEvent(new InputEvent('input',{bubbles:true,data:text,inputType:'insertText'}));await wait(()=>cell(ref)?.dataset.editing==='true','cell draft '+ref);cell(ref).dispatchEvent(new KeyboardEvent('keydown',{bubbles:true,key:'Enter'}));await wait(()=>cell(ref)?.dataset.editing==='false'&&document.querySelector('.desktop-workbook:not([hidden]) [data-home-command="home.undo"]')&&!document.querySelector('.desktop-workbook:not([hidden]) [data-home-command="home.undo"]').disabled,'accepted edit '+ref);await sleep(100)}
 async function sample(label){await sleep(1000);for(let index=0;index<3;index++){samples.push({label,index,...await native('standalone_memory')});if(index<2)await sleep(500)}}
 try{
  await wait(()=>window.xlsxDesktop&&(window.__XLSX_JOURNEY__==='conflict'?cell('A1')&&!cell('A1').readOnly:document.querySelector('.desktop-empty')),'standalone ready view')
  assert(/AppleWebKit/.test(navigator.userAgent),'Real standalone WKWebView renderer')
  assert((await native('standalone_status')).windowCount===1,'Independent standalone product owns one native window')
  assert(!window.api,'Standalone does not expose the full-app bridge')
  if(window.__XLSX_JOURNEY__==='conflict'){
   await wait(()=>cell('A1')?.value==='10'&&cell('C1')?.value==='30','launch-file source')
   const initial=await window.xlsxDesktop.current(),recent=(await window.xlsxDesktop.recentFiles()).entries[0]
   await edit('A1','42');await wait(()=>cell('C1')?.value==='62','accepted formula before conflict')
   const duplicate=await window.xlsxDesktop.openRecent(recent.id),active=await window.xlsxDesktop.activateOpen(duplicate.id)
   assert(active.id===initial.id&&active.documentId===initial.documentId&&(await window.xlsxDesktop.documents()).documents.length===1,'Reopening the same source retains one accepted document owner')
   await native('verification_report',{report:{pending:true,phase:'change-native-source',checks}})
   // The driver's next bounded poll replaces only this disposable source.
   // Save succeeding before replacement is a failure, never a passing conflict.
   await sleep(1500);click('Save')
   await wait(()=>document.querySelector('.desktop-conflict')&&cell('A1')?.value==='42'&&cell('C1')?.value==='62','visible native source conflict')
   assert((await window.xlsxDesktop.current()).conflict===true,'Failed Save exposes the retained original-file conflict')
   assert((await window.xlsxDesktop.current()).dirty===true,'Failed Save keeps accepted edits dirty and available')
   await native('verification_report',{report:{pending:true,phase:'save-native-conflict-copy',destination:window.__XLSX_CONFLICT_COPY__,checks}});click('Save a copy…')
   await wait(()=>document.querySelector('.desktop-notice')?.textContent.includes('Copy saved.'),'native conflict copy',90000)
   const copied=await window.xlsxDesktop.current()
   assert(copied.id===initial.id&&copied.documentId===initial.documentId&&copied.conflict&&copied.dirty,'Save a copy preserves the conflicted original binding and accepted owner')
   await window.xlsxDesktop.request(initial.id,{operation:'discard'});await window.xlsxDesktop.closeTab(initial.id)
   assert((await window.xlsxDesktop.recoveryFiles()).entries.length===0,'Explicit fixture discard removes its recovery without publishing over the source')
   assert((await native('standalone_status')).windowCount===1,'Conflict and native copy use one visible standalone window')
   await native('verification_report',{report:{passed:true,checks,samples,userAgent:navigator.userAgent}});await native('standalone_quit');return
  }
  if(window.__XLSX_COLD__){
   await wait(()=>document.querySelector('.desktop-drafts button.desktop-recent-open')&&!document.querySelector('.desktop-drafts button.desktop-recent-open').disabled,'cold recovery entry')
   document.querySelector('.desktop-drafts button.desktop-recent-open').click()
   await wait(()=>cell('C1')?.value==='35'&&cell('A1')?.value==='15'&&cell('B1')?.value==='20','cold recovered values')
   assert(!!document.querySelector('.desktop-recovery'),'Original standalone recovery UI is visible after cold launch')
   await wait(()=>!document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Undo]').disabled,'cold Undo ready')
   document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Undo]').click();await wait(()=>cell('C1')?.value==='','cold recovery Undo')
   await wait(()=>!document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Redo]').disabled,'cold Redo ready')
   document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Redo]').click();await wait(()=>cell('C1')?.value==='35','cold recovery Redo')
   assert(true,'Cold process recovery retains original values, formulas and undo/redo')
   await sample('standalone-cold-recovered')
   const recovered=await window.xlsxDesktop.current();await window.xlsxDesktop.request(recovered.id,{operation:'discard'})
   await native('verification_report',{report:{passed:true,checks,samples,userAgent:navigator.userAgent}});await native('standalone_quit');return
  }
  await sample('standalone-empty-home')
  click('New workbook');await wait(()=>cell('A1')&&!cell('A1').readOnly&&!document.querySelector('.desktop-header button').disabled,'new workbook')
  await edit('A1','15');await edit('B1','20');await edit('C1','=A1+B1');await wait(()=>cell('A1')?.value==='15'&&cell('B1')?.value==='20'&&cell('C1')?.value==='35','native standalone calculation')
  assert(true,'Original standalone grid edits and calculates through the native owner')
  const undo=document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Undo]');undo.click();await wait(()=>cell('C1')?.value==='','standalone Undo')
  await wait(()=>!document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Redo]').disabled,'standalone Redo ready')
  document.querySelector('.desktop-workbook:not([hidden]) button[aria-label=Redo]').click();await wait(()=>cell('C1')?.value==='35','standalone Redo')
  assert(true,'Original standalone undo/redo retains the formula result')
  await sample('standalone-loaded-dirty')
  if(window.__XLSX_JOURNEY__==='ime'){
   const input=await window.__PLATFORM_VERIFY_XLSX_INPUT__({cell,history:label=>document.querySelector('.desktop-workbook:not([hidden]) button[aria-label="'+label+'"]'),report:(phase,extra)=>native('verification_report',{report:{pending:true,phase,checks,...extra}})})
   checks.push(...input.checks)
   await native('verification_report',{report:{pending:true,phase:'save-native-ime-copy',destination:window.__XLSX_IME_OUTPUT__,checks,trace:input.trace}});click('Save a copy…')
   await wait(()=>document.querySelector('.desktop-notice')?.textContent.includes('Copy saved.'),'native IME archive copy',90000)
   const current=await window.xlsxDesktop.current();await window.xlsxDesktop.request(current.id,{operation:'discard'})
   await native('verification_report',{report:{passed:true,checks,trace:input.trace,samples,userAgent:navigator.userAgent}});await native('standalone_quit');return
  }
  if(window.__XLSX_JOURNEY__==='print'){
   const workbook=()=>document.querySelector('.desktop-workbook:not([hidden])'),panel=()=>workbook()?.querySelector('section[aria-label="Print and PDF"]'),action=text=>[...panel().querySelectorAll('button')].find(button=>button.textContent.trim()===text)
   const pending=(phase,extra={})=>native('verification_report',{report:{pending:true,phase,checks,...extra}})
   const clipboard='alpha\t中文\n3\t4';await window.xlsxDesktop.copyText(clipboard);assert(await window.xlsxDesktop.readClipboardText()===clipboard,'Standalone native clipboard preserves Unicode TSV')
   cell('E1').focus();await pending('paste-native-clipboard',{instruction:'Paste the synthetic TSV into focused E1; another app may have changed the system clipboard',clipboard})
   await wait(()=>cell('E1')?.value==='alpha'&&cell('F1')?.value==='中文'&&cell('E2')?.value==='3'&&cell('F2')?.value==='4','physical native clipboard paste',90000)
   assert(true,'Native Command-V pastes all four TSV cells through the original editor')
   const toggle=workbook().querySelector('button[aria-label="Print and PDF"]');assert(toggle,'Original standalone print controls are present');toggle.click()
   await wait(()=>panel()&&action('Preview pages')&&!action('Preview pages').disabled,'standalone print settings')
   const area=panel().querySelector('[data-print-facet="area"]');Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(area,'A1:F2');area.dispatchEvent(new Event('input',{bubbles:true}));await sleep(100)
   action('Preview pages').click();await wait(()=>action('Export PDF')&&!action('Export PDF').disabled,'accepted standalone print preview')
   const pageCount=Number(panel().textContent.match(/Page\s+1\s*\/\s*(\d+)/)?.[1]);assert(pageCount===1,'Standalone accepted A4 preview contains one page')
   assert(!!panel().querySelector('iframe')?.contentDocument?.querySelector('[data-xlsx-print-page]'),'Original standalone preview renders the accepted page')
   await pending('cancel-native-pdf');action('Export PDF').click();await wait(()=>/cancelled/i.test(panel().textContent)&&!action('Export PDF').disabled,'native PDF chooser Cancel',90000)
   assert(true,'Native PDF chooser Cancel keeps the accepted preview available')
   async function exportPdf(index){
    await pending('save-native-pdf-'+index,{destination:window.__XLSX_PDF_OUTPUTS__[index]});action('Export PDF').click()
    await wait(()=>panel().textContent.includes('PDF output finished.')&&!action('Export PDF').disabled,'standalone PDF saved',90000)
    assert((await native('standalone_status')).windowCount===1,'PDF rendering retains one visible standalone window')
   }
   await exportPdf(0)
   await pending('cancel-native-printer');action('Print').click();await wait(()=>/cancelled/i.test(panel().textContent)&&!action('Print').disabled,'native printer Cancel',90000)
   assert(true,'Native printer cancellation releases the print lease')
   await exportPdf(1)
   const current=await window.xlsxDesktop.current();await window.xlsxDesktop.request(current.id,{operation:'discard'})
   await native('verification_report',{report:{passed:true,checks,samples,pageCount,userAgent:navigator.userAgent}});await native('standalone_quit');return
  }
  click('New workbook');await wait(()=>document.querySelectorAll('.desktop-tabs [role=tab]').length===2&&cell('A1')?.value===''&&![...document.querySelectorAll('.desktop-tabs [role=tab]')].some(tab=>tab.disabled),'second workbook tab')
  const tabs=()=>[...document.querySelectorAll('.desktop-tabs [role=tab]')]
  tabs()[0].click();await wait(()=>cell('C1')?.value==='35','first independent tab');assert(true,'Independent workbook tabs preserve edited values')
  const current=await window.xlsxDesktop.current()
  if(window.__XLSX_JOURNEY__==='decisions'){
   await native('verification_report',{report:{pending:true,stage:'cancel-close',checks}})
   const cancelled=await window.xlsxDesktop.closeTab(current.id)
   assert(cancelled.cancelled&&!cancelled.closed,'Native Cancel preserves the dirty workbook')
   assert((await window.xlsxDesktop.documents()).documents.length===2,'Cancel preserves both tabs')
   await native('verification_report',{report:{pending:true,stage:'cancel-native-quit',checks}})
   await native('standalone_quit')
   await wait(async()=>{const status=await native('standalone_status');return status.pendingQuit&&status.dialog},'native Quit decision')
   await wait(async()=>!(await native('standalone_status')).pendingQuit,'cancelled native Quit resets the barrier')
   assert((await window.xlsxDesktop.documents()).documents.length===2,'Cancelled native Quit preserves both workbooks and permits another close')
   await native('verification_report',{report:{pending:true,stage:'keep-recovery',checks}})
   const kept=await window.xlsxDesktop.closeTab(current.id)
   assert(kept.closed&&!kept.cancelled,'Native Keep recovery closes only the selected workbook')
   assert((await window.xlsxDesktop.recoveryFiles()).entries.length===1,'Accepted recovery remains available after native Keep')
  }else if(window.__XLSX_JOURNEY__==='cold'){
   await native('verification_report',{report:{pending:true,stage:'keep-for-cold-restart',checks}})
   const kept=await window.xlsxDesktop.closeTab(current.id)
   assert(kept.closed&&!kept.cancelled,'Native Keep recovery prepares a cold restart')
  }else await window.xlsxDesktop.request(current.id,{operation:'discard'})
  // This first host fixture deliberately discards its synthetic draft. Save,
  // chooser decisions and cold recovery have their own native journey gates.
  await native('verification_report',{report:{passed:true,checks,samples,restartRequired:window.__XLSX_JOURNEY__==='cold',userAgent:navigator.userAgent}})
  await native('standalone_quit')
 }catch(error){await native('verification_report',{report:{passed:false,checks,samples,error:String(error),stack:error.stack,body:document.body.innerText.slice(0,3500)}});await native('standalone_quit')}
})()
