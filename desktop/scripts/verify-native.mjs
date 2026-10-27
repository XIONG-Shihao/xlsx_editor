import {assertGuiAllowed} from './tooling/gui-policy.mjs'
import {productRoot} from './tooling/stage.mjs'
import {spawn,execFileSync} from 'node:child_process'
import {mkdtemp,mkdir,cp,writeFile,readFile,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {resolve,join,basename} from 'node:path'
import assert from 'node:assert/strict'
import {workbook} from '../../fixtures/standalone.mjs'
import {unzipSync,zipSync,strFromU8,strToU8} from 'fflate'
import {verifyInputWorkbook} from './support/xlsxInput.mjs'
assertGuiAllowed()
const root=join(productRoot,'desktop'),journey=process.argv[2]??'basic'
if(!['basic','decisions','cold','print','conflict','ime'].includes(journey))throw Error('Unknown standalone native journey')
execFileSync('cargo',['build','--locked','--offline','--features','custom-protocol','--manifest-path',join(root,'src-tauri/Cargo.toml')],{stdio:'inherit'})
const temporary=await mkdtemp(join(tmpdir(),'tauri-xlsx-native-')),bundle=join(temporary,'XLSX Migration Verification.app'),report=join(temporary,'report.json'),exit=join(temporary,'exit.json'),log=join(temporary,'native.log')
await mkdir(join(bundle,'Contents/MacOS'),{recursive:true});await mkdir(join(bundle,'Contents/Resources'),{recursive:true})
await cp(join(root,'src-tauri/target/debug/xlsx-desktop'),join(bundle,'Contents/MacOS/xlsx-desktop'))
await writeFile(join(bundle,'Contents/Info.plist'),'<?xml version="1.0"?><plist version="1.0"><dict><key>CFBundleIdentifier</key><string>com.agentplatform.xlsx.tauri.verification</string><key>CFBundleName</key><string>XLSX Migration Verification</string><key>CFBundleExecutable</key><string>xlsx-desktop</string><key>CFBundlePackageType</key><string>APPL</string><key>CFBundleVersion</key><string>1</string><key>NSHighResolutionCapable</key><true/></dict></plist>')
execFileSync('codesign',['--force','--sign','-',bundle],{stdio:'pipe'})
const pdfOutputs=journey==='print'?[join(temporary,'standalone-a4.pdf'),join(temporary,'standalone-after-cancel.pdf')]:[]
const conflictSource=join(temporary,'external-change.xlsx'),conflictCopy=join(temporary,'accepted-copy.xlsx')
const externalParts=unzipSync(workbook());externalParts['xl/worksheets/sheet1.xml']=strToU8(strFromU8(externalParts['xl/worksheets/sheet1.xml']).replace('<v>10</v>','<v>88</v>').replace('<v>30</v>','<v>108</v>'))
const externalBytes=Buffer.from(zipSync(externalParts))
if(journey==='conflict')await writeFile(conflictSource,workbook())
const imeOutput=join(temporary,'native-ime.xlsx')
const injection=join(temporary,'journey.js');await writeFile(injection,'window.__XLSX_JOURNEY__='+JSON.stringify(journey)+';window.__XLSX_PDF_OUTPUTS__='+JSON.stringify(pdfOutputs)+';window.__XLSX_CONFLICT_COPY__='+JSON.stringify(conflictCopy)+';window.__XLSX_IME_OUTPUT__='+JSON.stringify(imeOutput)+';\n'+(journey==='ime'?await readFile(join(root,'scripts/native-xlsx-input.js'),'utf8')+'\n':'')+await readFile(join(root,'scripts/verify-native.js'),'utf8'))
const environment={XLSX_EDITOR_GUI_TESTS:'1',XLSX_EDITOR_USER_DATA:join(temporary,'profile'),XLSX_EDITOR_GUI_SCRIPT:injection,XLSX_EDITOR_GUI_REPORT:report,XLSX_EDITOR_GUI_EXIT_REPORT:exit}
let child,launcherExit
function launch(){child=spawn('/usr/bin/open',['-n','-W','--stdout',log,'--stderr',log,...Object.entries(environment).flatMap(([key,value])=>['--env',key+'='+value]),bundle,...(journey==='conflict'?['--args',conflictSource]:[])],{stdio:['ignore','pipe','pipe']});launcherExit=new Promise(resolve=>child.once('exit',(code,signal)=>resolve({code,signal})))}
launch()
console.log('Standalone native fixture: '+bundle)
const descendants=new Map()
function owned(){
 const rows=execFileSync('ps',['-axo','pid,ppid,lstart,command'],{encoding:'utf8'}).split('\n').flatMap(line=>{const match=line.match(/^\s*(\d+)\s+(\d+)\s+(.+)$/);return match?[{pid:Number(match[1]),parent:Number(match[2]),identity:match[3]}]:[]})
 for(const row of rows)if(row.identity.includes(basename(temporary)+'/XLSX Migration Verification.app/Contents/MacOS/xlsx-desktop'))descendants.set(row.pid,row.identity)
 const present=new Set([child.pid,...rows.filter(row=>descendants.get(row.pid)===row.identity).map(row=>row.pid)])
 let changed=true;while(changed){changed=false;for(const row of rows)if(present.has(row.parent)&&!present.has(row.pid)){present.add(row.pid);descendants.set(row.pid,row.identity);changed=true}}
 return rows.filter(row=>descendants.get(row.pid)===row.identity)
}
let result,externalChanged=false;const phases=[]
try{
 for(let phase=0;phase<(journey==='cold'?2:1);phase++){
 result=undefined
 const start=performance.now()
 let pendingPhase
 while(performance.now()-start<(['print','ime'].includes(journey)?240000:120000)){
  owned()
  try{result=JSON.parse(await readFile(report,'utf8'))}catch{}
  if(journey==='conflict'&&result?.pending&&result.phase==='change-native-source'&&!externalChanged){await writeFile(conflictSource,externalBytes);externalChanged=true}
  if(result?.pending&&result.phase!==pendingPhase){pendingPhase=result.phase;console.log(JSON.stringify(result))}
  if(result&&!result.pending)break
  if(child.exitCode!==null||child.signalCode!==null)throw Error('Native shell exited without report: '+await readFile(log,'utf8'))
  await new Promise(resolve=>setTimeout(resolve,250))
 }
 assert.ok(result,'No standalone native report')
 await Promise.race([launcherExit,new Promise((_,reject)=>setTimeout(()=>reject(Error('Standalone native Quit did not finish')),15000))])
 const stop=performance.now();while(owned().length&&performance.now()-stop<10000)await new Promise(resolve=>setTimeout(resolve,100))
 assert.equal(owned().length,0,'Standalone app and owned helpers exited')
 assert.equal(JSON.parse(await readFile(exit,'utf8')).approvedQuit,true,'Native close barrier approved Quit')
 assert.equal(child.exitCode,0)
 let holders;try{holders=execFileSync('/usr/sbin/lsof',['-t','--',join(temporary,'profile/.tauri-profile.lock')],{encoding:'utf8',stdio:['ignore','pipe','pipe']}).trim()}catch(error){if(error.status!==1)throw error;holders=''}
 assert.equal(holders,'','No profile writer remains after native Quit')
 assert.equal(result.passed,true,JSON.stringify(result))
 if(journey==='ime')result.checks.push(...verifyInputWorkbook(new Uint8Array(await readFile(imeOutput)),35))
 if(journey==='conflict'){
  assert.equal(externalChanged,true,'Driver actually changed the bound source')
  assert.deepEqual(await readFile(conflictSource),externalBytes,'Failed Save and accepted Copy never overwrite the externally changed source')
  execFileSync('python3',['-c','import sys,zipfile,xml.etree.ElementTree as E\nz=zipfile.ZipFile(sys.argv[1]); assert z.testzip() is None\nn={"x":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}; r=E.fromstring(z.read("xl/worksheets/sheet1.xml")); c={c.attrib["r"]:c.find("x:v",n).text for c in r.findall(".//x:c",n) if c.find("x:v",n) is not None}\nassert float(c["A1"])==42 and float(c["C1"])==62; assert r.find(".//x:c[@r=\'C1\']/x:f",n).text=="A1+B1"; assert z.read("custom/preserved.bin")==bytes([0,1,2,255])\nprint("Independent native conflict-copy archive verified")',conflictCopy],{stdio:'inherit'})
  result.checks.push('Independent external source bytes remain exact; conflict copy retains accepted cells, formula and unknown part after Quit')
 }
 if(journey==='print'){
  const {DOMMatrix,ImageData,Path2D}=await import('@napi-rs/canvas');Object.assign(globalThis,{DOMMatrix,ImageData,Path2D})
  const {getDocument}=await import('pdfjs-dist/legacy/build/pdf.mjs')
  assert.equal(result.pageCount,1,'Accepted standalone preview is one page')
  for(const path of pdfOutputs){
   const bytes=new Uint8Array(await readFile(path)),pdf=await getDocument({data:bytes,standardFontDataUrl:join(productRoot,'node_modules/pdfjs-dist/standard_fonts')+'/',disableFontFace:true}).promise
   try{
    assert.equal(pdf.numPages,result.pageCount);const page=await pdf.getPage(1),view=page.getViewport({scale:1})
    assert.ok(Math.abs(view.width-595.28)<0.5&&Math.abs(view.height-841.89)<0.5,'Native standalone PDF retains A4 dimensions')
    const text=(await page.getTextContent()).items.map(item=>item.str??'').join(' ').normalize('NFKC')
    for(const value of [15,20,35,3,4])assert.match(text,new RegExp('\\b'+value+'\\b'),'Independent PDF retains cell/formula output '+value)
    assert.match(text,/alpha/);assert.ok(text.replace(/\s/g,'').includes('中文'),'Independent PDF retains Chinese clipboard text')
   }finally{await pdf.destroy()}
  }
  result.checks.push('Both actual standalone PDF outputs independently retain A4 dimensions, formula and Chinese clipboard content')
 }
 phases.push(result)
 if(journey==='cold'&&phase===0){
  assert.equal(result.restartRequired,true)
  await rm(report);await rm(exit)
  await writeFile(injection,'window.__XLSX_JOURNEY__="cold";window.__XLSX_COLD__=true;\n'+await readFile(join(root,'scripts/verify-native.js'),'utf8'))
  launch()
 }
 }
 await rm(temporary,{recursive:true,force:true})
 console.log(JSON.stringify({...result,phases}))
 console.log('Standalone native acceptance complete; profile removed')
}catch(error){
 console.error('Kept for diagnosis: '+temporary+'; still running: '+JSON.stringify(owned()))
 throw error
}
