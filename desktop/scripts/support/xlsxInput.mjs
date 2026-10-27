import assert from 'node:assert/strict'
import {unzipSync,strFromU8} from 'fflate'
import {DOMParser} from '@xmldom/xmldom'
export function verifyInputWorkbook(bytes,formulaValue){
 const namespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
 const parts=unzipSync(bytes),parse=key=>new DOMParser({errorHandler:{warning:()=>{},error:message=>{throw Error(message)},fatalError:message=>{throw Error(message)}}}).parseFromString(strFromU8(parts[key]),'application/xml')
 const nodes=(root,tag)=>Array.from(root.getElementsByTagNameNS(namespace,tag))
 const sheet=parse('xl/worksheets/sheet1.xml'),cells=nodes(sheet,'c')
 const strings=parts['xl/sharedStrings.xml']?nodes(parse('xl/sharedStrings.xml'),'si').map(node=>node.textContent):[]
 const cell=ref=>{const selected=cells.filter(node=>node.getAttribute('r')===ref);assert.equal(selected.length,1,'Unique saved input cell '+ref);return selected[0]}
 const text=ref=>{const node=cell(ref),value=nodes(node,'v')[0]?.textContent;return node.getAttribute('t')==='s'?strings[Number(value)]:node.getAttribute('t')==='inlineStr'?nodes(node,'is')[0]?.textContent:value}
 assert.equal(text('E1'),'中文','Saved archive retains actual IME Chinese')
 assert.equal(text('E2'),'ok','Saved archive retains restored ABC text')
 assert.equal(Number(text('C1')),formulaValue)
 assert.equal(nodes(cell('C1'),'f')[0]?.textContent,'A1+B1')
 return ['Independent saved ZIP/XML retains native Chinese/ABC input and the original formula']
}
